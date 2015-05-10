using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Net.Mail;
using System.Net;

using OpenStack;
using OpenStack.Identity;
using OpenStack.Storage;

using net.openstack.Core.Domain;
using net.openstack.Core.Providers;
using net.openstack.Providers.Rackspace;

using CommandLine;
using CommandLine.Text;

//using CrystalDecisions.CrystalReports.Engine;
//using CrystalDecisions.Shared;


namespace RptDynamo
{
    class Program
    {
        static void Main(string[] args)
        {
            var options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                // consume Options instance properties
                if (options.Verbose)
                {
                    Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
                    Console.WriteLine(options.config);
                    Console.WriteLine(options.job);
                }
                else
                    Console.WriteLine("working ...");

                DataContractJsonSerializer serializer = null;

                // Process Config File
                Trace.WriteLine("Processing Config File");
                serializer = new DataContractJsonSerializer(typeof(RptDynamoConfig));
                RptDynamoConfig config = (RptDynamoConfig)serializer.ReadObject(File.OpenRead(options.config));

                // Process Job File
                Trace.WriteLine("Processing Job File");
                serializer = new DataContractJsonSerializer(typeof(RptJob));
                RptJob job = (RptJob)serializer.ReadObject(File.OpenRead(options.job));

                var tempPath = Environment.GetEnvironmentVariable("TEMP") + "//RptDynamo_" + Path.GetFileNameWithoutExtension(options.job);
                Directory.CreateDirectory(tempPath);
                Environment.SetEnvironmentVariable("TEMP", tempPath);
                Environment.SetEnvironmentVariable("TMP", tempPath);

                RptStatusAPI sAPI = new RptStatusAPI(config.apiUri);
                sAPI.processing(Guid.Parse(Path.GetFileNameWithoutExtension(options.job))).Wait();
                SentRpt(config, ProcessRpt(job, config.swiftCfg, sAPI), sAPI);
                Directory.Delete(tempPath, true);
            }
        }
        static MailMessage ProcessRpt(RptJob rptJob, OSSwift swiftCfg, RptStatusAPI sAPI)
        {
            RptEmail email = new RptEmail();

            // Setup Mail Message
            MailMessage mailMessage = new MailMessage();
            StringBuilder eMailBody = new StringBuilder();

            // Mail Message add 'TO' recpients
            if (rptJob.email.to != null)
            {
                rptJob.email.to.ForEach(delegate(String name)
                {
                    try { mailMessage.To.Add(name); }
                    catch (Exception e) { Trace.WriteLine(e.ToString()); }
                });
            }
            // Mail Message add 'CC' recipents
            if (rptJob.email.cc != null)
            {
                rptJob.email.cc.ForEach(delegate(String name)
                {
                    try { mailMessage.CC.Add(name); }
                    catch (Exception e) { Trace.WriteLine(e.ToString()); }
                });
            }


            // Custom Message Handling
            bool suppressParameters = false; //Controls if parameters should be listed in email
            if (rptJob.email.custom != null)
            {
                // Custom Subject Handling
                if (!String.IsNullOrEmpty(rptJob.email.custom.subject)) { mailMessage.Subject = "[RptDynamo] " + rptJob.email.custom.subject; }

                // Custom Body Handling
                if (!String.IsNullOrEmpty(rptJob.email.custom.body))
                {
                    eMailBody.AppendLine(rptJob.email.custom.body);
                    eMailBody.AppendLine("");
                }
                suppressParameters = rptJob.email.custom.supressparameters;

            }

            if (!suppressParameters)
            {
                foreach (Parameters rptParam in rptJob.report.Parameters)
                {
                    eMailBody.AppendLine(rptParam.Name + " is set to " + string.Join(", ", rptParam.SelectedValues));
                }
            }

            //Export report in seperate AppDomain
            String rptTitle = null;
            String rptExport = null;

            //Isolation for OutOfMemory Exception http://stackoverflow.com/questions/18042225/outofmemoryexception-trap-event-beside-try-catch
            using (Isolated<Facade> isolated = new Isolated<Facade>())
            {
                
                try
                {
                    rptExport = isolated.Value.Export(Newtonsoft.Json.JsonConvert.SerializeObject(rptJob.report));
                }
                catch (OutOfMemoryException e)
                {
                    System.Threading.Thread.Sleep(60000); //Wait for AppDomain Unload (Prevents Main Application Failure)
                    eMailBody.AppendLine("");
                    eMailBody.AppendLine(e.ToString());
                    mailMessage.Bcc.Add("aladuca@jhmc.org, mcallagh@jhmc.org");
                    sAPI.failed().Wait();
                }
                catch (Exception e)
                {
                    eMailBody.AppendLine("");
                    eMailBody.AppendLine(e.ToString());
                    mailMessage.Bcc.Add("aladuca@jhmc.org, mcallagh@jhmc.org");
                    sAPI.failed().Wait();
                }
            }

            string swiftaHRef = UploadRpt(rptExport, swiftCfg, rptJob);
            if (String.IsNullOrEmpty(swiftaHRef) && !String.IsNullOrEmpty(rptExport) && File.Exists(rptExport))
            {
                var fileInfo = new FileInfo(email.file);
                if (fileInfo.Length < 26214400)
                {
                    mailMessage.Attachments.Add(new Attachment(email.file));
                    sAPI.completed().Wait();
                }
                else
                {
                    sAPI.failed().Wait();
                    email.body.AppendLine("<br/><br/><font color=\"red\"><strong>Report export too large to be attached.</strong></font>");
                    File.Copy(email.file, "C:\\Users\\Public\\Documents\\RptPutty-RptDynamo\\Exports\\" + rptJob.JobID + Path.GetExtension(email.file));
                }
            }
            else if (!String.IsNullOrEmpty(swiftaHRef))
            {
                eMailBody.AppendLine(swiftaHRef);
                sAPI.completed().Wait();
            }

            // Subject Handling (Preserves Custom Subject)
            if (String.IsNullOrEmpty(mailMessage.Subject))
            {
                if (!String.IsNullOrEmpty(rptTitle)) { mailMessage.Subject = "[RptDynamo] " + rptTitle; }
                else if (!String.IsNullOrEmpty(rptJob.report.Title)) { mailMessage.Subject = "[RptDynamo] " + rptJob.report.Title; }
                else { mailMessage.Subject = "[RptDynamo] " + Path.GetFileNameWithoutExtension(rptJob.report.Filename); }
            }

            // Body to HTML
            mailMessage.Body = eMailBody.Replace(Environment.NewLine, Environment.NewLine + "<br/>").ToString();
            mailMessage.IsBodyHtml = true;

            return mailMessage;
        }
        static string UploadRpt(string filename, OSSwift swiftCfg, RptJob rptRqst)
        {
            if (!String.IsNullOrEmpty(filename) && File.Exists(filename) && swiftCfg != null)
            {
                Trace.WriteLine("Getting OpenStack Crudential");
                // Using openstack.net @ https://github.com/openstacknetsdk/openstack.net
                var ident = new OpenStackIdentityProvider(swiftCfg.authUri, new CloudIdentityWithProject()
                {
                    Username = swiftCfg.userName,
                    Password = swiftCfg.password,
                    ProjectName = "",
                    ProjectId = new ProjectId(swiftCfg.tenantId)
                });
                var cli = new CloudFilesProvider(null, ident);
                Trace.WriteLine("Uploading Object");
                var headers = new Dictionary<string, string>();
                headers.Add("X-Delete-After", "1209600");
                headers.Add("X-Object-Meta-Rpt-Filename", rptRqst.report.Filename);
                headers.Add("X-Object-Meta-Rpt-Email-TO", string.Join(", ", rptRqst.email.to));
                headers.Add("X-Object-Meta-Rpt-Email-CC", string.Join(", ", rptRqst.email.cc));
                cli.CreateObjectFromFile(swiftCfg.swiftContainer, filename, rptRqst.JobID + Path.GetExtension(filename), null, 4096, headers, swiftCfg.region);

                Trace.WriteLine("Get URL for Email");
                //Using openstack-sdk-dotnet @ https://github.com/stackforge/openstack-sdk-dotnet
                var credential = new OpenStackCredential(swiftCfg.authUri, swiftCfg.userName, swiftCfg.password, swiftCfg.tenantName, swiftCfg.region);
                var client = OpenStackClientFactory.CreateClient(credential);
                client.Connect().Wait();
                var swiftPublicEndpoint = credential.ServiceCatalog.GetPublicEndpoint("swift", swiftCfg.region);
                var swifturi = swiftPublicEndpoint + "/" + swiftCfg.swiftContainer + "/" + rptRqst.JobID + Path.GetExtension(filename);

                return "<a href=\"" + swifturi + "\">View Report</a>";
            }
            else { return null; }
        }
        static void SentRpt(RptDynamoConfig config, MailMessage mm, RptStatusAPI sAPI)
        {
            Trace.WriteLine("Emailing Report");

            // Transport configuration
            SmtpClient transport = new SmtpClient(config.smtp.server, config.smtp.port);
            if (config.smtp.username != null & config.smtp.password != null)
            { transport.Credentials = new NetworkCredential(config.smtp.username, config.smtp.password); }
            transport.EnableSsl = config.smtp.ssl;

            // Add From Address
            mm.From = new MailAddress(config.smtp.sender);

            try { transport.Send(mm); }
            catch { sAPI.failed().Wait(); }

            mm.Dispose();
            transport.Dispose();
        }
    }
    class Options
    {
        [Option('c', "config", Required = true, HelpText = "Specify Configuration File.")]
        public string config { get; set; }

        [Option('j', "job", Required = false, HelpText = "Specify Job File.")]
        public string job { get; set; }

        [Option('v', null, HelpText = "Print details during execution.")]
        public bool Verbose { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            // this without using CommandLine.Text
            //  or using HelpText.AutoBuild
            var usage = new StringBuilder();
            usage.AppendLine("RptDynamo Options");
            usage.AppendLine("Read user manual for usage instructions...");
            return usage.ToString();
        }
    }
}
