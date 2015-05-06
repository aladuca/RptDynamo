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

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;


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
                RptEmail email = ProcessRpt(job, sAPI);
                SentRpt(config, job, email, sAPI);
                Directory.Delete(tempPath, true);
            }
        }
        static RptEmail ProcessRpt(RptJob rptJob, RptStatusAPI sAPI)
        {
            RptEmail email = new RptEmail();
            ReportDocument rpt = new ReportDocument();

            // Open Report
            Trace.WriteLine("Opening Report");
            try
            {
                rpt.Load(rptJob.report.Filename);
                rpt.Refresh(); // Refresh data if saved
                Trace.WriteLine("Loaded Report");
            }
            catch
            {
                sAPI.failed().Wait();
                Trace.WriteLine("Error: failure loading report");
                email.subject = "[RptDynamo] Error " + Path.GetFileNameWithoutExtension(rptJob.report.Filename);
                email.body.AppendLine("Error loading: " + rptJob.report.Filename);
                email.body.AppendLine("\r\n Please contact system administrator");

                // Clean up Crystal Reports ReportDocument
                rpt.Close();
                rpt.Dispose();

                return email;
            }

            // Set Title based on if it specified in job > report > filename
            try { email.subject = rptJob.email.custom.subject; }
            catch
            {
                if (rpt.SummaryInfo.ReportTitle != null)
                    email.subject = "[RptDynamo] " + rpt.SummaryInfo.ReportTitle;
                else
                    email.subject = "[RptDynamo] " + Path.GetFileNameWithoutExtension(rptJob.report.Filename);
            }

            // Seed custom body
            try { email.body.AppendLine(rptJob.email.custom.body + "\r\n"); }
            catch { };

            // Pass Parameters to Report
            if (rptJob.report.Parameters != null)
            {
                Trace.WriteLine("Passing Parameters");
                foreach (Parameters rptParam in rptJob.report.Parameters)
                {
                    Trace.WriteLine("Parameter: " + rptParam.Name + " is set to " + string.Join(", ", rptParam.SelectedValues));
                    try { rpt.SetParameterValue(rptParam.Name, rptParam.SelectedValues.ToArray()); }
                    catch { Trace.WriteLine("Failed to set " + rptParam.Name); }
                    if (rptJob.email.custom == null)
                        email.body.AppendLine("Parameter: " + rptParam.Name + " is set to " + string.Join(", ", rptParam.SelectedValues));
                    else
                    {
                        if (rptJob.email.custom.supressparameters)
                            email.body.AppendLine("Parameter: " + rptParam.Name + " is set to " + string.Join(", ", rptParam.SelectedValues));
                    }
                }
            }
            else { Trace.WriteLine("No Parameters Passed"); }

            Trace.WriteLine("Creating temporary output directory");
            string tempdir = Path.Combine(Path.GetTempPath(), "export");
            Directory.CreateDirectory(tempdir);
            tempdir += "\\";
            Console.WriteLine("Created tmpdir " + tempdir);

            // Determine File Format and Output name
            ExportOptions CrExportOptions = rpt.ExportOptions;
            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            switch (rptJob.report.SelectedOutput)
            {
                case "9":
                    Trace.WriteLine("Output Format: Text");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".txt");
                    CrExportOptions.ExportFormatType = ExportFormatType.Text;

                    TextFormatOptions CrFormatTypeOptions = new TextFormatOptions();
                    CrFormatTypeOptions.LinesPerPage = 0;
                    CrExportOptions.ExportFormatOptions = CrFormatTypeOptions;
                    break;
                case "4":
                    Trace.WriteLine("Output Format: Excel Workbook (pre 2010)");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xls");
                    CrExportOptions.ExportFormatType = ExportFormatType.Excel;
                    break;
                case "8":
                    Trace.WriteLine("Output Format: Excel Workbook (pre 2010)");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xls");
                    CrExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                    break;
                case "5":
                    Trace.WriteLine("Output Format: Portable Document Format");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".pdf");
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    break;
                case "10":
                    Trace.WriteLine("Output Format: Comma-Seperated Values");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".csv");
                    CrExportOptions.ExportFormatType = ExportFormatType.CharacterSeparatedValues;
                    break;
                case "15":
                    Trace.WriteLine("Output Format: Excel Workbook (2010+)");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xlsx");
                    CrExportOptions.ExportFormatType = ExportFormatType.ExcelWorkbook;
                    break;
            }

            // Generate Report
            Trace.WriteLine("Exporting Report");
            try
            {
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                rpt.Export();

                email.file = CrDiskFileDestinationOptions.DiskFileName;
            }
            catch (CrystalDecisions.CrystalReports.Engine.InternalException e)
            {
                sAPI.failed().Wait();
                Trace.WriteLine("Error: " + e.Message);
                email.subject = "[RptDynamo] Error " + Path.GetFileNameWithoutExtension(rptJob.report.Filename);
                email.body.AppendLine("Error loading: " + rptJob.report.Filename);
                email.body.AppendLine(e.ToString());
                email.body.AppendLine("\r\n Please contact system administrator");
                email.body.AppendLine(e.Message);

            }
            catch (CrystalDecisions.CrystalReports.Engine.ParameterFieldCurrentValueException e)
            {
                sAPI.failed().Wait();
                Trace.WriteLine("Crystal Reports Error: " + e.Message);
                email.body.AppendLine("<br/><br/><font color=\"red\"><strong>Crystal Reports Error:</strong> " + e.Message + "</font>");
                email.body.AppendLine(e.ToString());
                email.body.AppendLine("\r\n Please contact system administrator");
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                sAPI.failed().Wait();
                Trace.WriteLine("COM Error: " + e.Message);
                email.body.AppendLine("<br/><br/><font color=\"red\"><strong>Crystal Reports Error:</strong> " + e.Message + "</font>");
                email.body.AppendLine(e.ToString());
                email.body.AppendLine("\r\n Please contact system administrator");
            }
            catch
            {
                sAPI.failed().Wait();
                Trace.WriteLine("ExportToDisk: Unknown Error ");
                email.body.AppendLine("<br/><br/><font color=\"red\"><strong>Unknown Error</strong></font>");
                email.body.AppendLine("\r\n Please contact system administrator");
            }

            // Clean up Crystal Reports ReportDocument
            rpt.Close();
            rpt.Dispose();

            return email;
        }
        static void SentRpt(RptDynamoConfig config, RptJob rptJob, RptEmail email, RptStatusAPI sAPI)
        {
            Trace.WriteLine("Emailing Report");
            MailMessage mm = new MailMessage();

            // Transport configuration
            SmtpClient transport = new SmtpClient(config.smtp.server, config.smtp.port);
            if (config.smtp.username != null & config.smtp.password != null)
            { transport.Credentials = new NetworkCredential(config.smtp.username, config.smtp.password); }
            transport.EnableSsl = config.smtp.ssl;

            mm.From = new MailAddress(config.smtp.sender);
            if (rptJob.email.to != null)
            {
                rptJob.email.to.ForEach(delegate(String name)
                {
                    mm.To.Add(name);
                });
            }
            else
            {
                Trace.WriteLine("No mail \"to\" addresses specified - Email not sent");
                return;
            }
            if (rptJob.email.cc != null)
            {
                rptJob.email.cc.ForEach(delegate(String name)
                {
                    mm.CC.Add(name);
                });
            }
            else { Trace.WriteLine("No mail \"cc\" addresses specified"); }
            mm.Subject = email.subject;

            if (email.file != null)
            {
                email.body.AppendLine();
                if (config.swiftCfg != null)
                {
                    Trace.WriteLine("Getting OpenStack Crudential");
                    // Using openstack.net @ https://github.com/openstacknetsdk/openstack.net
                    var ident = new OpenStackIdentityProvider(config.swiftCfg.authUri, new CloudIdentityWithProject()
                    {
                        Username = config.swiftCfg.userName,
                        Password = config.swiftCfg.password,
                        ProjectName = "",
                        ProjectId = new ProjectId(config.swiftCfg.tenantId)
                    });
                    var cli = new CloudFilesProvider(null, ident);
                    Trace.WriteLine("Uploading Object");
                    var headers = new Dictionary<string, string>();
                    headers.Add("X-Delete-After", "1209600");
                    headers.Add("X-Object-Meta-Rpt-Filename", rptJob.report.Filename);
                    headers.Add("X-Object-Meta-Rpt-Email-TO", string.Join(", ", rptJob.email.to));
                    headers.Add("X-Object-Meta-Rpt-Email-CC", string.Join(", ", rptJob.email.cc));
                    cli.CreateObjectFromFile(config.swiftCfg.swiftContainer, email.file, rptJob.JobID + Path.GetExtension(email.file), null, 4096, headers, config.swiftCfg.region);

                    Trace.WriteLine("Get URL for Email");
                    //Using openstack-sdk-dotnet @ https://github.com/stackforge/openstack-sdk-dotnet
                    var credential = new OpenStackCredential(config.swiftCfg.authUri, config.swiftCfg.userName, config.swiftCfg.password, config.swiftCfg.tenantName, config.swiftCfg.region);
                    var client = OpenStackClientFactory.CreateClient(credential);
                    client.Connect().Wait();
                    var swiftPublicEndpoint = credential.ServiceCatalog.GetPublicEndpoint("swift", config.swiftCfg.region);
                    var swifturi = swiftPublicEndpoint + "/" + config.swiftCfg.swiftContainer + "/" + rptJob.JobID + Path.GetExtension(email.file);

                    email.body.AppendLine("<a href=\"" + swifturi + "\">View Report</a>");
                }
                else
                {
                    var fileInfo = new FileInfo(email.file);
                    if (fileInfo.Length < 26214400) { mm.Attachments.Add(new Attachment(email.file)); }
                    else
                    {
                        email.body.AppendLine("<br/><br/><font color=\"red\"><strong>Report export too large to be attached.</strong></font>");
                        File.Copy(email.file, "C:\\Users\\Public\\Documents\\RptPutty-RptDynamo\\Exports\\" + rptJob.JobID + Path.GetExtension(email.file));
                    }
                }
            }

            mm.Body = email.body.Replace(Environment.NewLine, Environment.NewLine + "<br/>").ToString();
            mm.IsBodyHtml = true;

            transport.Send(mm);

            sAPI.completed().Wait();

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
