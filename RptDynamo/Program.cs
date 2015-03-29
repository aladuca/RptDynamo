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

                RptStatusAPI sAPI = new RptStatusAPI(config.apiUri);
                sAPI.processing(Guid.Parse(Path.GetFileNameWithoutExtension(options.job))).Wait();
                RptEmail email = ProcessRpt(job, sAPI);
                SentRpt(config, job, email, sAPI);

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
            string tempdir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempdir);
            tempdir += "\\";
            Console.WriteLine("Created tmpdir " + tempdir);

            // Determine File Format and Output name
            string outfile = null;
            ExportFormatType crformat = ExportFormatType.NoFormat;
            switch (rptJob.report.SelectedOutput)
            {
                case "4":
                    Trace.WriteLine("Output Format: Excel Workbook (pre 2010)");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xls");
                    crformat = ExportFormatType.Excel;
                    break;
                case "8":
                    Trace.WriteLine("Output Format: Excel Workbook (pre 2010)");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xls");
                    crformat = ExportFormatType.ExcelRecord;
                    break;
                case "5":
                    Trace.WriteLine("Output Format: Portable Document Format");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".pdf");
                    crformat = ExportFormatType.PortableDocFormat;
                    break;
                case "10":
                    Trace.WriteLine("Output Format: Comma-Seperated Values");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".csv");
                    crformat = ExportFormatType.CharacterSeparatedValues;
                    break;
                case "15":
                    Trace.WriteLine("Output Format: Excel Workbook (2010+)");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xlsx");
                    crformat = ExportFormatType.ExcelWorkbook;
                    break;
            }

            // Generate Report
            Trace.WriteLine("Exporting Report");
            try
            {
                rpt.ExportToDisk(crformat, outfile);
                email.file = outfile;
                sAPI.completed().Wait();
            }
            catch (InternalException e)
            {
                sAPI.failed().Wait();
                Trace.WriteLine(e.Message);
                email.subject = "[RptDynamo] Error " + Path.GetFileNameWithoutExtension(rptJob.report.Filename);
                email.body.AppendLine("Error loading: " + rptJob.report.Filename);
                email.body.AppendLine("\r\n Please contact system administrator");
                email.body.AppendLine(e.Message);

            }
            catch (CrystalDecisions.CrystalReports.Engine.ParameterFieldCurrentValueException e)
            {
                sAPI.failed().Wait();
                email.body.AppendLine("<br/><br/><font color=\"red\"><strong>Crystal Reports Error:</strong> " + e.Message + "</font>");
                email.body.AppendLine("\r\n Please contact system administrator");
                email.body.AppendLine(e.Message);
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                sAPI.failed().Wait();
                email.body.AppendLine("<br/><br/><font color=\"red\"><strong>Crystal Reports Error:</strong> " + e.Message + "</font>");
                email.body.AppendLine("\r\n Please contact system administrator");
                email.body.AppendLine(e.Message);
            }
            catch
            {
                sAPI.failed().Wait();
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

            mm.From = new MailAddress(config.smtp.username);
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
            mm.Body = email.body.Replace(Environment.NewLine, Environment.NewLine + "<br/>").ToString();
            mm.IsBodyHtml = true;
            if (email.file != null) mm.Attachments.Add(new Attachment(email.file));

            transport.Send(mm);

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
