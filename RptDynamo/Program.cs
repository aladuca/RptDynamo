﻿using System;
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

                RptEmail email = ProcessRpt(job);
                SentRpt(config, job, email);
                
            }
        }
        static RptEmail ProcessRpt(RptJob rptJob)
        {
            RptEmail email = new RptEmail();
            ReportDocument rpt = new ReportDocument();
            email.body = new StringBuilder();

            // Open Report
            Trace.WriteLine("Opening Report");
            rpt.Load(rptJob.report.Filename);
            email.subject = Path.GetFileNameWithoutExtension(rptJob.report.Filename);
            Trace.WriteLine("Loaded Report");

            // Pass Parameters to Report
            Trace.WriteLine("Passing Parameters");
            foreach (Parameter rptParam in rptJob.report.parameter)
            {
                Trace.WriteLine("Parameter: " + rptParam.Name + " is set to " + rptParam.text);
                email.body.AppendLine("Parameter: " + rptParam.Name + " is set to " + rptParam.text);
                rpt.SetParameterValue(rptParam.Name, rptParam.text);
            }

            // Create temporary output directory
            Trace.WriteLine("Creating temporary output directory");
            string tempdir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempdir);
            tempdir += "\\";
            Console.WriteLine("Created tmpdir " + tempdir);

            // Determine File Format and Output name
            string outfile = null;
            ExportFormatType crformat = ExportFormatType.NoFormat;
            switch (rptJob.report.output)
            {
                case "xlsx":
                    Trace.WriteLine("Output Format: Excel Workbook (2010+)");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xlsx");
                    crformat = ExportFormatType.ExcelWorkbook;
                    break;
                case "xls":
                    Trace.WriteLine("Output Format: Excel Workbook (pre 2010)");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xls");
                    crformat = ExportFormatType.Excel;
                    break;
                case "pdf":
                    Trace.WriteLine("Output Format: Portable Document Format");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".pdf");
                    crformat = ExportFormatType.PortableDocFormat;
                    break;
                case "csv":
                    Trace.WriteLine("Output Format: Comma-Seperated Values");
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".csv");
                    crformat = ExportFormatType.CharacterSeparatedValues;
                    break;
            }

            // Generate Report
            Trace.WriteLine("Exporting Report");
            rpt.ExportToDisk(crformat, outfile);
            email.file = outfile;

            // Clean up Crystal Reports ReportDocument
            rpt.Close();
            rpt.Dispose();

            return email;
        }
        static void SentRpt(RptDynamoConfig config, RptJob rptJob, RptEmail email)
        {
            Trace.WriteLine("Emailing Report");
            MailMessage mm = new MailMessage();
            SmtpClient transport = new SmtpClient(config.smtp.server, config.smtp.port)
            {
                Credentials = new NetworkCredential(config.smtp.username, config.smtp.password),
                EnableSsl = true
            };

            mm.From = new MailAddress(config.smtp.username);
            rptJob.email.to.ForEach(delegate(String name)
            {
                mm.To.Add(name);
            });
            rptJob.email.cc.ForEach(delegate(String name)
            {
                mm.CC.Add(name);
            });
            mm.Subject = email.subject;
            mm.Body = email.body.ToString();
            mm.IsBodyHtml = true;
            mm.Attachments.Add(new Attachment(email.file));

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
            usage.AppendLine("Quickstart Application 1.0");
            usage.AppendLine("Read user manual for usage instructions...");
            return usage.ToString();
        }
    }
}
