using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Runtime.Serialization.Json;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
                    Console.WriteLine(options.config);
                    Console.WriteLine(options.job);
                    string json = File.ReadAllText(options.job);
                    DataContractJsonSerializer jobser = new DataContractJsonSerializer(typeof(RptJob));
                    StreamReader file = File.OpenText(options.job);
                    Stream stream = File.OpenRead(options.job);
                    RptJob job = (RptJob)jobser.ReadObject(stream);
                    ProcessRpt(job);
                }
                else
                    Console.WriteLine("working ...");
            }
        }
        static void ProcessRpt(RptJob rptJob)
        {
            ReportDocument rpt = new ReportDocument();

            // Open Report
            rpt.Load(rptJob.report.Filename);
            Trace.WriteLine("Loaded Report");

            // Pass Parameters to Report
            foreach (Parameter rptParam in rptJob.report.parameter)
            {
                rpt.SetParameterValue(rptParam.Name, rptParam.text);
            }

            // Create temporary output directory
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
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xlsx");
                    crformat = ExportFormatType.ExcelWorkbook;
                    break;
                case "xls":
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".xls");
                    crformat = ExportFormatType.Excel;
                    break;
                case "pdf":
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".pdf");
                    crformat = ExportFormatType.PortableDocFormat;
                    break;
                case "csv":
                    outfile = tempdir + Path.ChangeExtension(Path.GetFileName(rptJob.report.Filename), ".csv");
                    crformat = ExportFormatType.CharacterSeparatedValues;
                    break;
            }

            // Generate Report
            rpt.ExportToDisk(crformat, outfile);

            // Clean up Crystal Reports ReportDocument
            rpt.Close();
            rpt.Dispose();
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
