using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;

using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace RptDynamo
{
    //Taken from http://www.superstarcoders.com/blogs/posts/executing-code-in-a-separate-application-domain-using-c-sharp.aspx
    public sealed class Isolated<T> : IDisposable where T : MarshalByRefObject
    {
        private AppDomain _domain;
        private T _value;

        public Isolated()
        {
            _domain = AppDomain.CreateDomain("Isolated:" + Guid.NewGuid(),
               null, AppDomain.CurrentDomain.SetupInformation);

            Type type = typeof(T);

            _value = (T)_domain.CreateInstanceAndUnwrap(type.Assembly.FullName, type.FullName);
        }

        public T Value
        {
            get
            {
                return _value;
            }
        }

        public void Dispose()
        {
            if (_domain != null)
            {
                AppDomain.Unload(_domain);

                _domain = null;
            }
        }
    }
    public class Facade : MarshalByRefObject
    {
        private SvcExport _existing = new SvcExport();

        public string Export(string rpt)
        {
            return _existing.RptExport(Newtonsoft.Json.JsonConvert.DeserializeObject<Report>(rpt));
        }
    } 
    class SvcExport
    {
        public string RptExport(Report rpt)
        {
            //Inialize Variables
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            ReportDocument CRRptDoc = new ReportDocument();

            //Load Report
            CRRptDoc.Load(rpt.Filename);

            //Pass Parameters
            if (rpt.Parameters != null)
            {
                foreach (Parameters parameter in rpt.Parameters)
                {
                    CRRptDoc.SetParameterValue(parameter.Name, parameter.SelectedValues.ToArray());
                }
            }

            //Create Export Directory
            Trace.WriteLine("ExportISO: " + Path.GetTempPath());
            string tempdir = Path.Combine(Path.GetTempPath(), "export");
            Directory.CreateDirectory(tempdir);
            tempdir += "\\";

            //Export Options
            ExportOptions CrExportOptions = CRRptDoc.ExportOptions;
            DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            switch (rpt.SelectedOutput)
            {
                case "9":
                    Trace.WriteLine("ExportISO: " + "Output Format: Text");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rpt.Filename), ".txt");
                    CrExportOptions.ExportFormatType = ExportFormatType.Text;

                    TextFormatOptions CrFormatTypeOptions = new TextFormatOptions();
                    CrFormatTypeOptions.LinesPerPage = 0;
                    CrExportOptions.ExportFormatOptions = CrFormatTypeOptions;
                    break;
                case "4":
                    Trace.WriteLine("ExportISO: " + "Output Format: Excel Workbook (pre 2010)");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rpt.Filename), ".xls");
                    CrExportOptions.ExportFormatType = ExportFormatType.Excel;
                    break;
                case "8":
                    Trace.WriteLine("ExportISO: " + "Output Format: Excel Workbook (pre 2010)");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rpt.Filename), ".xls");
                    CrExportOptions.ExportFormatType = ExportFormatType.ExcelRecord;
                    break;
                case "5":
                    Trace.WriteLine("ExportISO: " + "Output Format: Portable Document Format");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rpt.Filename), ".pdf");
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    break;
                case "10":
                    Trace.WriteLine("ExportISO: " + "Output Format: Comma-Seperated Values");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rpt.Filename), ".csv");
                    CrExportOptions.ExportFormatType = ExportFormatType.CharacterSeparatedValues;
                    break;
                case "15":
                    Trace.WriteLine("ExportISO: " + "Output Format: Excel Workbook (2010+)");
                    CrDiskFileDestinationOptions.DiskFileName = tempdir + Path.ChangeExtension(Path.GetFileName(rpt.Filename), ".xlsx");
                    ExcelFormatOptions crfomat = new ExcelFormatOptions();
                    CrExportOptions.ExportFormatType = ExportFormatType.ExcelWorkbook;
                    break;
            }

            CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
            CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
            CRRptDoc.Export();

            CRRptDoc.Close();
            CRRptDoc.Dispose();

            return CrDiskFileDestinationOptions.DiskFileName;
        }
    }
}
