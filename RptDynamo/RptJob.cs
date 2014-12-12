using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RptDynamo
{

    public class Parameters
    {
        public string Name { get; set; }
        public bool MultipleSelect { get; set; }
        public List<string> SelectedValues { get; set; }
    }

    public class Database
    {
        public string type { get; set; }
        public string resource { get; set; }
        public string password { get; set; }
    }

    public class Report
    {
        public string Filename { get; set; }
        public List<Parameters> Parameters { get; set; }
        public Database database { get; set; }
        public string SelectedOutput { get; set; }
    }

    public class Email
    {
        public Custom custom { get; set; }
        public List<string> to { get; set; }
        public List<string> cc { get; set; }
    }

    public class Custom
    {
        public string subject { get; set; }
        public string body { get; set; }
        public bool supressparameters { get; set; }
    }
    public class RptJob
    {
        public Guid JobID { get; set; }
        public Report report { get; set; }
        public Email email { get; set; }
    }

}
