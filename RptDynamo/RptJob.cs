using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RptDynamo
{

    public class Parameter
    {
        public string Name { get; set; }
        public string MultipleValues { get; set; }
        public object text { get; set; }
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
        public List<Parameter> parameter { get; set; }
        public Database database { get; set; }
        public string output { get; set; }
    }

    public class Email
    {
        public List<string> to { get; set; }
        public List<string> cc { get; set; }
    }

    public class RptJob
    {
        public Report report { get; set; }
        public Email email { get; set; }
    }

}
