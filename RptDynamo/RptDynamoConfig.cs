using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RptDynamo
{
    public class Smtp
    {
        public string server { get; set; }
        public string address { get; set; }
        public int port { get; set; }
        public string username { get; set; }
        public string password { get; set; }
        public string sender { get; set; }
        public bool ssl { get; set; }
    }

    public class Queue
    {
        public string type { get; set; }
        public string name { get; set; }
    }

    public class RptDynamoConfig
    {
        public Smtp smtp { get; set; }
        public Queue queue { get; set; }
        public String apiUri { get; set; }
    }
}
