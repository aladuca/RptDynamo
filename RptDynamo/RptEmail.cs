using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RptDynamo
{
    class RptEmail
    {
        public RptEmail()
        {
            body = new StringBuilder();
        }
        public String subject { get; set; }
        public StringBuilder body { get; set; }
        public String file { get; set; }
    }
}
