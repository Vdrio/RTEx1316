using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace R_TEx1316
{
    public class ExcelUser
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }

        public override string ToString()
        {
            string s = FirstName ?? "";
            s += " ";
            s += LastName ?? "";
            return s;
        }
    }
}
