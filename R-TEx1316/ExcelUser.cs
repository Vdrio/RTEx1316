using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace R_TEx1316
{
    public enum UserAccessLevel
    {
        Default = 0, 
        Admin = 1
    }

    public class ExcelUser
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public UserAccessLevel AccessLevel { get; set; }

        public ExcelUser(string firstName, string lastName, UserAccessLevel accessLevel = UserAccessLevel.Default)
        {

        }

        public override string ToString()
        {
            string s = FirstName ?? "";
            s += " ";
            s += LastName ?? "";
            return s;
        }
    }
}
