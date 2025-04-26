using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace Unit_Teszt.Controller
{
    public class AccountController
    {
        public bool ValidateUserName(string username)
        {
            if (username == null) return false;
            return Regex.IsMatch(username, "^admin$");
        }

        public bool ValidatePassword(string password)
        {
            if (password == null) return false;
            return Regex.IsMatch(password, "^KrumpliFozelek2025$");
        }
    }
}
