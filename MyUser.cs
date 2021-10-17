using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FP_Gen_1._0
{
    internal class MyUser
    {
        public string Username { get;set; }
        public string Password { get; set; }

        private static string error = "Username does not exist !";

        public static void ShowError()
        {
            System.Windows.Forms.MessageBox.Show(error);
        }
        public static bool IsEqual(MyUser user1, MyUser user2)
        {
            if (user1 == null || user2 == null) { return false; }

            if (user1.Username != user2.Username) { error = "Username does not exist !";
                return false;
            }

            if (user1.Password != user2.Password) { error = "Incorrect Password !";
                return false;
            }
            return true;
        }
    }
}
