using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FireSharp.Config;
using FireSharp.Response;
using FireSharp.Interfaces;

namespace FP_Gen_1._0
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }

        IFirebaseConfig ifc = new FirebaseConfig()
        {
            AuthSecret = "rFICeYegSf9pLssLvTqKGEC4gejYJzhyaFlWK0Dm",
            BasePath = "https://fp-gen-default-rtdb.europe-west1.firebasedatabase.app/"
        };

        IFirebaseClient client;
        private void logBtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(userBox.Text) &&
                string.IsNullOrWhiteSpace(passBox.Text))
            {
                MessageBox.Show("Please Fill All the Fields !");
                return;
            }

            FirebaseResponse res = client.Get(@"Users/" + userBox.Text);
            MyUser ResUser = res.ResultAs<MyUser>();
            MyUser CurUser = new MyUser()
            {
                Username = userBox.Text,
                Password = passBox.Text
            };

            if (MyUser.IsEqual(ResUser, CurUser))
            {
                main f1 = new main();
                this.Hide();
                f1.ShowDialog();
            }
            else
            {
                MyUser.ShowError();
            }

        }

        private void Login_Load(object sender, EventArgs e)
        {
            try
            {
                client = new FireSharp.FirebaseClient(ifc);
            }
            catch
            {
                MessageBox.Show("No internet connection");
            }
        }

        private void passBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
