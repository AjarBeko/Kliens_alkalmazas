using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Guna.UI2.WinForms;

namespace Szamla_Kliensalkalmazas
{
    public partial class LoginForm : Form
    {
        public LoginForm()
        {
            InitializeComponent();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            string username = guna2TextBox1.Text;
            string password = guna2TextBox2.Text;

            
            if (username == "admin" && password == "1234")
            {
                
                Form1 main = new Form1();
                this.Hide();
                main.ShowDialog();
                this.Close();
            }
            else
            {
                label2.Text = "Hibás felhasználónév vagy jelszó!";
                label2.ForeColor = Color.Red;
            }
        }

        private void guna2TextBox2_IconRightClick(object sender, EventArgs e)
        {
            if (guna2TextBox2.UseSystemPasswordChar)
            {
                guna2TextBox2.UseSystemPasswordChar = false;
                guna2TextBox2.IconRight = Properties.Resources.closed_eyes;
            }
            else
            {
                guna2TextBox2.UseSystemPasswordChar = true;
                guna2TextBox2.IconRight = Properties.Resources.eyeball_icon_png_eye_icon_1;
            }
        }

    }
}
