using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BoschTool
{
    public partial class Form1 : Form
    {
        WebPage web;
        public Dictionary<string, string> arrInputs;
        public Form1(WebPage page)
        {
            web = page;
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            arrInputs["txtAccountName"] = this.textBox1.Text;
            arrInputs["txtPassword"] = this.textBox2.Text;
            arrInputs["txtVerifyCode"] = this.textBox3.Text;
            bool state = web.LoginWebServer(arrInputs);
            Globals.Ribbons.Ribbon1.LoginState = state;
            if (state)
                MessageBox.Show("Login success!", "BOSCH", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
        }
    }
}
