using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ERPIntegration
{
    public partial class loginForm : Form
    {

        public loginForm()
        {
            InitializeComponent();
            this.TopMost = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();

            Form1.gpUser = textBox1.Text;
            Form1.gpPass = textBox2.Text;
            Form1.vaultUser = textBox3.Text;
            Form1.vaultPass = textBox4.Text;
        }
    }
}
