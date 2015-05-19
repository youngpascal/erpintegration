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
    public partial class Form1 : Form
    {
        public static string gpUser, gpPass, vaultUser, vaultPass = "";
        private importCSV import;
        private string filePath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            loginForm login = new loginForm();
            login.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {   //Import multiple Items button
            createCSVObject();
            import.processCSV("Parts");
            MessageBox.Show("Parts imported successfully");
        }

        private void createCSVObject()
        {//Get CSV path and create importCSV object
            OpenFileDialog findCSV = new OpenFileDialog();

            System.Windows.Forms.DialogResult dr = findCSV.ShowDialog();

            if (dr == DialogResult.OK)
            {
                filePath = findCSV.FileName;
            }

            import = new importCSV(gpUser, gpPass, vaultUser, vaultPass, filePath);
        }

        private void button2_Click(object sender, EventArgs e)
        {   //Import one item
            createCSVObject();
            import.processCSV("Part");
            MessageBox.Show("Part imported successfully");
        }

        private void button3_Click(object sender, EventArgs e)
        {   //Import Bill of Materials
            createCSVObject();
            import.processCSV("BOM");
            MessageBox.Show("BoM imported successfully");
        }
    }
}
