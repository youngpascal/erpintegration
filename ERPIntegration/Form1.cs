using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ERPIntegration
{
    public partial class Form1 : Form
    {
        private static string gpUser, gpPass, vaultUser, vaultPass = "";
        private importCSV import;
        private string filePath = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //On load
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
        }

        private void button1_Click(object sender, EventArgs e)
        {   //Import multiple Items button
            createCSVObject();
            import.processCSV("Parts");
            MessageBox.Show("Parts imported successfully");
        }

        private void createCSVObject()
        {   //Get CSV path and create importCSV object
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

        private void button5_Click(object sender, EventArgs e)
        {   //SQL Connection button
            gpUser = textBox1.Text;
            gpPass = textBox2.Text;
            vaultUser = textBox3.Text;
            vaultPass = textBox4.Text;

            SqlConnection testGP = new SqlConnection("Server=Sandbox-PC\\DTIGP;Database=DTI;User Id=" + gpUser + ";Password=" + gpPass);
            testGP.Open();
            if (testGP.State == ConnectionState.Open)
            {
                DTIGPStatus.Text = "Connected!";
            }
            else
            {
                DTIGPStatus.Text = "Connection failed.";
            }
            testGP.Close();

            SqlConnection testADSK = new SqlConnection("Server=Sandbox-PC\\AUTODESKVAULT;Database=Vault;User Id=" + vaultUser + ";Password=" + vaultPass);
            testADSK.Open();
            if (testADSK.State == ConnectionState.Open)
            {
                ADSKVStatus.Text = "Connected!";
            }
            else
            {
                ADSKVStatus.Text = "Connection failed.";
            }
            testADSK.Close();
        }
    }
}
