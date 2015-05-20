using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//for SQL
using System.Data.SqlClient;
//for item creation
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;

namespace ERPIntegration
{
    public partial class importCSV
    {
        private string csvPath = "";
        private string[] fields;
        private string dtiGPConnection;
        public string vaultConnection = "";
        private string dtigpUsername = "";
        private string dtigpPassword = "";
        private string vaultUsername = "";
        private string vaultPassword = "";

        //importCSV constructor
        public importCSV(string gpUser, string gpPass, string vaultUser, string vaultPass, string filePath)
        {
            dtigpUsername = gpUser;
            dtigpPassword = gpPass;
            vaultUsername = vaultUser;
            vaultPassword = vaultPass;
            csvPath = filePath;
            dtiGPConnection = "Server=Sandbox-PC\\DTIGP;Database=DTI;User Id="+gpUser+";Password="+gpPass;
            vaultConnection = "Server=Sandbox-PC\\AUTODESKVAULT;Database=Vault;User Id=" + vaultUser + ";Password=" + vaultPass;
            gpsql = new SqlConnection(dtiGPConnection);
            adsksql = new SqlConnection(vaultConnection);
        }

    

       
    }
}
