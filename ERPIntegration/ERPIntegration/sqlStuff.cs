using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data;

namespace ERPIntegration
{
    public partial class importCSV
    {
        private SqlConnection gpsql;
        private string updateExistingPartRevision = "UPDATE IVR10015 SET REVISIONLEVEL_I=@revisionlevel WHERE ITEMNMBR=@itemnumber";
        private string updateExistingPartPurchasing = "UPDATE SY03900 SET TXTFIELD =@purchasing WHERE NOTEINDX= (SELECT NOTEINDX FROM IV00101 WHERE ITEMNMBR =@itemnumber)";
        private string updateExistingPart = "UPDATE IV00101 SET ITEMDESC=@description WHERE ITEMNMBR=@itemnumber";
        static string checkForItemCommand = "SELECT COUNT(*)FROM IV00101 WHERE ITEMNMBR = @itemnumber";

        private void openGPSQL()
        {
            try
            {
                gpsql.Open();
            }
            catch (Exception sqlException)
            {
                MessageBox.Show("Error opening SQL Connection. Full error: \n" + sqlException.ToString());
            }
        }

        private void closeSQL()
        {
            gpsql.Close();
        }

        public bool checkExisting(string command, string a1)
        {
            openGPSQL();

            bool found = false;

            try
            {
                SqlCommand myCommand = new SqlCommand(command, gpsql);
                myCommand.Parameters.AddWithValue("@itemnumber", a1);
                int itemCheck = Convert.ToInt32(myCommand.ExecuteScalar());

                //If the row does not exist
                if (itemCheck == 0)
                    found = false;
                else
                    found = true;
            }
            catch (Exception sqlCommandEx)
            {
                MessageBox.Show("Error parsing sql command" + command + "full error: " + sqlCommandEx.ToString());
            }

            closeSQL();

            return found;
        }

        public void callIVRInsert(string a1)
        {
            openGPSQL();

            try
            {
                SqlCommand myCommand = new SqlCommand("IVR_Insert", gpsql);
                myCommand.CommandType = CommandType.StoredProcedure;
                myCommand.Parameters.AddWithValue("@ItemNumber", a1);
                int reader = myCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command CallIVRInsert. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();
        }

        public void updateExisting(part newpart)
        {
            openGPSQL();
            try
            {
                SqlCommand myCommand = new SqlCommand(updateExistingPart, gpsql);
                SqlCommand myCommand2 = new SqlCommand(updateExistingPartRevision, gpsql);
                SqlCommand myCommand3 = new SqlCommand(updateExistingPartPurchasing, gpsql);

                myCommand.Parameters.AddWithValue("@description", newpart.description);
                myCommand.Parameters.AddWithValue("@itemnumber", newpart.itemNumber);

                myCommand2.Parameters.AddWithValue("@itemnumber", newpart.itemNumber);
                myCommand2.Parameters.AddWithValue("@revisionlevel", newpart.revision);

                myCommand3.Parameters.AddWithValue("@itemnumber", newpart.itemNumber);
                myCommand3.Parameters.AddWithValue("@purchasing", newpart.purchasing);

                int reader = myCommand.ExecuteNonQuery();
                int reader2 = myCommand2.ExecuteNonQuery();
                int reader3 = myCommand3.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command updateExisting. Contact database Administrator: " + ex.ToString());
            }
            closeSQL();
        }
    }
}
