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
        private string checkForItemCommand = "SELECT COUNT(*)FROM IV00101 WHERE ITEMNMBR = @itemnumber";
        private string checkForExistingBOMLine = "SELECT COUNT(*) FROM IPG_AD_BOM_Line WHERE PPN_I = @ppn AND CPN_I = @cpn";
        private string deleteTempBoMLine = "DELETE FROM IPG_AD_BOM_Line";
        private string deleteTempBoMHeader = "DELETE FROM IPG_AD_BOM_Header";
        private string deleteFromActiveHeader = "DELETE FROM BM010415 WHERE ITEMNMBR = @ppn";
        private string deleteFromActiveLine = "DELETE FROM BM010115 WHERE (PPN_I = @ppn AND CPN_I = @cpn)";

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

        public bool checkExistingLine(string command, string a1, string a2)
        {
            openGPSQL();

            bool found = false;

            try
            {
                SqlCommand myCommand = new SqlCommand(command, gpsql);
                myCommand.Parameters.AddWithValue("@ppn", a1);
                myCommand.Parameters.AddWithValue("@cpn", a2);
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

        public void deleteBoMHeaderTemp()
        {
            openGPSQL();

            try
            {
                SqlCommand delete = new SqlCommand(deleteTempBoMHeader, gpsql);

                int reader = delete.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command deleteBoMHeaderTemp. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();
        }

        public void deleteBoMLineTemp()
        {
            openGPSQL();

            try
            {
                SqlCommand delete = new SqlCommand(deleteTempBoMLine, gpsql);

                int reader = delete.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command deleteTempBoMLine. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();
        }

        public void addBoMHeaderTemp(part headerPart)
        {
            openGPSQL();

            try
            {

                SqlCommand myCommand = new SqlCommand("Insert_BoMHeaderTemp", gpsql);
                myCommand.CommandType = CommandType.StoredProcedure;
                myCommand.Parameters.AddWithValue("@ItemNumber", headerPart.itemNumber);
                myCommand.Parameters.AddWithValue("@Revision", headerPart.revision);
                myCommand.Parameters.AddWithValue("@EffectiveDate", headerPart.effectiveDate);
                
                int reader = myCommand.ExecuteNonQuery();
                

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command AddTempBoMHeader. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();
        }

        public void addBoMLineTemp(part linePart)
        {
            openGPSQL();

            int subCat = 0;

            if (linePart.category.Equals("Part"))
            {
                subCat = 1;
            }

            if (linePart.position != "")
            {
                try
                {

                    SqlCommand myCommand = new SqlCommand("Insert_BoMLineTemp", gpsql);
                    myCommand.CommandType = CommandType.StoredProcedure;
                    myCommand.Parameters.AddWithValue("@ParentNumber", linePart.parent);
                    myCommand.Parameters.AddWithValue("@ChildNumber", linePart.itemNumber);
                    myCommand.Parameters.AddWithValue("@SubCat", subCat);
                    myCommand.Parameters.AddWithValue("@BoMSeq", linePart.position);
                    myCommand.Parameters.AddWithValue("@Position", Int32.Parse(linePart.position));
                    myCommand.Parameters.AddWithValue("@Quantity", float.Parse(linePart.quantity));
                    myCommand.Parameters.AddWithValue("@EffectiveDate", linePart.effectiveDate);
                    myCommand.Parameters.AddWithValue("@Units", linePart.units);
                    
                    int reader = myCommand.ExecuteNonQuery();
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error performing SQL command AddBoMTempLinevarchar . Contact database Administrator: " + ex.ToString());
                }
            }
            else
            {
                MessageBox.Show("Part number: " + linePart.itemNumber + " will not be added because no position value was given.");
            }

            closeSQL();
        }

        public void moveBoMHeaderToHistory()
        {
            openGPSQL();

            try
            {
                SqlCommand myCommand = new SqlCommand("MoveBoMHeaderToHistory", gpsql);
                myCommand.CommandType = CommandType.StoredProcedure;
                int reader = myCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command MoveBoMHeaderToHistory. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();
        }

        public void moveBoMLineToHistory()
        {
            openGPSQL();

            try
            {
                SqlCommand myCommand = new SqlCommand("MoveBoMLineToHistory", gpsql);
                myCommand.CommandType = CommandType.StoredProcedure;
                int reader = myCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command MoveBoMLineToHistory. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();
        }

        public void deleteActiveBoMLine(part newpart)
        {
           /* openGPSQL();

            try
            {
                SqlCommand myCommand = new SqlCommand("DeleteActiveBoMLine", gpsql);
                myCommand.CommandType = CommandType.StoredProcedure;
                int reader = myCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command deleteActiveBoMLine. Contact database Administrator: " + ex.ToString());
            }

            closeSQL(); */

            openGPSQL();
            try
            {
                SqlCommand myCommand = new SqlCommand(deleteFromActiveLine, gpsql);

                myCommand.Parameters.AddWithValue("@ppn", newpart.parent);
                myCommand.Parameters.AddWithValue("@cpn", newpart.itemNumber);

                int reader = myCommand.ExecuteNonQuery();


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command deleteActiveLine. Contact database Administrator: " + ex.ToString());
            }
            closeSQL();
        }

        public void deleteActiveBoMHeader(part newpart)
        {
            /*openGPSQL();

            try
            {
                SqlCommand myCommand = new SqlCommand("MoveBoMHeaderToHistory", gpsql);
                myCommand.CommandType = CommandType.StoredProcedure;
                int reader = myCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command deleteActiveBoMHeader. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();*/

            openGPSQL();
            try
            {
                SqlCommand myCommand = new SqlCommand(deleteFromActiveHeader, gpsql);

                myCommand.Parameters.AddWithValue("@ppn", newpart.itemNumber);

                int reader = myCommand.ExecuteNonQuery();


            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command deleteActiveHeader. Contact database Administrator: " + ex.ToString());
            }
            closeSQL();
        }

        public void insertActiveBoMHeader()
        {
            openGPSQL();

            try
            {
                SqlCommand myCommand = new SqlCommand("Insert_ActiveBoMHeader", gpsql);
                myCommand.CommandType = CommandType.StoredProcedure;
                int reader = myCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command deleteActiveBoMHeader. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();
        }

        public void insertActiveBoMLine()
        {
            openGPSQL();

            try
            {
                SqlCommand myCommand = new SqlCommand("Insert_ActiveBoMLine", gpsql);
                myCommand.CommandType = CommandType.StoredProcedure;
                int reader = myCommand.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error performing SQL command insertActiveBoMLine. Contact database Administrator: " + ex.ToString());
            }

            closeSQL();
        }

    }
}
