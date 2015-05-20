using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Globalization;
using Microsoft.Dynamics.GP.eConnect;
using Microsoft.Dynamics.GP.eConnect.Serialization;

namespace ERPIntegration
{
    public partial class importCSV
    {

        public void createNewItem(part newpart)
        {
            using (eConnectMethods eConCall = new eConnectMethods())
            {
                try
                {
                    // Create the eConnect document and store it in a file
                    serializeObject("C:\\newItem.xml", newpart);

                    // Load the eConnect XML document from the file into 
                    // and XML document object
                    XmlDocument xmldoc = new XmlDocument();
                    xmldoc.Load("C:\\newItem.xml");

                    // Create an XML string from the document object
                    string salesOrderDocument = xmldoc.OuterXml;

                    // Create a connection string
                    // Integrated Security is required. Integrated security=SSPI
                    // Update the data source and initial catalog to use the names of 
                    // data server and company database
                    string sConnectionString = "data source=Sandbox-PC\\DTIGP;initial catalog=DTI;integrated security=SSPI;persist security info=False;packet size=4096";

                    // Use the CreateTransactionEntity method to create the sales document in Microsoft Dynamics GP
                    // The method returns a string that contains the doc ID number of the new sales document
                    string salesOrder = eConCall.CreateTransactionEntity(sConnectionString, salesOrderDocument);
                }
                // The eConnectException class will catch any business logic related errors from eConnect_EntryPoint.
                catch (eConnectException exp)
                {
                    Console.Write(exp.ToString());
                }
                // Catch any system error that might occurr. Display the error to the user
                catch (Exception ex)
                {
                    Console.Write(ex.ToString());
                }
                finally
                {
                    // Use the Dipose method to release resources associated with the 
                    // eConnectMethods objects
                    eConCall.Dispose();
                }
            }
        }

        private static void serializeObject(string filename, part newpart)
        {
            // Create a datetime format object
            DateTimeFormatInfo dateFormat = new CultureInfo("en-US").DateTimeFormat;

            try
            {



                //create an eConnect schema object
                IVItemMasterType iv = new IVItemMasterType();

                taUpdateCreateItemRcd newItem = new taUpdateCreateItemRcd();


                newItem.ITEMNMBR = newpart.itemNumber; //Item number
                newItem.ITEMDESC = newpart.description; //Item description
                newItem.ITMSHNAM = "temp"; //short desc
                newItem.ITMGEDSC = "temp"; //general desc
                newItem.UseItemClass = 1;
                newItem.ITMCLSCD = newpart.category; //Part or Assembly
                newItem.ITEMTYPESpecified = true; //Say custom itemtype is being used
                newItem.ITEMTYPE = 1; //Sales item
                newItem.UOMSCHDL = newpart.units;
                newItem.DECPLCURSpecified = true; //Say a custom decplcur is being used
                newItem.DECPLCUR = 3; //Needed to make item in IVR10015
                newItem.NOTETEXT = newpart.purchasing;
                newItem.UpdateIfExists = 0;

                //Populate schema object with newItem info
                iv.taUpdateCreateItemRcd = newItem;

                // Create an array that holds ItemMasterType objects
                // Populate the array with the ItemMasterType schema object
                IVItemMasterType[] myItemType = { iv };

                // Create an eConnect XML document object and populate it 
                // with the ItemMasterType schema object
                eConnectType eConnect = new eConnectType();
                eConnect.IVItemMasterType = myItemType;

                // Create a file to hold the serialized eConnect XML document
                FileStream fs = new FileStream(filename, FileMode.Create);
                XmlTextWriter writer = new XmlTextWriter(fs, new UTF8Encoding());

                // Serialize the eConnect document object to the file using the XmlTextWriter.
                XmlSerializer serializer = new XmlSerializer(eConnect.GetType());
                serializer.Serialize(writer, eConnect);
                writer.Close();


            }
            //If an eConnect exception occurs, notify the user
            catch (eConnectException ex)
            {
                Console.Write(ex.ToString());
            }
        }
    }
}
