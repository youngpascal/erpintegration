using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ERPIntegration
{
    public partial class importCSV
    {
        public ExcelPackage bomPackage = new ExcelPackage();
        private Dictionary<part, string> bom = new Dictionary<part, string>();
        
        
        public void formatBoM()
        {
            ExcelWorksheet main = bomPackage.Workbook.Worksheets.Add("BoM");
            string newFilePath = "";
            int fOption = formatOption;

            //Add Headers
            main.Cells[1, 1].Value = "Parent";
            main.Cells[2, 1].Value = "Revision";
            main.Cells[4, 1].Value = "Item";
            main.Cells[4, 2].Value = "Part Number";
            main.Cells[4, 3].Value = "Units";
            main.Cells[4, 4].Value = "quantity";
            main.Cells[4, 5].Value = "Description";


            if (fOption == 1)
            {   //Simple format
                   
                newFilePath = Path.GetDirectoryName(csvPath) + "\\Simple Formatted" + Path.GetFileName(csvPath);
                int row = 5;

                for (int i = 1; i < bom.Count - 1; i++)
                {
                    foreach (KeyValuePair<part, string> part in bom)
                    {
                        if (part.Value.Equals("0"))
                        {
                            main.Cells[1, 2].Value = part.Key.itemNumber;
                            main.Cells[2, 2].Value = part.Key.revision;
                        }
                        else if (part.Value.Equals(i.ToString()))
                        {
                            main.Cells[row, 1].Value = part.Value;
                            main.Cells[row, 2].Value = part.Key.itemNumber;
                            main.Cells[row, 3].Value = part.Key.units;
                            main.Cells[row, 4].Value = part.Key.quantity;
                            main.Cells[row, 5].Value = part.Key.description;

                            row++;
                        }
                    }
                }
            }
            else if (fOption == 2)
            {   //Advanced format

                newFilePath = Path.GetDirectoryName(csvPath) + "\\Advanced Formatted" + Path.GetFileName(csvPath);
                //Advanced Headers
                main.Cells[4, 6].Value = "Type";
                main.Cells[4, 7].Value = "Comments";

                int row = 5;

                for (int i = 1; i < bom.Count - 1; i++)
                {
                    foreach (KeyValuePair<part, string> part in bom)
                    {
                        if (part.Value.Equals("0"))
                        {
                            main.Cells[1, 2].Value = part.Key.itemNumber;
                            main.Cells[2, 2].Value = part.Key.revision;
                        }
                        else if (part.Value.Equals(i.ToString()))
                        {
                            main.Cells[row, 1].Value = part.Value;
                            main.Cells[row, 2].Value = part.Key.itemNumber;
                            main.Cells[row, 3].Value = part.Key.units;
                            main.Cells[row, 4].Value = part.Key.quantity;
                            main.Cells[row, 5].Value = part.Key.description;
                            main.Cells[row, 6].Value = part.Key.dwg;
                            main.Cells[row, 7].Value = part.Key.purchasing;
                            row++;
                        }
                    }
                }
            }

            main.Cells["A:Z"].AutoFitColumns();

            bomPackage.SaveAs(new FileStream(newFilePath, FileMode.Create));
            bomPackage.Stream.Close();

            MessageBox.Show("Formatted BoM saved to \n" + newFilePath);
        }
    }
}
