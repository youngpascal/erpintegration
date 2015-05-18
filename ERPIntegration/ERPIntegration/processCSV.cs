﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//for parsing CSV files
using Microsoft.VisualBasic.FileIO;
namespace ERPIntegration
{
    public partial class importCSV
    {
        private List<part> parts = new List<part>();

        public void processCSV(string option)
        {
            parseCSV();

            if (option.Equals("Part"))
            {
                foreach (part newpart in parts)
                {
                    if (checkExisting(checkForItemCommand, newpart.itemNumber) == false)
                    {
                        createNewItem(newpart);
                        callIVRInsert(newpart.itemNumber);
                    }
                    else
                    {
                        updateExisting(newpart);
                    }
                }
            }
            else if (option.Equals("Parts"))
            {
                foreach (part newpart in parts)
                {
                    if (checkExisting(checkForItemCommand, newpart.itemNumber) == false)
                    {
                        createNewItem(newpart);
                        callIVRInsert(newpart.itemNumber);
                    }
                    else
                    {
                        updateExisting(newpart);
                    }
                }
            }
            else if (option.Equals("BOM"))
            {
                foreach (part newpart in parts)
                {
                    //do stuff
                }
            }
        }

        private static string reverseAndFormatEffDate(string s)
        {//Format the date to fit GP's needs
            char[] charArr = s.ToCharArray();
            char[] newD = new char[10];
            string newDate;

            for (int i = 0; i < charArr.Length; i++)
            {
                if (charArr[i].Equals('/'))
                {
                    if (i == 3)
                    {
                        newD[0] = charArr[4];
                        newD[1] = charArr[5];
                        newD[2] = charArr[6];
                        newD[3] = charArr[7];
                        newD[4] = '/';
                        newD[5] = '0';
                        newD[6] = charArr[0];
                        newD[7] = '/';
                        newD[8] = '0';
                        newD[9] = charArr[2];
                    }
                    else if (i == 4)
                    {
                        if (charArr[2].Equals('/'))
                        {
                            newD[0] = charArr[5];
                            newD[1] = charArr[6];
                            newD[2] = charArr[7];
                            newD[3] = charArr[8];
                            newD[4] = '/';
                            newD[5] = charArr[0];
                            newD[6] = charArr[1];
                            newD[7] = '/';
                            newD[8] = '0';
                            newD[9] = charArr[3];
                        }
                        else
                        {
                            newD[0] = charArr[5];
                            newD[1] = charArr[6];
                            newD[2] = charArr[7];
                            newD[3] = charArr[8];
                            newD[4] = '/';
                            newD[5] = '0';
                            newD[6] = charArr[0];
                            newD[7] = '/';
                            newD[8] = charArr[2];
                            newD[9] = charArr[3];
                        }
                    }
                    else if (i == 5)
                    {
                        newD[0] = charArr[6];
                        newD[1] = charArr[7];
                        newD[2] = charArr[8];
                        newD[3] = charArr[9];
                        newD[4] = '/';
                        newD[5] = charArr[0];
                        newD[6] = charArr[1];
                        newD[7] = '/';
                        newD[8] = charArr[3];
                        newD[9] = charArr[4];
                    }
                }
            }
            newDate = new string(newD);
            return newDate;
        }

        //Parse the CSV file and store in fields[]
        private void parseCSV()
        {
            TextFieldParser csvParser = new TextFieldParser(csvPath);
            csvParser.TextFieldType = FieldType.Delimited;
            csvParser.SetDelimiters(",");

            while (!csvParser.EndOfData)
            {   //Parse each line and add to the list of 'parts'
                fields = csvParser.ReadFields();
                parts.Add(new part(fields[0], fields[1], fields[2], fields[3],
                                       fields[4], fields[5], fields[6],
                                       fields[7], fields[8], fields[9], reverseAndFormatEffDate(fields[10])));
            }
            csvParser.Close();
        }//End parse CSV

        public struct part
        {
            public string level,
                          detailID,
                          position,
                          quantity,
                          itemNumber,
                          units,
                          description,
                          revision,
                          purchasing,
                          category,
                          effectiveDate;

            public part(string l, string d, string p,
                        string q, string i, string u,
                        string desc, string r, string pur,
                        string c, string e)
            {
                level = l;
                detailID = d;
                position = p;
                quantity = q;
                itemNumber = i;
                units = u;
                description = desc;
                revision = r;
                purchasing = pur;
                category = c;
                effectiveDate = e;
            }
        }
    }
}
