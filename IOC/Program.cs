using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace IOC
{
    class Program
    {
        private static int colonLang;
        static void Main()
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\MultilingIOC9.1b.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int lastSpecie = 4;

            string[] langNames1 = new string[] { "Name", "Catalan", "Czech", "Estonian", "German", "Indonesian", "Latvian", "Norwegian", "Russian", "Spanish", "Ukrainian" };
            string[] langNames2 = new string[] { "English", "Chinese", "Danish", "Finnish", "Hungarian", "Italian", "Lithuanian", "Polish", "Slovak", "Swedish" };
            string[] langNames3 = new string[] { "Afrikaans", "Chinese-Traditional", "Dutch", "French", "Icelandic", "Japanese", "Northern-Sami", "Portuguese", "Slovenian", "Thai" };

            // Generated from below
            int[] speciesFamily = new int[] { 4, 12, 20, 37, 51, 191, 726, 1632, 1649, 1705, 2151, 2222, 2242, 2253, 2312, 2672, 2860, 3663, 3743, 3754, 3762, 3771, 4345, 5514, 5564, 6598, 6603, 6674, 7123, 7864, 8235, 9689, 9709, 9840, 9845, 10383, 10610, 11955, 12158, 13357 };

            //List<int> speciesFamily = new List<int>();

            //for (int i = 2; i <= rowCount; i++)
            //{
            //    if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
            //    {
            //        speciesFamily.Add(i);
            //        Console.WriteLine(i);
            //        //Console.Write(xlRange.Cells[i, 2].Value2.ToString() + "\t");
            //    }
            //}

            //TextWriter tw = new StreamWriter("SavedList.txt");

            //foreach (int s in speciesFamily)
            //    tw.WriteLine(s);

            //tw.Close();

            //for (var i = 0; i < speciesFamily.Length; i++)
            //{
            //    if (xlRange.Cells[speciesFamily[i], 2] != null && xlRange.Cells[speciesFamily[i], 2].Value2 != null)
            //    {
            //        Console.WriteLine("Specie family: " + xlRange.Cells[speciesFamily[i], 2].Value2);
            //        Console.WriteLine("Family: " + xlRange.Cells[speciesFamily[i] + 1, 2 + 1].Value2);
            //    }

            //    Console.Write("\r\n");
            //}

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            XmlWriter writer = XmlWriter.Create(@"Birds.xml", settings);

            writer.WriteStartDocument();
            writer.WriteStartElement("Birds");

            for (int i = 4; i <= rowCount; i++)
            {
                Console.WriteLine(i);
                // Start at scientific name
                int j = 4;

                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    writer.WriteStartElement("Bird");

                    foreach (int id in speciesFamily)
                    {
                        if (i > id)
                        {
                            lastSpecie = id;
                        }
                    }

                    if (xlRange.Cells[lastSpecie, 2] != null && xlRange.Cells[lastSpecie, 2].Value2 != null)
                    {
                        //Console.WriteLine("Specie: " + xlRange.Cells[lastSpecie, 2].Value2);
                        writer.WriteElementString("Order", xlRange.Cells[lastSpecie, 2].Value2);
                    }

                    if (xlRange.Cells[lastSpecie, 2] != null && xlRange.Cells[lastSpecie, 2].Value2 != null)
                    {
                        //Console.WriteLine("Family: " + xlRange.Cells[lastSpecie + 1, 2 + 1].Value2);
                        writer.WriteElementString("Family", xlRange.Cells[lastSpecie + 1, 2 + 1].Value2);
                    }

                    for (int colonLang = 0; colonLang <= 3; colonLang++)
                    {

                        for (int rowLang = 0; rowLang <= 30; rowLang += 3) // Fix so it's not just 30 but the acutual rowlenght
                        {
                            if (xlRange.Cells[i + colonLang, j + rowLang + colonLang] != null && xlRange.Cells[i + colonLang, j + rowLang + colonLang].Value2 != null)
                            {

                                if (colonLang == 0)
                                {
                                    //Console.WriteLine(langNames1[(rowLang / 3)] + ": " + xlRange.Cells[i + colonLang, j + rowLang + colonLang].Value2.ToString());

                                    writer.WriteElementString(langNames1[(rowLang / 3)], xlRange.Cells[i + colonLang, j + rowLang + colonLang].Value2.ToString());
                                }
                                else if (colonLang == 1)
                                {
                                    //Console.WriteLine(langNames2[(rowLang / 3)] + ": " + xlRange.Cells[i + colonLang, j + rowLang + colonLang].Value2.ToString());

                                    writer.WriteElementString(langNames2[(rowLang / 3)], xlRange.Cells[i + colonLang, j + rowLang + colonLang].Value2.ToString());
                                }
                                else if (colonLang == 2)
                                {
                                    //Console.WriteLine(langNames3[(rowLang / 3)] + ": " + xlRange.Cells[i + colonLang, j + rowLang + colonLang].Value2.ToString());
                                    writer.WriteElementString(langNames3[(rowLang / 3)], xlRange.Cells[i + colonLang, j + rowLang + colonLang].Value2.ToString());
                                }
                            }
                        }
                    }

                    writer.WriteEndElement();
                }
            }

            writer.WriteEndElement();
            writer.WriteEndDocument();
            writer.Flush();
            writer.Close();

            Console.Write("End");
            Console.ReadLine();
        }
    }
}
