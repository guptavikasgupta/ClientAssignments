using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace CA_numWtubing
{
    class Program
    {
        static void Main(string[] args)
        {
            string rootDir = @"c:\users\vgupta\Desktop\testCA1\";
            string outputReport = @"c:\users\vgupta\Desktop\testCA1\Processoutput.csv";
            DirectoryExplorer d = new DirectoryExplorer();
            Console.WriteLine("Process Started");
            var strCSV = new StringBuilder();
            string str = "";
            string regxStr = "00119369PU-1|00119369PU-2";

            List<string> LatestRevs = d.getLatestRev(rootDir);

            if (LatestRevs != null && LatestRevs.Count > 0)
            {
                //Console.WriteLine("So this is the list of paths of Latest Revs:");
                try
                {
                    foreach (string listItem in LatestRevs)
                    {

                        Console.WriteLine("Looking in: " + listItem);
                        string PN = d.Find_PN(listItem, regxStr);

                        

                        string catno = listItem.Substring(listItem.LastIndexOf("\\") + 2, listItem.Length - (listItem.LastIndexOf("\\") + 2 + 3));

                       if (PN != "NO")
                        {
                            Console.WriteLine(PN + " is found.");
                            str = " Catalogue : " + catno + "Part :" + PN + " is found \n";
                            strCSV.AppendLine(str);
                        }
                        else
                        {
                            Console.WriteLine("Part nos are not found.");
                            str = "Part " +  PN + " is found \n";
                            strCSV.AppendLine(str);
                        }
                    }//end of for loop
                }//end of try
                catch (Exception ex)
                {
                Console.WriteLine(ex.Message);
                }
            }//end of if
            else
            {
                Console.WriteLine("Part nos are not found.");
                str = "Part nos are not found.";
                strCSV.AppendLine(str);
            }
            try
            {
                File.WriteAllText(outputReport, strCSV.ToString());
            }
            catch(Exception e)
            {

            }
            Console.ReadLine();
        }//end of main

    }//end of class
}//end of namespace
