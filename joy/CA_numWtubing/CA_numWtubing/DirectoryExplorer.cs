using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CA_numWtubing
{
    public class ConnectToSharedFolder : IDisposable
    {
        readonly string _networkName;

        public ConnectToSharedFolder(string networkName, NetworkCredential credentials)
        {
            _networkName = networkName;

            var netResource = new NetResource
            {
                Scope = ResourceScope.GlobalNetwork,
                ResourceType = ResourceType.Disk,
                DisplayType = ResourceDisplaytype.Share,
                RemoteName = networkName
            };

            var userName = string.IsNullOrEmpty(credentials.Domain)
                ? credentials.UserName
                : string.Format(@"{0}\{1}", credentials.Domain, credentials.UserName);

            var result = WNetAddConnection2(
                netResource,
                credentials.Password,
                userName,
                0);

            if (result != 0)
            {
                Console.WriteLine("Error connecting to remote share");
            }
        }

        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(NetResource netResource,
        string password, string username, int flags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags,
            bool force);

        public void Dispose()
        {
            //throw new NotImplementedException();
        }

        [StructLayout(LayoutKind.Sequential)]
        public class NetResource
        {
            public ResourceScope Scope;
            public ResourceType ResourceType;
            public ResourceDisplaytype DisplayType;
            public int Usage;
            public string LocalName;
            public string RemoteName;
            public string Comment;
            public string Provider;
        }

        public enum ResourceScope : int
        {
            Connected = 1,
            GlobalNetwork,
            Remembered,
            Recent,
            Context
        };

        public enum ResourceType : int
        {
            Any = 0,
            Disk = 1,
            Print = 2,
            Reserved = 8,
        }

        public enum ResourceDisplaytype : int
        {
            Generic = 0x0,
            Domain = 0x01,
            Server = 0x02,
            Share = 0x03,
            File = 0x04,
            Group = 0x05,
            Network = 0x06,
            Root = 0x07,
            Shareadmin = 0x08,
            Directory = 0x09,
            Tree = 0x0a,
            Ndscontainer = 0x0b
        }

    }

    class DirectoryExplorer
    {
        private Regex regEx;
        public PdfReader reader; // iTextSharp PDF object!
        private string sTempRev;
        private string[,] sTempFinalMB;
        private string sTempSheetText;
        private int iSheet;

        public List<string> getLatestRev(string rootDir)
        {
            List<string> latestRevs = new List<string>();
            
            try
            {
                Console.WriteLine("Getting Directories from " + rootDir);
                string[] subFolders = Directory.GetDirectories(rootDir);
                //string[] innerSubFolders = { }, innerFiles = { };
                foreach (string subFolder in subFolders)
                {                   
                    string folderName = subFolder.Remove(0, rootDir.Length);
                    string exc = folderName.Substring(0, 1);


                    //bool runProcess = true;
                    //Pause
                    //Quit
                    //Timer
                    //wqe should get output in file


                  //  if(runProcess()
                        

                    if (exc != "!")
                    {
                        Console.WriteLine("Opening Subfolder: " + subFolder);
                        string[] innerSubFolders = Directory.GetDirectories(subFolder);
                        string innerSubFolder;
                        for (int i = 0; i < innerSubFolders.Length; i++)
                        {
                            Console.WriteLine(innerSubFolders[i]);
                            innerSubFolder = innerSubFolders[i];



                            if (Directory.Exists(subFolder + @"/PDF CAD FILES"))
                            {
                                Console.WriteLine("The sub folder has a PDF CAD FILES folder");
                                // Get latest CA drawing




                                //    }
                                // foreach (string innerSubFolder in innerSubFolders)
                                //{
                                // string innerSubFolderName = innerSubFolder.Remove(0, subFolder.Length);
                                //if (innerSubFolderName == "/PDF CAD FILES")
                                //{
                                IDictionary<int, string> revs = new Dictionary<int, string>();
                                Console.WriteLine("Opening PDF CAD FILES inside Subfolder: " + subFolder);
                                string[] innerFiles = Directory.GetFiles(subFolder + @"/PDF CAD FILES");
                                foreach (string innerFile in innerFiles)
                                {
                                    string innerFileName = innerFile.Remove(0, subFolder.Length + 1);

                                    int found = innerFileName.IndexOf("CA_REV");
                                    int found1 = innerFileName.IndexOf("/");
                                    string catNo = innerFileName.Substring(found1 + 1, innerFileName.Length - (found1 + 2 + 3));



                                    if (found > 0)
                                    {
                                        //Console.WriteLine("CA REV found at index " + found + " of file: " + innerFileName);
                                        string croppedFileName = innerFileName.Remove(0, found + 6);

                                        if (croppedFileName.IndexOf(".pdf") == 1 || croppedFileName.IndexOf(".pdf") == 2)
                                        {
                                            if (croppedFileName.IndexOf(".pdf") == 1)
                                            {
                                                revs.Add(int.Parse(croppedFileName.Substring(0, 1)), innerFileName);
                                            }
                                            else if (croppedFileName.IndexOf(".pdf") == 2)
                                            {
                                                revs.Add(int.Parse(croppedFileName.Substring(0, 2)), innerFileName);
                                            }

                                        }
                                    }
                                }
                                int maxVal = 0;
                                if (revs != null)
                                { 
                                 maxVal = revs.Keys.Max();
                                Console.WriteLine("Latest Rev in " + folderName + " is :" + revs[maxVal]);
                            }
                                foreach (string innerFile in innerFiles)
                                {
                                    int found = innerFile.IndexOf(revs[maxVal]);
                                    if (found > 0)
                                    {
                                        latestRevs.Add(innerFile);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return latestRevs;
        }
        public string Find_PN(string sPath, string RegStr)
        {
            string str = null;

            string sDwgPath = sPath; // J.Mattrey added on 12-6-18               
            Regex Des = new Regex("DESCRIPTION"); //Find the master BOM
            Regex PN = new Regex(@"PART\sNUMBER"); //Find the master BOM
            Regex QTY = new Regex(@"QTY"); //Find the master BOM
            Regex Rev = new Regex(@"REVISIONS"); //Find the master BOM

            regEx = new Regex(RegStr, RegexOptions.IgnoreCase);   //@"(\w*[0-9]{4}L\w*-P)"
            //regEx = new Regex(@"([a-zA-Z0-9]+(\d[a-zA-Z0-9]+)+?((-?)+?([0-9]*)))");   

            // Get the PDFreader object(iTextSharp)
            try
            {
                NetworkCredential credentials = new NetworkCredential();
                //credentials.Domain = "DNNA";
                credentials.UserName = @"DNNA\S110489";
                credentials.Password = "April!2020";
                DateTime Latest_File = new DateTime(1900, 5, 1, 8, 30, 52); ;
                string Drawing_Path = null;
                // Console.WriteLine(System.IO.Path.GetDirectoryName(sPath));

                using (new ConnectToSharedFolder(System.IO.Path.GetDirectoryName(sPath), credentials))
                {
                    reader = new PdfReader(sPath);

                }

                SimpleTextExtractionStrategy simpleTextstrategy = new SimpleTextExtractionStrategy();

                sTempRev = null;
                sTempFinalMB = null;

                for (iSheet = 1; iSheet <= reader.NumberOfPages; iSheet++) // checking all the pages
                {

                    sTempSheetText = null;
                    sTempSheetText = PdfTextExtractor.GetTextFromPage(reader, iSheet);
                    //  Console.WriteLine(sTempSheetText); //convent the entire pdf into a string and print it out

                    if (Des.IsMatch(sTempSheetText) && PN.IsMatch(sTempSheetText) && QTY.IsMatch(sTempSheetText) && Rev.IsMatch(sTempSheetText) == false)
                    {
                        //Console.WriteLine(sTempSheetText); //convent the entire pdf into a string and print it out

                        if (regEx.IsMatch(sTempSheetText)) //for (int i = 0; i < DS.Tables[0].Rows.Count; i++) // change this to for loop??
                        {
                            var m = regEx.Matches(sTempSheetText);
                            foreach (Match mm in m)
                            {
                                Group g = mm.Groups[0];
                                //Console.WriteLine(g.ToString());
                                str = str + g.ToString() + ",";
                            }


                        }
                        break;
                    }
                }
                if (str != null)
                {
                    Console.WriteLine(str);
                    return str;
                }
                else
                    return "NO";
            }
            catch (Exception ex)
            {
                if (reader != null)
                    reader.Close();

        
                return "NO";
            
            }
        }

    }
}
