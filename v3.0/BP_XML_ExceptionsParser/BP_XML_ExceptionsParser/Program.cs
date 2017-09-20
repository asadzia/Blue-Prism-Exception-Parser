///
/// Author: Asad Zia
/// Version: 1.0
///

using System;
using System.Collections;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;

namespace BP_XML_ExceptionsParser
{
    /// <summary>
    /// Extract Exceptions - Blue Prism
    /// </summary>
    public class Program
    {
        /// Define paramters
        
        /// <summary>
        /// A method to escape characters
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public string checkEscape (string str)
        {
             if (str.Contains(",") || str.Contains("\""))
                {
                    str = '"' + str.Replace("\"", "\"\"") + '"';
                }
            return str;
        }

        /// <summary>
        /// The method for getting the file name from the file path
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public string getFileName (string fileName)
        {
            return fileName.Substring(fileName.LastIndexOf(@"\"));
        }

        /// <summary>
        /// The methodd for getting the file path minus the filename
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public string getFilePath (string fileName)
        {
            return fileName.Substring(0, fileName.LastIndexOf(@"\"));
        }

        /// <summary>
        /// The method for generating the file with the list of exceptions on the process layer.
        /// </summary>
        /// <param name="fileName"></param>
        public void createProcessFile (string fileName)
        {
            /// Hashtable variables
            Hashtable pageHashtable = new Hashtable();
            Hashtable processHashTable = new Hashtable();
             
            /// DataTable variables
            DataColumn column;
            DataRow row;

            /// CSV Variables
            DateTime now = DateTime.Now;
            string time = now.ToString(@"dd-MM-yyyy hh.mm.ss");
            string pathCSV = getFilePath(fileName) + @"\Process Exceptions List " + time + ".csv";
            StringBuilder sb = new StringBuilder();

            // Init DataTable Variables
            DataTable results = new DataTable();

            /// Add Type Column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = Config.TYPE_HEADER;
            results.Columns.Add(column);

            /// Add Details column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = Config.DETAILS_HEADER;
            results.Columns.Add(column);

            /// Add Process Column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = Config.PROCESS_HEADER;
            results.Columns.Add(column);

            /// Add Page Column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = Config.PAGE_HEADER;
            results.Columns.Add(column);

            XmlDocument doc = new XmlDocument();
            doc.Load(fileName);

            /// Get all the pages in the Blue Prism process and object layer
            XmlNodeList pageList = doc.GetElementsByTagName(Config.NAME_ATTRIBUTE);

            for (int k = 0; k < pageList.Count; k++)
            {
                if (pageList[k].ParentNode.Name == Config.SUBSHEET_TAG && pageList[k].ParentNode.ParentNode.ParentNode.Name == Config.PROCESS_TAG)
                {
                    pageHashtable[pageList[k].ParentNode.Attributes[Config.SUBSHEET_ID_ATTRIBUTE].Value] = pageList[k].InnerText;
                    processHashTable[pageList[k].ParentNode.Attributes[Config.SUBSHEET_ID_ATTRIBUTE].Value] = pageList[k].ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value;
                }
            }

            XmlNodeList elements = doc.GetElementsByTagName(Config.EXCEPTION_TAG);

            for (int j = 0; j < elements.Count; ++j)
            {
                bool flagMainPage = false;

                if (!(elements[j].Attributes[Config.USE_CURRENT_ATTRIBUTE] != null))
                {
                    if (elements[j].Attributes[Config.TYPE_ATTRIBUTE].Value == Config.EMPTY_STRING && elements[j].Attributes[Config.DETAIL_ATTRIBUTE].Value == Config.EMPTY_STRING)
                    {
                        continue;
                    }
                }

                foreach (XmlNode _temp in elements[j].ParentNode.ChildNodes)
                {
                    if (_temp.Name == Config.SUBSHEET_ID_ATTRIBUTE)
                    {
                        flagMainPage = true;
                        break;
                    }
                }

                if (elements[j].Attributes[Config.USE_CURRENT_ATTRIBUTE] != null)
                {
                    if (elements[j].ParentNode.Name == Config.STAGE_TAG && elements[j].ParentNode.FirstChild.Name == Config.SUBSHEET_ID_ATTRIBUTE && elements[j].ParentNode.ParentNode.ParentNode.Name == Config.PROCESS_TAG && flagMainPage)
                    {
                        row = results.NewRow();
                        row[Config.TYPE_HEADER] = Config.EXCEPTION_TYPE_PRESERVED;
                        row[Config.DETAILS_HEADER] = Config.EXCEPTION_DETAIL_PRESERVED;

                        if (elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value.Contains(processHashTable[elements[j].ParentNode.FirstChild.InnerText].ToString()))
                        {
                            row[Config.PROCESS_HEADER] = processHashTable[elements[j].ParentNode.FirstChild.InnerText];
                        }
                        else
                        {
                            /// Handling cases in which a process has been copied COMPLETELY from another object by using the Save As Option
                            /// This causes the subsheetID to be the same for both the processes even though the names are different.
                            /// This is handled by looking at the name of the process instead.
                            row[Config.PROCESS_HEADER] = elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value;
                        }

                        
                        row[Config.PAGE_HEADER] = pageHashtable[elements[j].ParentNode.FirstChild.InnerText];
                        results.Rows.Add(row);
                    }

                    /// For exceptions present on the main page.
                    if (!flagMainPage)
                    {
                        row = results.NewRow();
                        row[Config.TYPE_HEADER] = Config.EXCEPTION_TYPE_PRESERVED;
                        row[Config.DETAILS_HEADER] = Config.EXCEPTION_DETAIL_PRESERVED;
                        row[Config.PROCESS_HEADER] = elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value;
                        row[Config.PAGE_HEADER] = "Main Page";
                        results.Rows.Add(row);
                    }
                }
                else
                {
                    if (flagMainPage && elements[j].ParentNode.Name == Config.STAGE_TAG && elements[j].ParentNode.FirstChild.Name == Config.SUBSHEET_ID_ATTRIBUTE && elements[j].ParentNode.ParentNode.ParentNode.Name == Config.PROCESS_TAG)
                    {
                        row = results.NewRow();
                        row[Config.TYPE_HEADER] = elements[j].Attributes[Config.TYPE_ATTRIBUTE].Value;
                        row[Config.DETAILS_HEADER] = elements[j].Attributes[Config.DETAIL_ATTRIBUTE].Value;

                        if (elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value.Contains(processHashTable[elements[j].ParentNode.FirstChild.InnerText].ToString()))
                        {
                            row[Config.PROCESS_HEADER] = processHashTable[elements[j].ParentNode.FirstChild.InnerText];
                        }
                        else
                        {
                            /// Handling cases in which a process has been copied COMPLETELY from another object by using the Save As Option
                            /// This causes the subsheetID to be the same for both the processes even though the names are different.
                            /// This is handled by looking at the name of the process instead.
                            row[Config.PROCESS_HEADER] = elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value;
                        }

                        row[Config.PAGE_HEADER] = pageHashtable[elements[j].ParentNode.FirstChild.InnerText];
                        results.Rows.Add(row);
                    }

                    /// For exceptions present on the main page.
                    if (!flagMainPage)
                    {
                        row = results.NewRow();
                        row[Config.TYPE_HEADER] = elements[j].Attributes[Config.TYPE_ATTRIBUTE].Value;
                        row[Config.DETAILS_HEADER] = elements[j].Attributes[Config.DETAIL_ATTRIBUTE].Value;
                        row[Config.PROCESS_HEADER] = elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value;
                        row[Config.PAGE_HEADER] = "Main Page";
                        results.Rows.Add(row);
                    }
                }
            }

            /// Generate CSV
            string exceptionString;
            string value;

            foreach (DataRow dr in results.Rows)
            {
                exceptionString = String.Empty;
                value = checkEscape(dr[Config.TYPE_HEADER].ToString());
                exceptionString = exceptionString + value + ",";
                value = checkEscape(dr[Config.DETAILS_HEADER].ToString());
                exceptionString = exceptionString + value + ",";
                value = checkEscape(dr[Config.PROCESS_HEADER].ToString());
                exceptionString = exceptionString + value + ",";
                value = checkEscape(dr[Config.PAGE_HEADER].ToString());
                exceptionString = exceptionString + value;

                // Create Line
                sb.Append(exceptionString + Environment.NewLine);
            }

            using (System.IO.StreamWriter file = File.CreateText(pathCSV))
            {
                string csvHeader = string.Format("{0}, {1}, {2}, {3}", Config.TYPE_HEADER, Config.DETAILS_HEADER, Config.PROCESS_HEADER, Config.PAGE_HEADER);
                string csvRow = sb.ToString();

                file.WriteLine(csvHeader);
                file.Write(csvRow);
            }
        }

        /// <summary>
        /// The method for generating the file with the list of exceptions in the object layer.
        /// </summary>
        /// <param name="fileName"></param>
        public void createObjectFile(string fileName)
        {
            /// Hashtable variables
            Hashtable pageHashtable = new Hashtable();
            Hashtable objectHashtable = new Hashtable();

            /// DataTable variables
            DataColumn column;
            DataRow row;

            /// CSV Variables
            DateTime now = DateTime.Now;
            string time = now.ToString(@"dd-MM-yyyy hh.mm.ss");
            string pathCSV = getFilePath(fileName) + @"\Objects Exceptions List " + time + ".csv";
            StringBuilder sb = new StringBuilder();

            // Init DataTable Variables
            DataTable results = new DataTable();

            /// Add Type Column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = Config.TYPE_HEADER;
            results.Columns.Add(column);

            /// Add Details column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = Config.DETAILS_HEADER;
            results.Columns.Add(column);

            /// Add Object Column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = Config.OBJECT_HEADER;
            results.Columns.Add(column);

            /// Add Page Column
            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = Config.ACTION_HEADER;
            results.Columns.Add(column);

            XmlDocument doc = new XmlDocument();
            doc.Load(fileName);

            /// Get all the pages in the Blue Prism process and object layer
            XmlNodeList pageList = doc.GetElementsByTagName(Config.NAME_ATTRIBUTE);

            for (int k = 0; k < pageList.Count; k++)
            {
                if (pageList[k].ParentNode.Name == Config.SUBSHEET_TAG && pageList[k].ParentNode.ParentNode.ParentNode.Name == Config.OBJECT_TAG)
                {
                    pageHashtable[pageList[k].ParentNode.Attributes[Config.SUBSHEET_ID_ATTRIBUTE].Value] = pageList[k].InnerText;
                    objectHashtable[pageList[k].ParentNode.Attributes[Config.SUBSHEET_ID_ATTRIBUTE].Value] = pageList[k].ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value;
                }
            }

                XmlNodeList elements = doc.GetElementsByTagName(Config.EXCEPTION_TAG);

                for (int j = 0; j < elements.Count; ++j)
                {
                    if (!(elements[j].Attributes[Config.USE_CURRENT_ATTRIBUTE] != null))
                    {
                        if (elements[j].Attributes[Config.TYPE_ATTRIBUTE].Value == Config.EMPTY_STRING && elements[j].Attributes[Config.DETAIL_ATTRIBUTE].Value == Config.EMPTY_STRING)
                        {
                            continue;
                        }
                    }

                    if (elements[j].Attributes[Config.USE_CURRENT_ATTRIBUTE] != null)
                    {
                        if (elements[j].ParentNode.ParentNode.ParentNode.Name == Config.OBJECT_TAG)
                        {
                            row = results.NewRow();
                            row[Config.TYPE_HEADER] = Config.EXCEPTION_TYPE_PRESERVED;
                            row[Config.DETAILS_HEADER] = Config.EXCEPTION_DETAIL_PRESERVED;

                            if (elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value.Contains(objectHashtable[elements[j].ParentNode.FirstChild.InnerText].ToString()))
                            {
                                row[Config.OBJECT_HEADER] = objectHashtable[elements[j].ParentNode.FirstChild.InnerText];
                            }
                            else
                            {
                                /// Handling cases in which an object has been copied COMPLETELY from another object by using the Save As Option
                                /// This causes the subsheetID to be the same for both the objects even though the names are different.
                                /// This is handled by looking at the name of the object instead.
                                row[Config.OBJECT_HEADER] = elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value;
                            }

                            row[Config.ACTION_HEADER] = pageHashtable[elements[j].ParentNode.FirstChild.InnerText];
                            results.Rows.Add(row);
                        }
                    }
                    else
                    {
                        if (elements[j].ParentNode.ParentNode.ParentNode.Name == Config.OBJECT_TAG)
                        {
                            row = results.NewRow();
                            row[Config.TYPE_HEADER] = elements[j].Attributes[Config.TYPE_ATTRIBUTE].Value;
                            row[Config.DETAILS_HEADER] = elements[j].Attributes[Config.DETAIL_ATTRIBUTE].Value;

                            if (elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value.Contains(objectHashtable[elements[j].ParentNode.FirstChild.InnerText].ToString()))
                            {
                                row[Config.OBJECT_HEADER] = objectHashtable[elements[j].ParentNode.FirstChild.InnerText];
                            }
                            else
                            {
                                /// Handling cases in which an object has been copied COMPLETELY from another object by using the Save As Option
                                /// This causes the subsheetID to be the same for both the objects even though the names are different.
                                /// This is handled by looking at the name of the object instead.
                                row[Config.OBJECT_HEADER] = elements[j].ParentNode.ParentNode.ParentNode.Attributes[Config.NAME_ATTRIBUTE].Value;
                            }

                            row[Config.ACTION_HEADER] = pageHashtable[elements[j].ParentNode.FirstChild.InnerText];
                            results.Rows.Add(row);
                        }
                    }
                }

            /// Generate CSV
            string exceptionString;
            string value;

            foreach (DataRow dr in results.Rows)
            {
                exceptionString = String.Empty;
                value = checkEscape(dr[Config.TYPE_HEADER].ToString());
                exceptionString = exceptionString + value + ",";
                value = checkEscape(dr[Config.DETAILS_HEADER].ToString());
                exceptionString = exceptionString + value + ",";
                value = checkEscape(dr[Config.OBJECT_HEADER].ToString());
                exceptionString = exceptionString + value + ",";
                value = checkEscape(dr[Config.ACTION_HEADER].ToString());
                exceptionString = exceptionString + value;

                // Create Line
                sb.Append(exceptionString + Environment.NewLine);
            }

            using (System.IO.StreamWriter file = File.CreateText(pathCSV))
            {
                string csvHeader = string.Format("{0}, {1}, {2}, {3}", Config.TYPE_HEADER, Config.DETAILS_HEADER, Config.OBJECT_HEADER, Config.ACTION_HEADER);
                string csvRow = sb.ToString();

                file.WriteLine(csvHeader);
                file.Write(csvRow);
            }
        }

        /// <summary>
        /// Main Method
        /// </summary>
        /// <param name="args"></param>
        public static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                System.Console.WriteLine("Please enter the complete file path");
                return 1;
            }

            string fileName = args[0];

            if (File.Exists(fileName) == false)
            {
                Console.WriteLine("File Path entered is invalid");
                return 1;
            }

            /// Init object
            Program obj = new Program();

            if (obj.getFileName(fileName).Contains(".bprelease") == false)
            {
                Console.WriteLine("This application only works with BPRELEASE files");
                return 1;
            }

            // Initiation
            Console.WriteLine("Init...");
            Thread.Sleep(200);
            Console.WriteLine("Flush enviornment...");
            Thread.Sleep(200);
            Console.WriteLine("Set up Modules...");
            Thread.Sleep(200);
            Console.WriteLine("...");
            Console.WriteLine("");

            Console.WriteLine("Init Process exceptions extraction...");
            Thread.Sleep(200);
            Console.WriteLine(".");
            Thread.Sleep(200);
            Console.WriteLine("..");
            Thread.Sleep(200);
            Console.WriteLine("...");
            Thread.Sleep(200);
            obj.createProcessFile(fileName);
            Console.WriteLine("Process File Generated");
            Console.WriteLine(Environment.NewLine);

            Thread.Sleep(300);
            Console.WriteLine("Init Object exceptions extraction...");
            Thread.Sleep(200);
            Console.WriteLine(".");
            Thread.Sleep(200);
            Console.WriteLine("..");
            Thread.Sleep(200);
            Console.WriteLine("...");
            Thread.Sleep(200);
            obj.createObjectFile(fileName);
            Console.WriteLine("Object File Generated");

            Thread.Sleep(200);
            Console.WriteLine("Done");

            return 0;
        }
    }
}