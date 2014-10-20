using System;
using System.Collections.Generic;
using System.Linq;
using TSEnvironmentLib_x86;
using System.IO;
using System.Xml.Linq;

namespace GetTaskSequenceVariables
{
    class Program
    {
        static void Main(string[] args)
        {
            bool _outputCsv = false;
            bool _outputTxt = false;
            bool _outputConsole = false;
            bool _outputHtml = false;
            bool _outputXml = false;
            string outputFile = "";

            foreach (string s in args)
            {

                if (s.ToLower().Contains("out"))
                {
                    outputFile = ParseOutputFile(s);
                }
                else if (s.ToLower().Contains("?") || s.Contains("help"))
                {
                    ShowUsage();
                }
                else if (s.ToLower().Contains("csv"))
                {
                    _outputCsv = true;
                }
                else if (s.ToLower().Contains("txt"))
                {
                    _outputTxt = true;
                }
                else if (s.ToLower().Contains("html"))
                {
                    _outputHtml = true;
                }
                else if (s.ToLower().Contains("xml"))
                {
                    _outputXml = true;
                }
                else if (s.ToLower().Contains("console"))
                {
                    _outputConsole = true;
                }
                
            }

            if (!(_outputConsole) && !(_outputCsv) && !(_outputHtml) && !(_outputTxt) && !(_outputXml))
            {
                _outputConsole = true;
            }

            if (outputFile.Length < 1 && !(_outputConsole))
            {
                Console.WriteLine("Output file not specified!\n\n");
                ShowUsage();
            }

            Dictionary<string, string> tsVars = BuildTsVariableDictionary();

            if (_outputConsole)
            {
                OutputToConsole(tsVars);
            }
            else if (_outputCsv)
            {
                OutputToCsv(tsVars, outputFile);
            }
            else if (_outputTxt)
            {
                OutputToTxt(tsVars, outputFile);
            }
            else if (_outputHtml)
            {
                OutputToHtml(tsVars, outputFile);
            }
            else if (_outputXml)
            {
                OutputToXml(tsVars, outputFile);
            }
        }

        public static Dictionary<string,string> BuildTsVariableDictionary()
        {
            ITSEnvClass TsEnv = null;

            try
            {
                TsEnv = new TSEnvClass();
            }
            catch
            {
                Console.WriteLine("Unable to locate Task Sequence Environment");
                Environment.Exit(1);
            }

            Dictionary<string, string> tsVars = new Dictionary<string, string>();

            object[] x = (object[])TsEnv.GetVariables();

            for (int i = 0; i < x.Length; i++)
            {
                tsVars.Add(x[i].ToString(), TsEnv[x[i].ToString()]);
            }

            return tsVars;
        }

        public static void ShowUsage()
        {
            Console.WriteLine("GetTaskSequenceVariablesDotNet\n");
            Console.WriteLine("Outputs all variables of a running Task Sequence.\n.NET must be installed in your WinPE image!");
            Console.WriteLine("\n");
            Console.WriteLine("Must include one of the following:");
            Console.WriteLine("\t/csv /out=<filepath>");
            Console.WriteLine("\t/txt /out=<filepath>");
            Console.WriteLine("\t/html /out=<filepath>");
            Console.WriteLine("\t/xml /out=<filepath>");
            Console.WriteLine("\t/console\n");
            Console.WriteLine("Examples:");
            Console.WriteLine("\tGetTaskSequenceVariablesDotNet.exe /csv /out=c:\\outputfile.csv");
            Console.WriteLine("\tGetTaskSequenceVariablesDotNet.exe /html /out=c:\\outputfile.html");
            Console.WriteLine("\tGetTaskSequenceVariablesDotNet.exe /console\n");


            Environment.Exit(1);
        }

        public static string ParseOutputFile(string s)
        {
            string[] temp = s.Split('=');
            return temp[1];
        }

        public static void OutputToCsv(Dictionary<string,string> dic, string filePath)
        {
            Console.WriteLine("Creating CSV File to: {0}", filePath);
            StreamWriter oFile = new StreamWriter(filePath, true);

            // Create File Header
            oFile.WriteLine("Variable Name,VariableValue");

            // Output Each Variable
            foreach (KeyValuePair<string, string> p in dic)
            {
                oFile.WriteLine(p.Key + "," + p.Value);
            }

            // Close the file
            oFile.Close();
            Console.WriteLine("Successfully created file: {0}", filePath);
        }

        public static void OutputToTxt(Dictionary<string, string> dic, string filePath)
        {
            Console.WriteLine("Creating TXT File to: {0}", filePath);
            StreamWriter oFile = new StreamWriter(filePath, true);
            
            // Output Each Variable
            foreach (KeyValuePair<string, string> p in dic)
            {
                oFile.WriteLine(p.Key + " : " + p.Value);
            }

            // Close the file
            oFile.Close();
            Console.WriteLine("Successfully created file: {0}", filePath);
        }

        public static void OutputToHtml(Dictionary<string, string> dic, string filePath)
        {
            Console.WriteLine("Creating HTML File to: {0}", filePath);
            StreamWriter oFile = new StreamWriter(filePath, true);

            string tsCompName = string.Join("",from i in dic
                                               where i.Key == "OSDComputerName"
                                               select i.Value);
            string tsManufac = string.Join("",from i in dic
                                                  where i.Key == "Make"
                                                  select i.Value);
            string tsModel = string.Join("",from i in dic
                                                where i.Key == "Model"
                                                select i.Value);
            string tsRam = string.Join("",from i in dic
                                                where i.Key == "Memory"
                                                select i.Value);

            oFile.WriteLine(@"<html><head><title>Task Sequence Variables</title>");
            oFile.WriteLine(@"<style type='text/css'>table{border:0;border-collapse:collapse;margin-left:50px;}tr:nth-child(even){background:#E7E7E7;}tr:nth-child(odd){background:#FFFFFF;}td{padding-left:5;padding-top:3;padding-right:5;padding-bottom:3;}#tblHeader{background:#7E7E7E;color:#FFFFFF;font-weight:bold;}</style></head>");
            oFile.WriteLine(@"<body><div><h2>Computer Information:</h2><table><tr id='tblHeader'>");
            oFile.WriteLine(@"<td>Computer Name:</td><td>" + tsCompName + "</td></tr>");
            oFile.WriteLine(@"<tr><td>Manufacturer</td><td>" + tsManufac + "</td></tr>");
            oFile.WriteLine(@"<tr><td>Model</td><td>" + tsModel + "</td></tr>");
            oFile.WriteLine(@"<tr><td>RAM</td><td>" + tsRam + " MB</td></tr>");
            oFile.WriteLine(@"</table></div><div><h2>Task Sequence Variables:</h2><table><tr id='tblHeader'><td>Variable Name</td><td>Variable Value</td></tr>");

            foreach (KeyValuePair<string, string> p in dic)
            {
                oFile.WriteLine(@"<tr><td>" + p.Key +"</td><td>" + p.Value + "</td></tr>");
            }

            oFile.WriteLine(@"</table></div></body></html>");

            Console.WriteLine("Successfully created HTML file: {0}", filePath);

        }

        public static void OutputToXml(Dictionary<string, string> dic, string filePath)
        {
            Console.WriteLine("Creating XML File to: {0}", filePath);

            var xdoc = new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                new XElement("Root",
                    dic.Select(v => new XElement("TSVariable",
                        new XElement("Name",v.Key),
                        new XElement("Value",v.Value))
                        )
                        ));

            xdoc.Save(filePath);

            Console.WriteLine("Successfully created XML file: {0}", filePath);
        }

        public static void OutputToConsole(Dictionary<string,string> dic)
        {
            foreach (KeyValuePair<string, string> p in dic)
            {
                Console.WriteLine("{0}\t:\t{1}\n", p.Key, p.Value);
            }
        }
    }
}
