using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Timers;
using System.ComponentModel;

namespace ExcelReportsMaker
{
    class CurrencyNode
    {
        public string Currency { get; set; }
        public string Values { get; set; }

        public string Node { get; set; }
    }

    class BuildCSVFromExcelFiles
    {
        public BuildCSVFromExcelFiles(string[] paths) { Paths = paths; }
        public BuildCSVFromExcelFiles() { }

        private string[] Paths { get; }

        private string[] GetFilesFromFolder(string path_)
        {
            return Directory.GetFiles(path_, "*.xlsm");
        }

        public void PrintFilesInfoToCSV(string[] xlBooks)
        {
            string old = string.Empty;
            if (File.Exists(@".\IndexedFiles.csv")) { 
                using (StreamReader sr = new StreamReader(@".\IndexedFiles.csv"))
                {
                    old = sr.ReadToEnd();
                    sr.Close();
                }
            }

            using(StreamWriter sw = new StreamWriter(@".\IndexedFiles.csv"))
            {
                if(!old.Contains("CreateDate,LatestModificationDate,FileName"))
                    sw.WriteLine("CreateDate,LatestModificationDate,FileName");
                if (old != string.Empty) sw.Write(old);
                foreach(var item in xlBooks)
                    sw.WriteLine($"{File.GetCreationTime(item)},{File.GetLastWriteTime(item)},{Path.GetFileName(item)}");
            }
        }

        public string ReadOneFile(string path)
        {
            string resultStr = "";
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWbook = xlApp.Workbooks.Open(path);
            string filename = xlWbook.Name;
            Excel._Worksheet xlUSD = xlWbook.Sheets["USD"];
            Excel._Worksheet xlEUR = xlWbook.Sheets["EUR"];
            Excel._Worksheet xlGBP = xlWbook.Sheets["GBP"];
            Excel._Worksheet xlJPY = xlWbook.Sheets["JPY"];
            Excel._Worksheet xlCHF = xlWbook.Sheets["CHF"];
            Excel._Worksheet xlNZD = xlWbook.Sheets["NZD"];
            Excel._Worksheet xlCAD = xlWbook.Sheets["CAD"];
            Excel._Worksheet xlAUD = xlWbook.Sheets["AUD"];

            Excel.Range xlRangeUSD = xlUSD.Range["A3", "E19"];
            Excel.Range xlRangeEUR = xlEUR.Range["A3", "E9"];
            Excel.Range xlRangeGPB = xlGBP.Range["A3", "E10"];
            Excel.Range xlRangeJPY = xlJPY.Range["A3", "E9"];
            Excel.Range xlRangeCHF = xlCHF.Range["A3", "E10"];
            Excel.Range xlRangeNZD = xlNZD.Range["A3", "E10"];
            Excel.Range xlRangeCAD = xlCAD.Range["A3", "E10"];
            Excel.Range xlRangeAUD = xlAUD.Range["A3", "E10"];

            List<CurrencyNode> currenciesInNode = new List<CurrencyNode>();
            currenciesInNode.AddRange(ReadList(xlRangeUSD));
            currenciesInNode.AddRange(ReadList(xlRangeEUR));
            currenciesInNode.AddRange(ReadList(xlRangeGPB));
            currenciesInNode.AddRange(ReadList(xlRangeJPY));
            currenciesInNode.AddRange(ReadList(xlRangeCHF));
            currenciesInNode.AddRange(ReadList(xlRangeNZD));
            currenciesInNode.AddRange(ReadList(xlRangeCAD));
            currenciesInNode.AddRange(ReadList(xlRangeAUD));

            foreach (var item in currenciesInNode)
                resultStr += $"{System.DateTime.Now.Date.ToString("dd.MM.yyyy")},{filename.Replace(".xlsm", "")},{item.Node},{item.Currency},{item.Values}\n";

            xlWbook.Close(0);
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            return resultStr;
        }

        public void ReadSelected()
        {
            ListCounter.Count = 0;
            ListCounter.Finished = false;
            string[] xlBooks = WhatToDo.filenames;
            int length = 0;
            if (xlBooks != null)
                length = xlBooks.Length;
            Thread t = new Thread(() => new ScanProgress(length).ShowDialog());
            try
            {
                if (length > 0)
                {
                    t.IsBackground = true;
                    t.Start();

                    string resultStr = "DateOfBuild,DateOfFile,Node,Pair,W,D,H4,H1\n";
                    foreach (var i in xlBooks)
                        if (!i.Contains("~$"))
                        {
                            ListCounter.Count++;
                            resultStr += ReadOneFile(i);
                        }
                    PrintToFile(resultStr);
                    PrintFilesInfoToCSV(xlBooks);
                    ListCounter.Finished = true;
                    t.Abort();
                }
                else 
                    ListCounter.Finished = true;
            }
            catch (NullReferenceException e)
            {
                //ListCounter.Finished = true;
            }
        }

        public void Read()
        {           
            List<string> ExcelFiles = new List<string>();
            foreach (var i in Paths)
                ExcelFiles.AddRange(GetFilesFromFolder(i));

            string[] xlBooks = ExcelFiles.ToArray();

            ListCounter.Count = 0;
            ListCounter.Finished = false;
            
            new Thread(() => new ScanProgress(xlBooks.Length).ShowDialog()).Start();
            
            string resultStr = "DateOfBuild,DateOfFile,Node,Pair,W,D,H4,H1\n";
            foreach (var i in xlBooks)
                if (!i.Contains("~$"))
                {
                    ListCounter.Count++;
                    resultStr += ReadOneFile(i);
                }

            PrintToFile(resultStr);
            PrintFilesInfoToCSV(xlBooks);
            ListCounter.Finished = true;
        }

        private List<CurrencyNode> ReadList(Excel.Range range) 
        {
            List<CurrencyNode> currenciesInNode = new List<CurrencyNode>();
            for (int i = 1; i <= (range.Rows.Count); i++)
            {
                string name = string.Empty;
                string values = string.Empty;

                for (int j = 1; j <= (range.Columns.Count); j++)
                    try
                    {
                        if (j == 1)
                            name = range.Cells[i, j].Value2.ToString();
                        else if (j != range.Columns.Count)
                            values += range.Cells[i, j].Value2.ToString() + ',';
                        else
                            values += range.Cells[i, j].Value2.ToString();
                    }
                    catch { }

                if (name != string.Empty && values != string.Empty)
                    currenciesInNode.Add(new CurrencyNode() { Currency = name, Values = values, Node =  range.Worksheet.Name});
            }

            return currenciesInNode;
        }

        void PrintToFile(string resultStr)
        {
            string currDate = System.DateTime.Now.Date.ToString("dd.MM.yyyy");
            Directory.CreateDirectory(".\\reports");
            int count = Directory.GetFiles(@".\reports\", $@"*{currDate}*").Count();
            try
            {
                using (StreamReader sr = new StreamReader(@".\reports\report-" + currDate + "-" + count + ".csv"))
                {
                    string tmp = sr.ReadLine();
                    if (!tmp.Contains("DateOfBuild,DateOfFile,Node,Pair,W,D,H4,H1\n"))
                        using (StreamWriter sw = new StreamWriter(@".\reports\report-" + currDate + "-" + count + ".csv"))
                        {
                            sw.Write("DateOfBuild,DateOfFile,Node,Pair,W,D,H4,H1\n");
                            sw.Close();
                        }
                    sr.Close();
                }
            }
            catch { }

            using (StreamWriter sw = new StreamWriter(@".\reports\report-" + currDate + "-" + count + ".csv"))
            {
                sw.Write(resultStr);
                sw.Close();
            }
        }
    }
}
