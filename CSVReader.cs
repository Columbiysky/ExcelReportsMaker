using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportsMaker
{
    class CSVReader
    {
        static public string[,] ReadCSV(string path)
        {
            string report = string.Empty;
            using (StreamReader sr = new StreamReader(path))
            {
                report = sr.ReadToEnd();
                sr.Close();
            }

            string[] lines = report.Split('\n');
            report = report.Replace(lines[0] + "\n", "");
            lines = report.Split('\n');
            int rows = lines.Length - 1;
            int cols = lines[0].Split(',').Length;
            string[,] report_ = new string[rows, cols];
            for (int i = 0; i < rows; i++)
            {
                string[] colsTemp = lines[i].Split(',');
                for (int j = 0; j < cols; j++)
                    report_[i, j] = colsTemp[j];
            }

            return report_;
        }
    }
}
