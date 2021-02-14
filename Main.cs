using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace ExcelReportsMaker
{
    public partial class MainForm : Form
    {
        private string[] directories = null;
        private string[] files = null;
        private string[] shortFiles = null;
        string currDate = System.DateTime.Now.Date.ToString("dd.MM.yyyy");
        public MainForm()
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            ListCounter.Count = 0;
            ListCounter.Finished = false;
            this.MinimumSize = new Size(869, 559);
            InitializeComponent();
            CheckSettings();
            ShowLatestReport();
            UpdateOrCreateNewReport();
            if (shortFiles != null)
            {
                SetComboboxAndFilesArrays();
            }
        }

        private void SetComboboxAndFilesArrays()
        {
            comboBox1.Items.Clear();
            var dirInfo = new DirectoryInfo(@".\reports").GetFiles("report-*.csv").OrderByDescending(d => d.CreationTime).ToList();
            List<string> fullnames = new List<string>();
            foreach (var item in dirInfo)
                fullnames.Add(item.FullName);
            files = fullnames.ToArray();
            shortFiles = new string[files.Length];
            for (int i = 0; i < shortFiles.Length; i++)
                shortFiles[i] = files[i].Replace(".csv", "").Remove(0, files[i].IndexOf("report-")).Replace("report-", "");
            comboBox1.Items.AddRange(shortFiles);
            foreach (var item in comboBox1.Items)
                if (DateLabel.Text.Contains(item.ToString()))
                    comboBox1.SelectedItem = item;
        }

        private void UpdateOrCreateNewReport()
        {
            //try
            //{
            if (File.Exists(".\\IndexedFiles.csv"))
            {
                var currentInfoAboutFilesArray = CSVReader.ReadCSV(".\\IndexedFiles.csv");
                string currentInfoAboutFiles = string.Empty;
                foreach (var item in currentInfoAboutFilesArray)
                    currentInfoAboutFiles += item + ",";
                List<string> filesInDirsList = new List<string>();
                foreach (var dir in directories)
                    filesInDirsList.AddRange(Directory.GetFiles(dir, "*.xlsm"));

                List<string> NotIncludedFiles = new List<string>();
                bool notIncluded = false;
                foreach (var item in filesInDirsList)
                {
                    if (!currentInfoAboutFiles.Contains(Path.GetFileName(item)))
                    {
                        NotIncludedFiles.Add(item);
                        notIncluded = true;
                    }
                }

                if (notIncluded)
                {
                    NewReportForm nrf = new NewReportForm(NotIncludedFiles);
                    nrf.ShowDialog();
                    if (WhatToDo.CreateNew)
                    {
                        MakeCSVFromExcelFiles();
                        return;
                    }
                    else if (WhatToDo.Update)
                    {
                        ListCounter.Count = 0;
                        new Thread(() => new ScanProgress(NotIncludedFiles.Count).ShowDialog()).Start();
                        var processes = from p in Process.GetProcessesByName("EXCEL")
                                        select p;
                        if (processes.Count() == 0)
                        {
                            foreach (var item in NotIncludedFiles)
                            {                            
                                AddCSVRerpotFromExcelFile(item);
                                ListCounter.Count++;                                
                            }                            
                            processes = from p in Process.GetProcessesByName("EXCEL")
                                        select p;
                            foreach (var process in processes)
                                process.Kill();
                        }
                        else
                        {
                            MessageBox.Show("Закройте все Excel!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                        ListCounter.Finished = true;
                        
                    }
                }
            }
            //}
            //catch { }
        }

        private void ShowLatestReport()
        {
            try 
            {
                var dirInfo = new DirectoryInfo(@".\reports").GetFiles("report-*.csv").OrderByDescending(d => d.CreationTime).ToList();
                List<string> fullnames = new List<string>();
                foreach (var item in dirInfo)
                    fullnames.Add(item.FullName);
                files = fullnames.ToArray();
                shortFiles = new string[files.Length];
                for (int i = 0; i < shortFiles.Length; i++)
                    shortFiles[i] = files[i].Replace(".csv", "").Remove(0, files[i].IndexOf("report-")).Replace("report-", "");

                Print(files[0]);
                SetComboboxAndFilesArrays();

            }
            catch { }
        }

        private void Print(string path)
        {
            DateLabel.Text = "Дата: " + path.Replace(".csv", "").Remove(0,path.IndexOf("report-")).Replace("report-","");
            if (File.Exists(path))
            {
                var res = CSVReader.ReadCSV(path);

                dataGridView1.RowCount = res.GetLength(0);
                dataGridView1.ColumnCount = res.GetLength(1);
                dataGridView1.Columns[0].HeaderText = "Дата создания отчета";
                dataGridView1.Columns[1].HeaderText = "Дата создания файла";
                dataGridView1.Columns[2].HeaderText = "Лист валюты";
                dataGridView1.Columns[3].HeaderText = "Валютная пара";
                dataGridView1.Columns[4].HeaderText = "W";
                dataGridView1.Columns[5].HeaderText = "D";
                dataGridView1.Columns[6].HeaderText = "H4";
                dataGridView1.Columns[7].HeaderText = "H1";
                dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.DefaultCellStyle.Font = new Font("Microsoft Sans Serif", 15F, GraphicsUnit.Pixel);

                for (int i = 0; i < res.GetLength(0); i++)
                    for (int j = 0; j < res.GetLength(1); j++)
                    {
                        string val = res[i, j];
                        bool v = int.TryParse(val, out int ress);
                        bool digit = v;
                        if (digit && ress < 0)
                            dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.FromArgb(255, 104, 104);
                        else if (digit)
                            dataGridView1.Rows[i].Cells[j].Style.BackColor = Color.FromArgb(104, 179, 255);
                        dataGridView1.Rows[i].Cells[j].Value = val;
                    }
            }
        }

        private void CheckSettings()
        {
            if (!File.Exists(@".\PathConfig.cfg"))
            {
                MessageBox.Show("Укажите папку с Excel файлами, из которых необходимо сформировать отчет", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                string path = string.Empty;
                using (var folderBrowser = new System.Windows.Forms.FolderBrowserDialog())
                {
                    System.Windows.Forms.DialogResult res = folderBrowser.ShowDialog();
                    if (res == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowser.SelectedPath))
                        if (Directory.GetFiles(folderBrowser.SelectedPath, "*.xlsm").Length > 0)
                            path = folderBrowser.SelectedPath;
                }

                using (StreamWriter sw = new StreamWriter(@".\PathConfig.cfg"))
                {
                    sw.WriteLine("Dir0 = \"{0}\" ", path);
                    sw.Close();
                }
            }
            else
            {
                bool isFileEmpty = false;
                //List<string> dirs = new List<string>();
                using (StreamReader sr = new StreamReader(@".\PathConfig.cfg"))
                {
                    string txt = sr.ReadToEnd();
                    if (txt != "")
                    {
                        directories = ReadConfig();
                        sr.Close();
                    }
                    else 
                    {
                        isFileEmpty = true;
                        sr.Close();
                    }
                }

                if (isFileEmpty)
                {
                    MessageBox.Show("Укажите папку с Excel файлами, из которых необходимо сформировать отчет", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    string path = string.Empty;
                    using (var folderBrowser = new System.Windows.Forms.FolderBrowserDialog())
                    {
                        System.Windows.Forms.DialogResult res = folderBrowser.ShowDialog();
                        if (res == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowser.SelectedPath))
                            if (Directory.GetFiles(folderBrowser.SelectedPath, "*.xlsm").Length > 0)
                                path = folderBrowser.SelectedPath;
                    }

                    using (StreamWriter sw = new StreamWriter(@".\PathConfig.cfg"))
                    {
                        sw.WriteLine("Dir0 = \"{0}\" ", path);
                        sw.Close();
                    }
                    directories = ReadConfig();
                }
            }
        }

        private string[] ReadConfig()
        {
            List<string> dirs = new List<string>();
            using (StreamReader sr = new StreamReader(@".\PathConfig.cfg"))
            {
                while (sr.Peek() >= 0)
                {
                    string line = sr.ReadLine();
                    bool write = false;
                    string tmp = string.Empty;
                    for (int i = 0; i < line.Length; i++)
                    {
                        if (line[i] == '\"' && write == false)
                            write = true;
                        else if (line[i] == '\"' && write == true)
                            write = false;
                        if (write && line[i] != '\"')
                            tmp += line[i];
                    }
                    line = tmp;
                    dirs.Add(line);
                }
            }
            return dirs.ToArray();
        }

        private void BuildReportBtn_Click(object sender, EventArgs e)
        {
            ListCounter.Count = 0;
            ListCounter.Finished = false;
            MakeCSVFromExcelFiles();
            SetComboboxAndFilesArrays();
        }

        private void AddCSVRerpotFromExcelFile(string path)
        {
            
            
                var dirInfo = new DirectoryInfo(@".\reports").GetFiles("report-*.csv").OrderByDescending(d => d.CreationTime).ToList();
                if (dirInfo.Count > 0)
                {
                    List<string> fullnames = new List<string>();
                    foreach (var item in dirInfo)
                        fullnames.Add(item.FullName);
                    files = fullnames.ToArray();
                }

                string old = string.Empty;
                using(StreamReader sr = new StreamReader(files[0]))
                {
                    old = sr.ReadToEnd();
                    sr.Close();
                }

                using (StreamWriter sw = new StreamWriter(files[0]))
                {
                    sw.Write(old);
                    BuildCSVFromExcelFiles CSVBuilder = new BuildCSVFromExcelFiles();
                    var result = CSVBuilder.ReadOneFile(path);
                    sw.Write(result);
                }

                string tmp = string.Empty;
                using (StreamReader sr = new StreamReader(@".\IndexedFiles.csv")) 
                {
                    tmp = sr.ReadToEnd();
                    sr.Close();
                }

                using (StreamWriter sw = new StreamWriter(@".\IndexedFiles.csv"))
                {
                    sw.Write(tmp);
                    sw.WriteLine($"{File.GetCreationTime(path)},{File.GetLastWriteTime(path)},{Path.GetFileName(path)}");
                    sw.Close();
                }

                Print(files[0]);
            
        }

        private void MakeCSVFromExcelFiles()
        {
            currDate = System.DateTime.Now.Date.ToString("dd.MM.yyyy");
            var processes = from p in Process.GetProcessesByName("EXCEL") select p;

            if (processes.Count() == 0)
            {
                directories = ReadConfig();
                SelectFilesForm sff = new SelectFilesForm(directories);
                sff.ShowDialog();
                BuildCSVFromExcelFiles mkCSV = new BuildCSVFromExcelFiles();
                mkCSV.ReadSelected();

                while (!ListCounter.Finished){}
                int count = Directory.GetFiles(@".\reports\", $@"*{currDate}*").Count() - 1;
                Print(@".\reports\report-" + currDate + "-" + count + ".csv");

                processes = from p in Process.GetProcessesByName("EXCEL")
                            select p;
                foreach (var process in processes)
                    process.Kill();
            }
            else
            {
                MessageBox.Show("Закройте все Excel!!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void Back_Btn_Click(object sender, EventArgs e)
        {
            int i = Array.FindIndex(files, val => val.Contains(DateLabel.Text.Replace("Дата: ", "")));
            if (i == 0)
                MessageBox.Show("Это самый новый отчет.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                Print(files[i - 1]);
                SetComboboxAndFilesArrays();
            }
        }

        private void Right_Btn_Click(object sender, EventArgs e)
        {
            int i = Array.FindIndex(files, val => val.Contains(DateLabel.Text.Replace("Дата: ", "")));
            if (i == files.Length - 1)
                MessageBox.Show("Это самый старый отчет.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else 
            {
                Print(files[i + 1]);
                SetComboboxAndFilesArrays();
            }
        }

        private void AddFolder_Btn_Click(object sender, EventArgs e)
        {
            string path = string.Empty;
            int lines = 0;
            string old = string.Empty;
            using (var folderBrowser = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult res = folderBrowser.ShowDialog();
                if (res == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(folderBrowser.SelectedPath))
                    if (Directory.GetFiles(folderBrowser.SelectedPath, "*.xlsm").Length > 0)
                        path = folderBrowser.SelectedPath;
            }

            using (StreamReader sr = new StreamReader(@".\PathConfig.cfg"))
            {
                old = sr.ReadToEnd();
                lines = old.Trim().Split('\n').Length;
                sr.Close();
            }

            using (StreamWriter sw = new StreamWriter(@".\PathConfig.cfg"))
            {
                sw.Write(old);
                if(path!= string.Empty)
                    sw.WriteLine($"Dir{lines} = \"{path}\" ");
                sw.Close();
            }
            directories = ReadConfig();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ShowLatestReport();
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            int i = Array.FindIndex(files, val => val.Contains(comboBox1.SelectedItem.ToString()));
            Print(files[i]);
        }
    }
}
