using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelReportsMaker
{
    public partial class NewReportForm : Form
    {
        public NewReportForm()
        {
            InitializeComponent();
        }

        private List<string> NotIncluded { get; }

        public NewReportForm(List<string> notIncluded)
        {
            //NotIncluded = NotIncluded;
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            string[] shortNames = new string[notIncluded.Count];
            for (int i = 0; i < notIncluded.Count; i++)
                shortNames[i] = notIncluded[i].Split('\\')[notIncluded[i].Split('\\').Length -2] +"\\"+ notIncluded[i].Split('\\').Last();
            listBox1.Items.AddRange(shortNames);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            WhatToDo.Update = true;
            this.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            WhatToDo.CreateNew = true;
            this.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
