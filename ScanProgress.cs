using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace ExcelReportsMaker
{
    public partial class ScanProgress : Form
    {
        public ScanProgress(int max)
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
            progressBar1.Maximum = max;
            timer1.Start();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBar1.Value = ListCounter.Count;
            label1.Text = $"Идет сканирование докумена: {ListCounter.Count}/{progressBar1.Maximum}";
            if (ListCounter.Finished)
                this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var processes = from p in Process.GetProcessesByName("EXCEL")
                        select p;

            foreach (var process in processes)
            {
                process.Kill();
            }
            this.Dispose();
        }
    }
}
