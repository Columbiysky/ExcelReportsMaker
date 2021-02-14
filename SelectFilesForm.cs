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

namespace ExcelReportsMaker
{
    public partial class SelectFilesForm : Form
    {
        public SelectFilesForm()
        {
            InitializeComponent();
        }
        private string[] Directories { get; }

        public SelectFilesForm(string[] dirs)
        {
            this.StartPosition = FormStartPosition.CenterScreen;
            Directories = dirs;
            InitializeComponent();
            FillTreeView();
        }

        private void FillTreeView()
        {
            List<string> dirNames = new List<string>();
            Dictionary<string, string> insideFiles = new Dictionary<string, string>();
            for(int i =0;i< Directories.Length; i++)
            {
                dirNames.Add(Path.GetFullPath(Directories[i]));
                var files = Directory.GetFiles(Directories[i], "*.xlsm");
                for (int j = 0; j < files.Length; j++)
                    insideFiles.Add(files[j], Directories[i]);
            }

            TreeNode[] dirs = new TreeNode[dirNames.Count];
            for (int i = 0; i < dirs.Length; i++)
                dirs[i] = new TreeNode(dirNames[i].Split('\\').Last());
            treeView1.Nodes.AddRange(dirs);
            
            for(int i =0; i < treeView1.Nodes.Count; i++)
            {
                foreach(var key in insideFiles.Keys)
                {
                    string val = string.Empty;
                    insideFiles.TryGetValue(key, out val);
                    if (treeView1.Nodes[i].Text == val.Split('\\').Last())
                    {
                        treeView1.Nodes[i].Nodes.Add(key.Split('\\').Last());
                    }
                }
                treeView1.Nodes[i].Expand();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (TreeNode n in treeView1.Nodes)
            {
                n.Checked = true;
                CheckChildren(n, true);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (TreeNode n in treeView1.Nodes)
            {
                n.Checked = false;
                CheckChildren(n, false);
            }
        }

        private void CheckChildren(TreeNode rootNode, bool isChecked)
        {
            foreach (TreeNode node in rootNode.Nodes)
            {
                CheckChildren(node, isChecked);
                node.Checked = isChecked;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string dirName = string.Empty;
            string fileName = string.Empty;
            List<string> filePaths = new List<string>();
            foreach (TreeNode dirNode in treeView1.Nodes)
            {
                
                dirName = dirNode.Text;
                foreach (TreeNode fileNode in dirNode.Nodes)
                {
                    if (fileNode.Checked)
                    {
                        int index = 0;
                        for (int i = 0; i < Directories.Length; i++)
                            if (Directories[i].Contains(dirName))
                            {
                                index = i;
                                break;
                            }
                        var dirFiles = Directory.GetFiles(Directories[index]);
                        foreach (var i in dirFiles)
                            if (i.Contains(fileNode.Text))
                                filePaths.Add(Directories[index]+"\\"+fileNode.Text);
                    }
                }
            }

            WhatToDo.filenames = filePaths.ToArray();
            this.Dispose();
        }

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            CheckChildren(e.Node, e.Node.Checked);
        }
    }
}
