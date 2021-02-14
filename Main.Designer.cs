
namespace ExcelReportsMaker
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Back_Btn = new System.Windows.Forms.Button();
            this.Right_Btn = new System.Windows.Forms.Button();
            this.DateLabel = new System.Windows.Forms.Label();
            this.BuildReportBtn = new System.Windows.Forms.Button();
            this.AddFolder_Btn = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(7, 86);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.Size = new System.Drawing.Size(834, 422);
            this.dataGridView1.TabIndex = 0;
            // 
            // Back_Btn
            // 
            this.Back_Btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.Back_Btn.Location = new System.Drawing.Point(111, 43);
            this.Back_Btn.Name = "Back_Btn";
            this.Back_Btn.Size = new System.Drawing.Size(79, 37);
            this.Back_Btn.TabIndex = 1;
            this.Back_Btn.Text = "Вперед";
            this.Back_Btn.UseVisualStyleBackColor = true;
            this.Back_Btn.Click += new System.EventHandler(this.Back_Btn_Click);
            // 
            // Right_Btn
            // 
            this.Right_Btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.Right_Btn.Location = new System.Drawing.Point(7, 43);
            this.Right_Btn.Name = "Right_Btn";
            this.Right_Btn.Size = new System.Drawing.Size(78, 37);
            this.Right_Btn.TabIndex = 2;
            this.Right_Btn.Text = "Назад";
            this.Right_Btn.UseVisualStyleBackColor = true;
            this.Right_Btn.Click += new System.EventHandler(this.Right_Btn_Click);
            // 
            // DateLabel
            // 
            this.DateLabel.AutoSize = true;
            this.DateLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.DateLabel.Location = new System.Drawing.Point(8, 11);
            this.DateLabel.Name = "DateLabel";
            this.DateLabel.Size = new System.Drawing.Size(52, 20);
            this.DateLabel.TabIndex = 3;
            this.DateLabel.Text = "Дата:";
            // 
            // BuildReportBtn
            // 
            this.BuildReportBtn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BuildReportBtn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.BuildReportBtn.Location = new System.Drawing.Point(454, 43);
            this.BuildReportBtn.Name = "BuildReportBtn";
            this.BuildReportBtn.Size = new System.Drawing.Size(134, 37);
            this.BuildReportBtn.TabIndex = 4;
            this.BuildReportBtn.Text = "Создать отчет";
            this.BuildReportBtn.UseVisualStyleBackColor = true;
            this.BuildReportBtn.Click += new System.EventHandler(this.BuildReportBtn_Click);
            // 
            // AddFolder_Btn
            // 
            this.AddFolder_Btn.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.AddFolder_Btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.AddFolder_Btn.Location = new System.Drawing.Point(594, 3);
            this.AddFolder_Btn.Name = "AddFolder_Btn";
            this.AddFolder_Btn.Size = new System.Drawing.Size(247, 39);
            this.AddFolder_Btn.TabIndex = 5;
            this.AddFolder_Btn.Text = "Добавить папку с файлами";
            this.AddFolder_Btn.UseVisualStyleBackColor = true;
            this.AddFolder_Btn.Click += new System.EventHandler(this.AddFolder_Btn_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.button1.Location = new System.Drawing.Point(594, 43);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(247, 37);
            this.button1.TabIndex = 6;
            this.button1.Text = "Показать последний отчет";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(242, 11);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(121, 28);
            this.comboBox1.TabIndex = 7;
            this.comboBox1.SelectedValueChanged += new System.EventHandler(this.comboBox1_SelectedValueChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(853, 520);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.AddFolder_Btn);
            this.Controls.Add(this.BuildReportBtn);
            this.Controls.Add(this.DateLabel);
            this.Controls.Add(this.Right_Btn);
            this.Controls.Add(this.Back_Btn);
            this.Controls.Add(this.dataGridView1);
            this.Name = "MainForm";
            this.Text = "Main";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button Back_Btn;
        private System.Windows.Forms.Button Right_Btn;
        private System.Windows.Forms.Label DateLabel;
        private System.Windows.Forms.Button BuildReportBtn;
        private System.Windows.Forms.Button AddFolder_Btn;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox comboBox1;
    }
}

