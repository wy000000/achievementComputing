
namespace achievementComputing
{
    partial class Form1
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
			this.label1 = new System.Windows.Forms.Label();
			this.textBox_excelfile = new System.Windows.Forms.TextBox();
			this.button_file = new System.Windows.Forms.Button();
			this.listBox_sheets = new System.Windows.Forms.ListBox();
			this.button_compute = new System.Windows.Forms.Button();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
			this.SuspendLayout();
			// 
			// dataGridView1
			// 
			this.dataGridView1.AllowUserToAddRows = false;
			this.dataGridView1.AllowUserToDeleteRows = false;
			this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView1.Location = new System.Drawing.Point(0, 1);
			this.dataGridView1.Name = "dataGridView1";
			this.dataGridView1.ReadOnly = true;
			this.dataGridView1.RowHeadersWidth = 51;
			this.dataGridView1.RowTemplate.Height = 27;
			this.dataGridView1.Size = new System.Drawing.Size(1170, 566);
			this.dataGridView1.TabIndex = 0;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(12, 597);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(77, 15);
			this.label1.TabIndex = 1;
			this.label1.Text = "excel文件";
			// 
			// textBox_excelfile
			// 
			this.textBox_excelfile.Location = new System.Drawing.Point(15, 627);
			this.textBox_excelfile.Name = "textBox_excelfile";
			this.textBox_excelfile.ReadOnly = true;
			this.textBox_excelfile.Size = new System.Drawing.Size(519, 25);
			this.textBox_excelfile.TabIndex = 2;
			// 
			// button_file
			// 
			this.button_file.Location = new System.Drawing.Point(104, 585);
			this.button_file.Name = "button_file";
			this.button_file.Size = new System.Drawing.Size(75, 27);
			this.button_file.TabIndex = 3;
			this.button_file.Text = "打开";
			this.button_file.UseVisualStyleBackColor = true;
			this.button_file.Click += new System.EventHandler(this.button_file_Click);
			// 
			// listBox_sheets
			// 
			this.listBox_sheets.FormattingEnabled = true;
			this.listBox_sheets.ItemHeight = 15;
			this.listBox_sheets.Location = new System.Drawing.Point(595, 573);
			this.listBox_sheets.Name = "listBox_sheets";
			this.listBox_sheets.Size = new System.Drawing.Size(276, 154);
			this.listBox_sheets.TabIndex = 4;
			this.listBox_sheets.SelectedIndexChanged += new System.EventHandler(this.listBox_sheets_SelectedIndexChanged);
			// 
			// button_compute
			// 
			this.button_compute.Enabled = false;
			this.button_compute.Location = new System.Drawing.Point(899, 620);
			this.button_compute.Name = "button_compute";
			this.button_compute.Size = new System.Drawing.Size(98, 29);
			this.button_compute.TabIndex = 5;
			this.button_compute.Text = "计算并保存";
			this.button_compute.UseVisualStyleBackColor = true;
			this.button_compute.Click += new System.EventHandler(this.button_compute_Click);
			// 
			// openFileDialog1
			// 
			this.openFileDialog1.FileName = "openFileDialog1";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1169, 736);
			this.Controls.Add(this.button_compute);
			this.Controls.Add(this.listBox_sheets);
			this.Controls.Add(this.button_file);
			this.Controls.Add(this.textBox_excelfile);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dataGridView1);
			this.Name = "Form1";
			this.Text = "达成度计算";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_excelfile;
        private System.Windows.Forms.Button button_file;
        private System.Windows.Forms.ListBox listBox_sheets;
        private System.Windows.Forms.Button button_compute;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}