using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace achievementComputing
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
		//ExcelPackage EPexcel;
		//ExcelWorksheet sheet;
		achievementComputing_class achievementComputing;
		private void button_file_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            dataGridView1.DataSource = null;
            listBox_sheets.Items.Clear();
            button_compute.Enabled = false;
            openFileDialog1.Filter = "Excel Workbook|*.xlsx|Excel 97-2003 Workbook|*.xls";//"*.xlsx|*.xls";
            if (openFileDialog1.ShowDialog()==DialogResult.OK)
            {
                textBox_excelfile.Text = openFileDialog1.FileName;
                //FileInfo excelfile = new FileInfo(textBox_excelfile.Text);
                //EPexcel = new ExcelPackage(excelfile);
                achievementComputing = new achievementComputing_class(textBox_excelfile.Text);
				listBox_sheets.Items.Clear();
                //for (int i = 1; i <= EPexcel.Workbook.Worksheets.Count; i++)
                //{
                //    listBox_sheets.Items.Add(EPexcel.Workbook.Worksheets[i].Name);
                //}
                List<string> list = achievementComputing.getSheetsNames();
                foreach(string s in list)
				{
                    listBox_sheets.Items.Add(s);
				}
            }
            this.Enabled = true;
    }

        private void listBox_sheets_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Enabled = false;
            dataGridView1.DataSource =
                achievementComputing.getDataTableFromSheet(listBox_sheets.SelectedItem.ToString());
            this.Enabled = true;
            button_compute.Enabled = true;
        }

        private void button_compute_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            achievementComputing.computeAchievement();
            dataGridView1.DataSource = null;
                //achievementComputing.getDataTableFromSheet(listBox_sheets.SelectedItem.ToString());
            this.Enabled = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
			//MessageBox.Show("adgg", "aaa", MessageBoxButtons.OK,MessageBoxIcon.Error);
			////////测试用
			//achievementComputing = new achievementComputing_class();
			//////////

		}
	}
}
