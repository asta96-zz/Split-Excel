using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDemo;
using OfficeOpenXml;

namespace Split_Excel
{
    public partial class Form1 : Form
    {
        internal   string FilePath { get; set; }
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Text Files",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xlsx",
                Filter = " Excel Files(.xlsx)|*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                label2.Text = openFileDialog1.SafeFileName;
                label2.Visible = true;
                FilePath = openFileDialog1.FileName;
                Preview.Visible = true;
            }
            
        }

        private void SplitButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(FilePath))
            {

            }
        }

        private void Preview_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(FilePath))
            {
                var dt = ExcelHelper.GetDataTableFromExcel(FilePath, true);
                dataGridView1.DataSource = dt;
                dataGridView1.Visible = true;
            }
            else
            {

            }
        }
    }
}
