using ExcelDemo;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Split_Excel
{
    public partial class Form1 : Form
    {
        internal string FilePath { get; set; }
        internal string OutFilePath { get; set; }
        internal int defaultRecordCount = 50000;
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            progressBar1.Visible = false;
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
            }

        }

        private async void SplitButton_ClickAsync(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(FilePath))
            {
                var _filepath = new FileInfo(FilePath);
                var outPutPath = _filepath.Directory + "\\out";
                _ = Directory.CreateDirectory(outPutPath);
                OutFilePath = outPutPath;
                string _tempName = _filepath.Name.Substring(0, _filepath.Name.IndexOf("."));
                DataTable dt = ExcelHelper.GetDataTableFromExcel(FilePath, true);
                int fileCount = 1;
                int RecordCountPerFile;
                RecordCountPerFile = int.TryParse(textBox1.Text, out RecordCountPerFile) ? RecordCountPerFile : defaultRecordCount;
                int takeCount = RecordCountPerFile;
                backgroundWorker1.WorkerReportsProgress = true;
                backgroundWorker1.RunWorkerAsync();
                progressBar1.Show();               
                textBox2.Show();
                textBox2.Text = ".... in progress";
                //if(true)
                //{
                //    string tempPath = Path.Combine(outPutPath, string.Concat(_tempName, "-", fileCount.ToString(), ".xlsx"));
                //    splitByGroup(dt, tempPath);
                //    return;
                //}
                for (int skipCount = 0; skipCount < dt.Rows.Count; skipCount += RecordCountPerFile)
                {
                    string tempPath = Path.Combine(outPutPath, string.Concat(_tempName, "-", fileCount.ToString(), ".xlsx"));
                    DataTable _table = dt.AsEnumerable().Skip(skipCount).Take(takeCount).CopyToDataTable();
                    await ExcelHelper.SaveExcelFile( new FileInfo(tempPath), true,"Sheet1", _table);
                    fileCount++;
                }

            }
        }

        private async void splitByGroup(DataTable dt, string tempPath, string groupByColName)
        {
            var GroupedByName = dt.AsEnumerable()
                .GroupBy(r => new { Col1 = r[groupByColName] })
                .Select(g => new { value = g.Key, ColumnValues = g.CopyToDataTable() })
                .ToList();
                for (int i=0;i<GroupedByName.Count;i++)
            {
                string name = GroupedByName[i].value.Col1.ToString();
                var groupedTable = (DataTable)GroupedByName[i].ColumnValues;
                string sheetName = name.Length < 31 ? name : string.Concat("sheet",$"-${i}");
                if (i == 0)
                {                   
                    await ExcelHelper.SaveExcelFile(new FileInfo(tempPath), true, sheetName, groupedTable);
                }
                else
                {
                    await ExcelHelper.SaveExcelFile(new FileInfo(tempPath), false, sheetName, groupedTable);

                }

            }
                //.Select(g => g.OrderBy(r => r["Rep"]).First())
                
            //dt.AsEnumerable().GroupBy(x=>x.)
        }
        

        private void Preview_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(FilePath))
            {
                var dt = ExcelHelper.GetDataTableFromExcel(FilePath, true);
                dataGridView1.DataSource = dt;
                dataGridView1.Visible = true;
                List<string> colmOptions = new List<string>();
              var getColmName=  dt.AsEnumerable().Select(g=>g.Table.Columns).ToList().FirstOrDefault();
                foreach(var x in getColmName)
                {
                    colmOptions.Add(x.ToString());
                }
                comboBox1.Items.AddRange(colmOptions.ToArray());
            }
            else
            {

            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            for (int i = 1; i <= 100; i++)
            {
                // Wait 50 milliseconds.  
                Thread.Sleep(40);
                // Report progress.  
                backgroundWorker1.ReportProgress(i);
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar   
            progressBar1.Value = e.ProgressPercentage;
            progressBar1.Show();
            // Set the text.  
            label4.Text = e.ProgressPercentage.ToString() + "%";
            if (label4.Text == "100%")
            {

                textBox2.Text = "File Split Completed Successfully" + Environment.NewLine + "Path: " + OutFilePath.ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var _filepath = new FileInfo(FilePath);
            var outPutPath = _filepath.Directory + "\\out";
            _ = Directory.CreateDirectory(outPutPath);
            OutFilePath = outPutPath;
            string _tempName = _filepath.Name.Substring(0, _filepath.Name.IndexOf("."));
            DataTable dt = ExcelHelper.GetDataTableFromExcel(FilePath, true);
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.RunWorkerAsync();
            progressBar1.Show();
            textBox2.Show();
            textBox2.Text = ".... in progress";
            if (true)
            {
                string tempPath = Path.Combine(outPutPath, string.Concat(_tempName, ".xlsx"));
                splitByGroup(dt, tempPath,comboBox1.SelectedItem.ToString());
                return;
            }
        }
    }
}
