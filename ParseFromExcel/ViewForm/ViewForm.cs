
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using LinqToExcel;
using Model;
using Newtonsoft.Json;
using System.IO;

namespace ViewForm
{
    public partial class ViewForm : Form
    {
        private string _openPath;
        private string _jsonDirectoryPath;
        private string _messageText;

        public ViewForm()
        {
            InitializeComponent();
        }

        private void openFileDialog_FileOk(object sender, CancelEventArgs e)
        {


        }


        private void OpenFile_click(object sender, EventArgs e)
        {
            GetOpenFile();
            OpenFilePathBox.Text = _openPath;

        }

        private void GetSaveFile()
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                _jsonDirectoryPath = folderBrowserDialog.SelectedPath;
              
            }
            
        }

        private void GetOpenFile()
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "D:\\";
            openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)

            {
                var excel = new ExcelQueryFactory();
                excel.FileName = openFileDialog.FileName;
                _openPath = excel.FileName; 
            }
        }

        private void StartConwert_Click(object sender, EventArgs e)
        {
            string excelPath = _openPath;
            string worksheetName = "Sheet2";

            var excel = new ExcelQueryFactory(excelPath);
            var columnNames = excel.GetColumnNames(worksheetName).ToList();

            excel.AddMapping<Values>(x => x.Key, columnNames[0]);

            for (int i = 1; i < columnNames.Count(); i++)
            {
                string path = String.Format(@"{0}\{1}.json", _jsonDirectoryPath, columnNames[i]);
                excel.AddMapping<Values>(x => x.Value, columnNames[i]);

                List<Values> workflows = (from x in excel.Worksheet<Values>(worksheetName)
                                          select x)
                .Where(x => x.Key != null)
                .ToList();
                MessageText.Text = "Ready!";
                _messageText = MessageText.Text;
                SerializeToJson(workflows, path);
            }
        }

        private void SaveFile_Click(object sender, EventArgs e)
        {
            GetSaveFile();
            SaveFilePathBox.Text = _jsonDirectoryPath;

        }

       

        private static void SerializeToJson(List<Values> workflows, string path)
        {
            Dictionary<string, string> points = new Dictionary<string, string>();

            foreach (var item in workflows)
            {
                points.Add(item.Key, item.Value);
            }

            string json = JsonConvert.SerializeObject(points, Formatting.Indented);

            using (var file = File.CreateText(path))
            {
                file.Write(json);
               
            }
        }

        private void folderBrowserDialog_HelpRequest(object sender, EventArgs e)
        {

        }

        private void ViewForm_Load(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {
           
        }
    }

}
