
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
        public ViewForm()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }


        private void OpenFile_click(object sender, EventArgs e)
        {
            Text = openFileDialog1.FileName;
            GetOpenFile();
           
        }

        private static void GetSaveFile()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            //   saveFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                // Code to write the stream goes here.
            }
        }

        private static void GetOpenFile()
        {

            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = "D:\\ParseFromExcel\\ParseFromExcel\\bin\\Debug";
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            //  openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            //   excel.FileName = "Info.xlsx";


            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                var excel = new ExcelQueryFactory();

                excel.FileName = openFileDialog1.FileName;


                JsonSerializer serializer = new JsonSerializer();
                string path;
                path = @"d:\en.json";

            }
        }

        private void StartConwert_Click(object sender, EventArgs e)
        {
            string excelPath = @"d:\Info.xlsx";
            string worksheetName = "Sheet2";

            string jsonDirectory = @"d:";

            var excel = new ExcelQueryFactory(excelPath);
            var columnNames = excel.GetColumnNames(worksheetName).ToList();

            excel.AddMapping<Values>(x => x.Key, columnNames[0]);

            for (int i = 1; i < columnNames.Count(); i++)
            {
                string path = String.Format(@"{0}\{1}.json", jsonDirectory, columnNames[i]);
                excel.AddMapping<Values>(x => x.Value, columnNames[i]);

                List<Values> workflows = (from x in excel.Worksheet<Values>(worksheetName)
                                          select x)
                .Where(x => x.Key != null)
                .ToList();

                SerializeToJson(workflows, path);
            }
        }

        private void SaveFile_Click(object sender, EventArgs e)
        {
            GetSaveFile();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            Text = openFileDialog1.FileName;
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
    }

}
