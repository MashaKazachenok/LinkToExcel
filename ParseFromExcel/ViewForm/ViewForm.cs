
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Model;
using LinqToExcel;
using Newtonsoft.Json;

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


        private void button2_Click(object sender, EventArgs e)
        {
            GetOpenFile();
            GetSaveFile();
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
    }

}
