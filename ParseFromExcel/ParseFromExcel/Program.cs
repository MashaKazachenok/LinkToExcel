using System;
using System.IO;
using System.Linq;
using LinqToExcel;
using Model;
using Newtonsoft.Json;
using System.Collections.Generic;
using Microsoft.SqlServer.Server;

namespace ParseFromExcel
{
    class Program
    {
        static void Main(string[] args)
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

            Console.ReadKey();
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