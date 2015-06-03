

using System;
using System.IO;
using System.Linq;
using LinqToExcel;
using Model;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace ParseFromExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // var path = @"D:\Person.xlsx";
            var excel = new ExcelQueryFactory("Info.xlsx");
            var columnNames = excel.GetColumnNames("Sheet1").ToList();
            string path;
            JsonSerializer serializer = new JsonSerializer();
            excel.AddMapping<Values>(x => x.Key, columnNames[0]);

            for (int i = 1; i < columnNames.Count(); i++)
            {
                path = String.Format("d:\\{0}.json", columnNames[i]);
                excel.AddMapping<Values>(x => x.Value, columnNames[i]);
                List<Values> workflows = (from x in excel.Worksheet<Values>("Sheet1")
                                          select x).ToList();

                SerializeToJson(serializer, workflows, path);
            }

            Console.ReadKey();
        }

        private static void SerializeToJson(JsonSerializer serializer, List<Values> workflows, string path)
        {
            Dictionary<string, string> points = new Dictionary<string, string>();

            foreach (var item in workflows)
            {
                points.Add(item.Key, item.Value);
            }

            using (StreamWriter strim = new StreamWriter(path))
            {

                using (JsonWriter writer = new JsonTextWriter(strim))
                {

                    string json = JsonConvert.SerializeObject(points);
                    serializer.Serialize(writer, json);

                }
            }
        }
    }
}
