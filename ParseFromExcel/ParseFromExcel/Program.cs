

using System;
using System.IO;
using System.Linq;
using LinqToExcel;
using Model;
using Newtonsoft.Json;

namespace ParseFromExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            // var path = @"D:\Person.xlsx";
            var excel = new ExcelQueryFactory("Info.xlsx");
            var columnNames = excel.GetColumnNames("Sheet1");
            string path;
            JsonSerializer serializer = new JsonSerializer();

            foreach(var name in columnNames)
            {  
                path = String.Format( "d:\\{0}.json", name);
                excel.AddMapping<Values>(x => x.Value, name);
                var workflow = from x in excel.Worksheet<Values>("Sheet1")
                               select x;
                SerializeToJson(serializer, workflow.ToList(), path);
            }

            Console.ReadKey();
        }

        private static void SerializeToJson(JsonSerializer serializer, Object workflow, string path)
        {

            using (StreamWriter strim = new StreamWriter(path))

            using (JsonWriter writer = new JsonTextWriter(strim))
            {
                serializer.Serialize(writer, workflow);
            }
            
            string json = JsonConvert.SerializeObject(workflow);
            Console.Write(json);
        }
    }
}
