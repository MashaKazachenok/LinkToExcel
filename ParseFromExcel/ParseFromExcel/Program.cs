

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
            var excel = new ExcelQueryFactory();
            excel.FileName = "Info.xlsx";
            excel.AddMapping<BaseModel>(x => x.Key, "KEY");
            excel.AddMapping<RuModel>(x => x.RuValue, "RU_VALUE");
            excel.AddMapping<EnModel>(x => x.EnValue, "EN_VALUE");

            var workflowEn = from x in excel.Worksheet<EnModel>("Sheet1")
                           select x;
            var workflowRu = from x in excel.Worksheet<RuModel>("Sheet1")
                select x;
                            
         
            JsonSerializer serializer = new JsonSerializer();
            string path;
            path = @"d:\en_json.txt";
            SerializeToJson(serializer, workflowEn.ToList(), path);
            path = @"d:\ru_json.txt";
            SerializeToJson(serializer, workflowRu.ToList(), path);

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
