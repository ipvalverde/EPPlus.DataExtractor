using EntityFrameworkCoreSample.Model;
using EPPlus.DataExtractor;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EntityFrameworkCoreSample
{
    class Program
    {
        static void Main(string[] args)
        {
            var spreadsheetPath = @".\Resource\EntityFrameworkSampleData.xlsx";

            Console.WriteLine($"Loading data from spreadsheet in '{spreadsheetPath}'.");
            var data = LoadDataFromSpreadsheet(spreadsheetPath);
            Console.WriteLine($"{data.Count()} records loaded.");

            Console.WriteLine("Saving data into database...");
            SaveDataIntoDatabase(data);
            Console.WriteLine("Data saved");

            Console.WriteLine("");
            Console.WriteLine("Dumping data from database:");
            DumpDatabaseData();
            Console.WriteLine("");
            Console.WriteLine("Press enter to exit.");
            Console.ReadLine();
        }

        public static void DumpDatabaseData()
        {
            using (var context = new VehicleStoreContext())
            {
                foreach(var branch in context.Branches)
                {
                    Console.WriteLine(branch);
                }
            }
        }

        private static void SaveDataIntoDatabase(IEnumerable<BranchEntity> data)
        {
            using (var context = new VehicleStoreContext())
            {
                foreach (var branch in data)
                {
                    context.Branches.Add(branch);
                }
                context.SaveChanges();
            }
        }

        private static IEnumerable<BranchEntity> LoadDataFromSpreadsheet(string spreadsheetPath)
        {
            var fileInfo = new FileInfo(spreadsheetPath);
            using (var package = new ExcelPackage(fileInfo))
            {
                var data = package.Workbook.Worksheets["Sheet1"]
                    .Extract<BranchEntity>()
                    .WithProperty(b => b.Id, "A")
                    .WithProperty(b => b.Name, "B")
                    .WithProperty(b => b.Location, "C")
                    .WithProperty(b => b.Phone, "D")
                    //.WithCollectionProperty(b => b.Revenues,
                    //    r => r.MonthYear, 2,
                    //    r => r.Value, "E", "P")
                    .GetData(3, 4)
                    .ToList();

                return data;
            }
        }
    }
}
