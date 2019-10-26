using EntityFrameworkCoreSample.Model;
using EPPlus.DataExtractor;
using Microsoft.EntityFrameworkCore;
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

            try
            {
                Console.WriteLine("Saving data into database...");
                SaveDataIntoDatabase(data);
                Console.WriteLine("Data saved");

                Console.WriteLine("");
                Console.WriteLine("Dumping data from database:");
                DumpDatabaseData();
                Console.WriteLine("");

                Console.WriteLine("Cleaning database...");
                ClearDatabase();
                Console.WriteLine("Database cleared");

                Console.WriteLine("");
                Console.WriteLine("Press enter to exit.");
                Console.ReadLine();
            }
            catch (Exception)
            {
                Console.WriteLine("An exception has been thrown, trying to clear database");
                try
                {
                    ClearDatabase();
                    Console.WriteLine("Database cleared");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error clearing database: {ex}");
                }

                throw;
            }
        }

        private static void ClearDatabase()
        {
            using (var context = new VehicleStoreContext())
            {
                var deleteRevenuesSql = $"DELETE FROM {MonthlyRevenueEntity.TableName}";
                context.Database.ExecuteSqlCommand(deleteRevenuesSql);

                var deleteBranchesSql = $"DELETE FROM {BranchEntity.TableName}";
                context.Database.ExecuteSqlCommand(deleteBranchesSql);
            }
        }

        private static void DumpDatabaseData()
        {
            using (var context = new VehicleStoreContext())
            {
                var branchesWithRevenues = context.Branches.Include(b => b.Revenues);
                foreach (var branch in branchesWithRevenues)
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
                    .WithCollectionProperty(b => b.Revenues,
                        r => r.MonthYear, 2,
                        r => r.Value, "E", "P")
                    .GetData(3, 4)
                    .ToList();

                return data;
            }
        }
    }
}
