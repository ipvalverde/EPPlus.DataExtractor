using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace EPPlus.DataExtractor.Tests
{
    [TestClass]
    public class EPPlusExtensionsTests
    {
        public class SimpleRowData
        {
            public string CarName { get; set; }

            public double Value { get; set; }

            public DateTime CreationDate { get; set; }
        }

        public class RowDataWithColumnBeingRow
        {
            public string Name { get; set; }

            public int Age { get; set; }

            public List<ColumnData> MoneyData { get; set; }
        }

        public class ColumnData
        {
            public double ReceivedMoney { get; set; }

            public DateTime Date { get; set; }
        }

        private static FileInfo GetSpreadsheetFileInfo()
        {
            var fileInfo = new FileInfo(@"spreadsheets\WorkbookTest.xlsx");
            if (!fileInfo.Exists)
            {
                Assert.Inconclusive("The spreadsheet file could not be found.");
                return null;
            }

            return fileInfo;
        }

        [TestMethod]
        public void ExtractSimpleData()
        {
            var fileInfo = GetSpreadsheetFileInfo();
            using (var package = new ExcelPackage(fileInfo))
            {
                var cars = package.Workbook.Worksheets["MainWorksheet"]
                    .Extract<SimpleRowData>()
                    .WithProperty(p => p.CarName, "B")
                    .WithProperty(p => p.Value, "C")
                    .WithProperty(p => p.CreationDate, "D")
                    .GetData(4, 6)
                    .ToList();

                Assert.AreEqual(3, cars.Count);

                Assert.IsTrue(cars.Any(i =>
                    i.CarName == "Pegeut 203" && i.Value == 20000 && i.CreationDate == new DateTime(2017, 07, 01)));

                Assert.IsTrue(cars.Any(i =>
                    i.CarName == "i30" && i.Value == 21000 && i.CreationDate == new DateTime(2017, 04, 15)));

                Assert.IsTrue(cars.Any(i =>
                    i.CarName == "Etios" && i.Value == 17575 && i.CreationDate == new DateTime(2015, 07, 21)));
            }
        }

        [TestMethod]
        public void ExtractDataTransformingColumnsIntoRows()
        {
            var fileInfo = GetSpreadsheetFileInfo();
            using (var package = new ExcelPackage(fileInfo))
            {
                var data = package.Workbook.Worksheets["MainWorksheet"]
                    .Extract<RowDataWithColumnBeingRow>()
                    .WithProperty(p => p.Name, "F")
                    .WithProperty(p => p.Age, "G")
                    //.WithProperty(p => p.CreationDate, "D")
                    .GetData(2, 4)
                    .ToList();

                Assert.AreEqual(3, data.Count);
            }
        }
    }
}