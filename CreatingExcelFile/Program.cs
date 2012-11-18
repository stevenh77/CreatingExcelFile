using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CreatingExcelFile
{
    class Program
    {
        static void Main(string[] args)
        {
            // Make a copy of the template file.
            File.Copy(@"blank.xlsx", @"generated.xlsx", true);

            // Open the copied template workbook. 
            using (SpreadsheetDocument myWorkbook = SpreadsheetDocument.Open(@"generated.xlsx", true))
            {
                // Access the main Workbook part, which contains all references.
                WorkbookPart workbookPart = myWorkbook.WorkbookPart;

                // Get the first worksheet. 
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(2);

                // The SheetData object will contain all the data.
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // output our headers
                var data = new[] { "ID", "TRADE", "PRICE"};
                var row = CreateNewRow(1, data);

                sheetData.AppendChild(row);

                // Begining Row pointer                       
                int rowIndex = 2;

                // For each item in the database, add a Row to SheetData.
                foreach (var trade in GetTradeData())
                {
                    // the data written in each row
                    data = new [] {trade.Id.ToString(), trade.Asset, trade.Price.ToString()};
                    
                    // create the Excel row with our data
                    row = CreateNewRow(rowIndex, data );

                    // Append Row to SheetData
                    sheetData.AppendChild(row);

                    // increase row pointer
                    rowIndex++;                
                }

                // save
                worksheetPart.Worksheet.Save();
            }
        }

        private static Row CreateNewRow(int rowIndex, params string[] data)
        {
            // New Row
            Row row = new Row { RowIndex = (UInt32)rowIndex };

            for (int i = 0; i < data.Length; i++)
            {
                // A = 65 for the first column, B = 66, C = 67...
                string column = ((char) (65 + i)).ToString();

                // New Cell
                Cell cell = new Cell
                                {
                                    DataType = CellValues.InlineString,
                                    CellReference = column + rowIndex
                                };

                // Create Text object
                Text t = new Text {Text = data[i]};

                // Append Text to InlineString object
                InlineString inlineString = new InlineString();
                inlineString.AppendChild(t);

                // Append InlineString to Cell
                cell.AppendChild(inlineString);

                // Append Cell to Row
                row.AppendChild(cell);
            }
            return row;
        }

        public static IList<Trade> GetTradeData()
        {
            var list = new List<Trade>(6)
                           {
                               new Trade() {Id = Guid.NewGuid(), Asset = "APPLE", Price = (decimal) 33.89},
                               new Trade() {Id = Guid.NewGuid(), Asset = "BMW", Price = (decimal) 1.23},
                               new Trade() {Id = Guid.NewGuid(), Asset = "CAPCOM", Price = (decimal) 87.46},
                               new Trade() {Id = Guid.NewGuid(), Asset = "FUSE", Price = (decimal) 4.24},
                               new Trade() {Id = Guid.NewGuid(), Asset = "IBM", Price = (decimal) 103.66},
                               new Trade() {Id = Guid.NewGuid(), Asset = "MICROSOFT", Price = (decimal) 45.55}
                           };
            return list;
        }
    }

    public class Trade
    {
        public Guid Id { get; set; }
        public string Asset { get; set; }
        public decimal Price { get; set; }
    }
}