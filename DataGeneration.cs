using Bogus;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGen
{
    public class DataGeneration
    {
        private static readonly Faker _faker = new();

        public static void InsertDataRows(WorksheetPart worksheetPart, int rowCount)
        {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            for (int i = 1; i <= rowCount; i++)
            {
                var newRow = new Row();
                newRow.Append(
                    CreateCell(_faker.Random.Int(0 - 10000).ToString(), CellValues.String),
                    CreateCell(_faker.Name.FullName(), CellValues.String),
                    CreateCell(_faker.Phone.PhoneNumber(), CellValues.Number),
                    CreateCell(_faker.Date.Past(2).ToShortDateString(), CellValues.String),
                    CreateCell(_faker.Random.Int(10, 39).ToString(), CellValues.Number)
                );

                sheetData?.AppendChild(newRow);
            }
        }

        private static Cell CreateCell(string value, CellValues dataType)
        {
            return new Cell
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
    }
}
