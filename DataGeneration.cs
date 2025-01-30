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
                    CreateCell(_faker.Random.Int(0,99999).ToString(), CellValues.Number),
                    CreateCell(_faker.Name.FullName(), CellValues.String),
                    CreateCell(_faker.Phone.PhoneNumberFormat(1), CellValues.Number),
                    CreateCell(_faker.Date.Past(2).ToShortDateString(), CellValues.String),
                    CreateCell(_faker.Random.Int(10, 39).ToString(), CellValues.Number)
                );

                sheetData?.AppendChild(newRow);
            }
        }

        private static Cell CreateCell(string value, CellValues dataType)
        {
            var cell = new Cell();

            if (dataType == CellValues.Number)
            {
                // Ensure numeric values are properly formatted
                if (double.TryParse(value, out double numericValue))
                {
                    cell.CellValue = new CellValue(numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                }
                else
                {
                    // Fallback to string if parsing fails
                    cell.CellValue = new CellValue(value);
                    cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
            }
            else
            {
                // Handle string values correctly
                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }

            return cell;
        }

    }
}
