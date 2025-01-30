using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReportGen
{
    public class ReportGeneration
    {
        public static Task GenerateReport(string templatePath, string outputPath, int rowCount)
        {
            try
            {
                // Ensure the output directory exists
                string directory = Path.GetDirectoryName(outputPath) ?? "";
                if (!Directory.Exists(directory) && !string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Create a new Excel file
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
                {
                    // Add a workbook part to the document
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    // Add a worksheet part to the workbook
                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    // Create Sheets collection
                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Report"
                    };
                    sheets.Append(sheet);

                    // Insert data into the new sheet
                    DataGeneration.InsertDataRows(worksheetPart, rowCount);

                    // Save the workbook
                    worksheetPart.Worksheet.Save();
                    workbookPart.Workbook.Save();
                }

                Console.WriteLine($"Report successfully generated: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while generating report: {ex.Message}");
            }

            return Task.CompletedTask;
        }
    }
}
