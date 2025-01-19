using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelReportGenerator
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Excel Report Generator - Starting...");

            // Input arguments: template path and row count
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: dotnet run <TemplatePath> <RowCount>");
                return;
            }

            string templatePath = args[0] ?? "exceltemp.xlsx";
            int rowCount = int.Parse(args[1]) > 0 ? int.Parse(args[1]) : 10;

            // Validate the template file
            if (!File.Exists(templatePath))
            {
                Console.WriteLine("Error: Template file not found.");
                return;
            }

            string outputPath = GenerateOutputFilePath(templatePath);

            // Generate report
            try
            {
                Console.WriteLine("Generating report...");
                await Task.Run(() => GenerateReport(templatePath, outputPath, rowCount));
                Console.WriteLine($"Report generated successfully: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during report generation: {ex.Message}");
            }
        }

        private static string GenerateOutputFilePath(string templatePath)
        {
            string directory = Path.GetDirectoryName(templatePath) ?? "";
            string fileName = Path.GetFileNameWithoutExtension(templatePath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return Path.Combine(directory, $"{fileName}_{timestamp}.xlsx");
        }

        private static Task GenerateReport(string templatePath, string outputPath, int rowCount)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(templatePath, false))
            {
                // Open the template for reading
                var workbookPart = document.WorkbookPart;
                if (workbookPart == null)
                {
                    throw new InvalidOperationException("WorkbookPart is null.");
                }
                var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet == null)
                {
                    throw new InvalidOperationException("No sheets found in the template.");
                }

                // Copy the template to create the output file
                File.Copy(templatePath, outputPath, true);

                // Open the output file for modification
                using (SpreadsheetDocument outputDoc = SpreadsheetDocument.Open(outputPath, true))
                {
                    var outputWorkbookPart = outputDoc.WorkbookPart;
                    if (outputWorkbookPart != null)
                    {
                        var outputSheet = outputWorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                        if (outputSheet != null && outputSheet.Id != null)
                        {
                            var worksheetPart = (WorksheetPart)outputWorkbookPart.GetPartById(outputSheet.Id);
                            InsertDataRows(worksheetPart, rowCount);
                        }

                        // Save changes
                        outputWorkbookPart.Workbook.Save();
                    }
                }
            }

            return Task.CompletedTask;
        }

        private static void InsertDataRows(WorksheetPart worksheetPart, int rowCount)
        {
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            // Dummy data generation (Example: Replace this with dynamic generation logic)
            for (int i = 1; i <= rowCount; i++)
            {
                var newRow = new Row();
                newRow.Append(
                    CreateCell($"Customer_{i}", CellValues.String),
                    CreateCell($"customer{i}@example.com", CellValues.String),
                    CreateCell((1000 + i).ToString(), CellValues.Number),
                    CreateCell(DateTime.Now.AddDays(i).ToShortDateString(), CellValues.String),
                    CreateCell(RandomNumberGenerator.GetInt32(10, 39).ToString(), CellValues.Number)
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
