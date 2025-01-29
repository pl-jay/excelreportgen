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
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(templatePath, false))
            {
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
                            DataGeneration.InsertDataRows(worksheetPart, rowCount);
                        }

                        outputWorkbookPart.Workbook.Save();
                    }
                }
            }

            return Task.CompletedTask;
        }
    }
}
