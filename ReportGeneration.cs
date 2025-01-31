using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace ExcelReportGen
{
    public class ReportGeneration
    {
        public static Task GenerateReport(string templatePath, string outputPath, int rowCount)
        {
            try
            {
                List<string> headers = ExtractHeadersFromTemplate(templatePath); // ✅ Extract headers

                using (SpreadsheetDocument document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Generated Report"
                    };
                    sheets.Append(sheet);

                    DataGeneration.InsertDataRows(worksheetPart, rowCount, headers); // ✅ Pass extracted headers

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


        private static List<string> ExtractHeadersFromTemplate(string templatePath)
        {
            List<string> headers = new List<string>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(templatePath, false))
            {
                var workbookPart = document.WorkbookPart;
                var sheet = workbookPart?.Workbook.Descendants<Sheet>().FirstOrDefault();
                if (sheet == null || workbookPart == null)
                {
                    throw new InvalidOperationException("No sheets found in the template.");
                }

                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // ✅ Extract Headers (First Row)
                var firstRow = sheetData?.Elements<Row>().FirstOrDefault();
                if (firstRow != null)
                {
                    foreach (var cell in firstRow.Elements<Cell>())
                    {
                        string cellValue = GetCellValue(workbookPart, cell);
                        headers.Add(cellValue);
                    }
                }
            }

            return headers;
        }

        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            if (cell.CellValue == null)
                return string.Empty;

            string value = cell.CellValue.Text;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStringTable != null)
                {
                    value = sharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }

            return value;
        }

    }
}
