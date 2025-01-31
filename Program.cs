using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExcelReportGen;
using Microsoft.Extensions.Logging;

namespace ExcelReportGenerator
{
    class Program
    {
        private static ILogger<Program>? _logger;

        static async Task Main(string[] args)
        {
            ConfigureLogging();

            _logger.LogInformation("Excel Report Generator - Starting...");

            // Input validation: Ensure there are at least two arguments and that they form valid pairs
            if (args.Length < 2 || args.Length % 2 != 0)
            {
                _logger.LogError("Usage: dotnet run <TemplatePath1> <RowCount1> <TemplatePath2> <RowCount2> ...");
                return;
            }

            // Parse input arguments into pairs of <templatePath, rowCount>
            var reportRequests = ParseInputArgs(args);
            if (!reportRequests.Any())
            {
                _logger.LogError("No valid input pairs provided. Please check the paths and row counts.");
                return;
            }

            try
            {
                // Generate reports concurrently
                _logger.LogInformation("Starting asynchronous report generation...");
                var tasks = reportRequests.Select(request => GenerateReportAsync(request.TemplatePath, request.RowCount));

                // Wait for all reports to be generated
                await Task.WhenAll(tasks);
                _logger.LogInformation("All reports generated successfully.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred during report generation.");
            }
        }

        private static void ConfigureLogging()
        {
            using var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder
                    .AddConsole()
                    .AddDebug()
                    .SetMinimumLevel(LogLevel.Information);
            });

            _logger = loggerFactory.CreateLogger<Program>();
        }

        private static List<ReportRequest> ParseInputArgs(string[] args)
        {
            var reportRequests = new List<ReportRequest>();

            // Edge case: No arguments provided
            if (args.Length == 0)
            {
                _logger.LogError("No arguments provided. Please specify at least one template path and row count.");
                return reportRequests;
            }

            string? currentTemplatePath = null;

            foreach (var arg in args)
            {
                // Check if argument is a valid file path
                if (File.Exists(arg))
                {
                    // Store the current template path
                    currentTemplatePath = arg;
                }
                else if (int.TryParse(arg, out int rowCount) && rowCount > 0)
                {
                    // If a valid row count is found, associate it with the previous template path
                    if (currentTemplatePath != null)
                    {
                        reportRequests.Add(new ReportRequest(currentTemplatePath, rowCount));
                        currentTemplatePath = null; 
                    }
                    else
                    {
                        _logger.LogWarning("Row count {RowCount} provided without a preceding template path. Skipping.", rowCount);
                    }
                }
                else
                {
                    _logger.LogWarning("Invalid argument detected: {Argument}. Skipping.", arg);
                }
            }

            if (currentTemplatePath != null)
            {
                _logger.LogWarning("Template path {TemplatePath} does not have an associated row count. Skipping.", currentTemplatePath);
            }

            if (reportRequests.Count > 0)
            {
                _logger.LogInformation("Parsed {Count} valid report request(s).", reportRequests.Count);
            }
            else
            {
                _logger.LogWarning("No valid report requests were parsed.");
            }

            return reportRequests;
        }


        private static async Task GenerateReportAsync(string templatePath, int rowCount)
        {
            string outputPath = GenerateOutputFilePath(templatePath);

            _logger.LogInformation("Generating report for {TemplatePath} with {RowCount} rows...", templatePath, rowCount);

            try
            {
                await Task.Run(() => ReportGeneration.GenerateReport(templatePath, outputPath, rowCount));
                _logger.LogInformation("Report generated successfully: {OutputPath}", outputPath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating report for {TemplatePath}", templatePath);
            }
        }

        private static string GenerateOutputFilePath(string templatePath)
        {
            string directory = Path.GetDirectoryName(templatePath) ?? "";
            string fileName  = Path.GetFileNameWithoutExtension(templatePath);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            return Path.Combine(directory, $"{fileName}_{timestamp}.xlsx");
        }
    }

    public record ReportRequest(string TemplatePath, int RowCount);
}
