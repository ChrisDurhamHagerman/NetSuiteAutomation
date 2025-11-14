using Microsoft.AspNetCore.Mvc;
using NetSuiteAutomation.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetSuiteAutomation.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class WorkflowMessageTransactionController : ControllerBase
    {
        private readonly ProjectActivityService _projectActivityService;
        private readonly LogService _logService;

        private readonly string _folderPath = @"C:\ADSK-Automation\WorkflowMessageTransaction";
        private readonly string _archiveFolder = @"C:\ADSK-Automation\WorkflowMessageTransaction\OldImports";
        private readonly string _logPath = @"C:\ADSK-Automation\Logs\WorkflowImportIssues.txt";

        public WorkflowMessageTransactionController(ProjectActivityService projectActivityService, LogService logService)
        {
            _projectActivityService = projectActivityService;
            _logService = logService;
        }

        [HttpPost("import")]
        public IActionResult Import([FromBody] List<Dictionary<string, string>> jsonData)
        {
            _logService.Log("Received WorkflowMessageTransaction import request.");

            if (jsonData == null || jsonData.Count == 0)
            {
                _logService.Log("No data received in request.");
                return BadRequest("No data received.");
            }

            try
            {
                Directory.CreateDirectory(_folderPath);
                Directory.CreateDirectory(_archiveFolder);

                // Archive prior CSVs
                foreach (var file in Directory.GetFiles(_folderPath, "*.csv"))
                {
                    string name = Path.GetFileNameWithoutExtension(file);
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string archiveName = $"{name}_{timestamp}.csv";
                    System.IO.File.Move(file, Path.Combine(_archiveFolder, archiveName));
                }

                string baseName = $"workflow_import_{DateTime.Now:yyyyMMdd_HHmmss}";
                string debugCsvPath = Path.Combine(_folderPath, $"{baseName}.csv");
                string accessCsvPath = Path.Combine(_folderPath, $"{baseName}_access.csv");

                // 1) Debug CSV (fully quoted, \n shown as \n for log readability)
                WriteDebugCsv(jsonData, debugCsvPath, _logPath);

                // 2) Access-safe CSV (exactly the format you want)
                WriteAccessCsv(jsonData, accessCsvPath);

                System.IO.File.AppendAllText(_logPath, $"[{DateTime.Now}] INFO: Debug CSV written: {debugCsvPath}{Environment.NewLine}");
                System.IO.File.AppendAllText(_logPath, $"[{DateTime.Now}] INFO: Access CSV written: {accessCsvPath}{Environment.NewLine}");

                // Ensure file handles closed before Access
                GC.Collect();
                GC.WaitForPendingFinalizers();

                bool success = _projectActivityService.ImportWorkflowCsvToAccess(accessCsvPath);
                if (!success)
                {
                    _logService.Log("Workflow import failed. See WorkflowImportIssues.txt for details.");
                    return StatusCode(500, "Access import failed.");
                }

                _logService.Log("WorkflowMessageTransaction import completed successfully.");
                return Ok(new { message = "Import completed successfully.", rowCount = jsonData.Count });
            }
            catch (Exception ex)
            {
                _logService.Log($"Unexpected error during workflow import: {ex.Message}");
                return StatusCode(500, "An unexpected error occurred.");
            }
        }

        // Fully quoted, escapes quotes; flattens newlines to \n for inspection
        private void WriteDebugCsv(List<Dictionary<string, string>> data, string filePath, string logPath)
        {
            var headers = data.First().Keys.ToList();

            using var writer = new StreamWriter(filePath, false, Encoding.UTF8);
            writer.NewLine = "\r\n";
            writer.WriteLine(string.Join(",", headers));
            System.IO.File.AppendAllText(logPath, $"[{DateTime.Now}] INFO: CSV Headers: {string.Join(" | ", headers)}{Environment.NewLine}");

            int rowNumber = 1;
            foreach (var row in data)
            {
                try
                {
                    string line = string.Join(",", headers.Select(h =>
                    {
                        string value = row.GetValueOrDefault(h, "") ?? "";
                        value = value.Replace("\r\n", "\\n").Replace("\r", "\\n").Replace("\n", "\\n");
                        return $"\"{value.Replace("\"", "\"\"")}\"";
                    }));

                    writer.WriteLine(line);
                    System.IO.File.AppendAllText(logPath, $"[{DateTime.Now}] ROW {rowNumber} - Raw JSON: {System.Text.Json.JsonSerializer.Serialize(row)}{Environment.NewLine}");
                    System.IO.File.AppendAllText(logPath, $"[{DateTime.Now}] ROW {rowNumber} - CSV Line: {line}{Environment.NewLine}");
                }
                catch (Exception ex)
                {
                    System.IO.File.AppendAllText(logPath, $"[{DateTime.Now}] ERROR processing row {rowNumber}: {ex.Message}{Environment.NewLine}");
                }
                rowNumber++;
            }
        }

        // ACCESS FILE FORMAT YOU REQUESTED:
        // - No quotes for any field except the last column (Comment), which stays quoted (to keep commas)
        // - All newlines -> single spaces
        // - All quotes removed from content
        // - Commas removed from non-Comment fields
        // - Write as ANSI (Encoding.Default) to play nicely with ACE text driver
        private void WriteAccessCsv(List<Dictionary<string, string>> data, string filePath)
        {
            var headers = data.First().Keys.ToList();
            using var writer = new StreamWriter(filePath, false, Encoding.Default);
            writer.NewLine = "\r\n";
            writer.WriteLine(string.Join(",", headers));

            foreach (var row in data)
            {
                var fields = new List<string>();
                for (int i = 0; i < headers.Count; i++)
                {
                    string header = headers[i];
                    string value = row.GetValueOrDefault(header, "") ?? "";

                    // Normalize
                    value = value.Replace("\r\n", " ")
                                 .Replace("\r", " ")
                                 .Replace("\n", " ")
                                 .Replace("\"", "")
                                 .Trim();

                    if (i == headers.Count - 1)
                    {
                        // Keep commas inside Comment by quoting that one field
                        fields.Add($"\"{value}\"");
                    }
                    else
                    {
                        // Remove commas entirely from other fields
                        value = value.Replace(",", " ");
                        fields.Add(value);
                    }
                }

                writer.WriteLine(string.Join(",", fields));
            }
        }
    }
}
