using Microsoft.AspNetCore.Mvc;
using NetSuiteAutomation.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetSuiteAutomation.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CsvImportController : ControllerBase
    {
        private readonly AccessImportService _accessImportService;
        private readonly ImportFormaService _importFormaService;
        private readonly LogService _log;
        private readonly AccessMacroService _accessMacroService;
        private readonly string _importFolder = @"C:\ADSK-Automation\NetSuiteImports";

        public CsvImportController(
            AccessImportService accessImportService,
            ImportFormaService importFormaService,
            AccessMacroService accessMacroService,
            LogService log)
        {
            _accessImportService = accessImportService;
            _importFormaService = importFormaService;
            _log = log;
            _accessMacroService = accessMacroService;
        }

        [HttpPost("import")]
        public IActionResult ImportFromJson([FromBody] List<Dictionary<string, string>> jsonData)
        {
            _log.Log("📥 Received import request from NetSuite.");

            if (jsonData == null || jsonData.Count == 0)
            {
                _log.Log("⚠️ No data received in request body.");
                return BadRequest("No data received.");
            }

            _log.Log($"📊 Processing {jsonData.Count} records from NetSuite.");

            try
            {
                // Convert to CSV
                var csv = ConvertJsonToCsv(jsonData);
                _log.Log("🛠 CSV conversion completed.");

                // Ensure folder exists
                Directory.CreateDirectory(_importFolder);

                // Save CSV to file
                var fileName = $"netsuite_import_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                var filePath = Path.Combine(_importFolder, fileName);
                System.IO.File.WriteAllText(filePath, csv);
                _log.Log($"💾 CSV file written to: {filePath}");

                // Import into Access
                _log.Log("📂 Starting import into Access database...");
                bool importSuccess = _accessImportService.ImportCsvToAccess(filePath);

                if (!importSuccess)
                {
                    _log.Log("❌ Access import failed. Check NetsuiteImportIssues.txt for detailed error info.");
                    return StatusCode(500, "Access import failed.");
                }

                _log.Log("✅ Access import completed successfully. Data inserted into [dbo_autodesk_Sales_Orders].");

                // 🟢 Immediately return response to NetSuite
                var response = Ok(new
                {
                    message = "Data received and CSV imported. Background processing started.",
                    rowCount = jsonData.Count,
                    filePath = filePath
                });

                // 🔄 Start Forma import and macro execution in background
                Task.Run(async () =>
                {
                    try
                    {
                        _log.Log("📥 Beginning Forma CSV import...");
                        bool formaSuccess = await _importFormaService.ImportFormaDataAsync();

                        if (formaSuccess)
                        {
                            _log.Log("✅ Forma CSV import completed.");
                            _log.Log("🧩 Starting Access macro execution and export to CSV...");

                            await Task.Delay(10000); // Allow file locks to release

                            try
                            {
                                _accessMacroService.RunAccessMacroAndExport();
                                _log.Log("✅ Macros ran and CSV exported.");
                            }
                            catch (Exception macroEx)
                            {
                                _log.Log($"❌ Error running macro and exporting CSV: {macroEx.Message}");
                            }
                        }
                        else
                        {
                            _log.Log("❌ Forma import failed. Check FormaImportIssues.txt for detailed error info.");
                        }
                    }
                    catch (Exception bgEx)
                    {
                        _log.Log($"❌ Background task error: {bgEx.Message}");
                    }
                });

                return response;
            }
            catch (Exception ex)
            {
                _log.Log("❌ Error during CSV import pipeline: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the data.");
            }
        }
    

        [HttpGet("download/NetsuiteImportData.csv")]
        public IActionResult DownloadNetsuiteImportCsv()
        {
            try
            {
                string filePath = @"C:\ADSK-Automation\Logs\NetsuiteImportData.csv";

                if (!System.IO.File.Exists(filePath))
                {
                    _log.Log("❌ Requested CSV file does not exist: " + filePath);
                    return NotFound("File not found.");
                }

                byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
                _log.Log("📤 NetsuiteImportData.csv served to caller.");

                // ✅ Schedule background task AFTER file is sent to NetSuite
                HttpContext.Response.OnCompleted(() =>
                {
                    try
                    {
                        var runner = @"C:\ADSK-Automation\Release\AccessMacroRunner.exe";
                        var psi = new ProcessStartInfo
                        {
                            FileName = runner,
                            Arguments = "emailreports",
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };
                        Process.Start(psi);
                        _log.Log("🚀 Launched AccessMacroRunner.exe emailreports");
                    }
                    catch (Exception ex)
                    {
                        _log.Log("❌ Could not launch AccessMacroRunner: " + ex.Message);
                    }
                    return Task.CompletedTask;
                });

                return File(fileBytes, "text/csv", "NetsuiteImportData.csv");
            }
            catch (Exception ex)
            {
                _log.Log("❌ Error serving NetsuiteImportData.csv: " + ex.Message);
                return StatusCode(500, "An error occurred while retrieving the file.");
            }
        }



        private string ConvertJsonToCsv(List<Dictionary<string, string>> data)
        {
            var sb = new StringBuilder();
            var headers = data.First().Keys.ToList();
            sb.AppendLine(string.Join(",", headers));

            foreach (var row in data)
            {
                var line = string.Join(",", headers.Select(h => Sanitize(row.GetValueOrDefault(h, ""))));
                sb.AppendLine(line);
            }

            return sb.ToString();
        }

        private string Sanitize(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            input = input.Replace("\"", "\"\"");
            return input.Contains(",") ? $"\"{input}\"" : input;
        }
    }
}
