using Microsoft.AspNetCore.Mvc;
using NetSuiteAutomation.Services;
using System.Diagnostics;
using System.Text;

namespace NetSuiteAutomation.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class TwoQuarterImportController : ControllerBase
    {
        private static readonly string[] CsvHeaders =
        {
            "InternalID", "Customer", "SO #", "Autodesk Quote #", "Line ID",
            "Autodesk Line #", "Item", "Quantity", "Item Rate", "Unit Customer Price",
            "AD Start Date", "AD End Date", "Order Date"
        };

        private const string ImportFolder = @"C:\ADSK-Automation\2Qtrs";
        private const string ImportFileName = "JMHAutodeskSalesOrders2QtrsAllSyncResults.csv";
        private const string RunnerPath = @"C:\ADSK-Automation\Release\AccessMacroRunner.exe";

        private readonly LogService _log;

        public TwoQuarterImportController(LogService log)
        {
            _log = log;
        }

        [HttpPost("import")]
        public IActionResult Import([FromBody] List<Dictionary<string, string>> jsonData)
        {
            if (jsonData == null || jsonData.Count == 0)
            {
                _log.Log("Two-quarter import rejected: no data received.");
                return BadRequest("No data received.");
            }

            string temporaryPath = string.Empty;
            try
            {
                ValidatePayload(jsonData);

                Directory.CreateDirectory(ImportFolder);
                string archiveFolder = Path.Combine(ImportFolder, "Old");
                Directory.CreateDirectory(archiveFolder);

                temporaryPath = Path.Combine(
                    ImportFolder,
                    ".twoqtrs_" + Guid.NewGuid().ToString("N") + ".tmp");

                WriteCsv(jsonData, temporaryPath);

                foreach (string existingFile in Directory.GetFiles(ImportFolder, "*.csv", SearchOption.TopDirectoryOnly))
                    ArchiveFile(existingFile, archiveFolder);

                string importPath = Path.Combine(ImportFolder, ImportFileName);
                System.IO.File.Move(temporaryPath, importPath);
                temporaryPath = string.Empty;

                if (!System.IO.File.Exists(RunnerPath))
                    throw new FileNotFoundException("AccessMacroRunner executable not found.", RunnerPath);

                var process = Process.Start(new ProcessStartInfo
                {
                    FileName = RunnerPath,
                    Arguments = "twoqtrs",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(RunnerPath) ?? ImportFolder
                });

                if (process == null)
                    throw new InvalidOperationException("AccessMacroRunner could not be started.");

                _log.Log(
                    "Two-quarter import accepted. Saved " + jsonData.Count
                    + " rows to " + importPath + " and started AccessMacroRunner PID " + process.Id + ".");

                return Accepted(new
                {
                    message = "Two-quarter data received. Access import and report processing started.",
                    rowCount = jsonData.Count,
                    filePath = importPath,
                    processId = process.Id
                });
            }
            catch (InvalidDataException ex)
            {
                _log.Log("Two-quarter import validation failed: " + ex.Message);
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                _log.Log("Two-quarter import failed: " + ex);
                return StatusCode(500, "Two-quarter import failed.");
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(temporaryPath) && System.IO.File.Exists(temporaryPath))
                {
                    try { System.IO.File.Delete(temporaryPath); } catch { }
                }
            }
        }

        private static void ValidatePayload(IEnumerable<Dictionary<string, string>> rows)
        {
            int rowNumber = 0;
            foreach (Dictionary<string, string> row in rows)
            {
                rowNumber++;
                if (row == null)
                    throw new InvalidDataException("Row " + rowNumber + " is null.");

                foreach (string header in CsvHeaders)
                {
                    if (!TryGetValue(row, header, out _))
                        throw new InvalidDataException("Row " + rowNumber + " is missing field '" + header + "'.");
                }

                TryGetValue(row, "InternalID", out string internalId);
                if (string.IsNullOrWhiteSpace(internalId))
                    throw new InvalidDataException("Row " + rowNumber + " has a blank InternalID.");
            }
        }

        private static void WriteCsv(IEnumerable<Dictionary<string, string>> rows, string path)
        {
            using (var writer = new StreamWriter(path, false, new UTF8Encoding(false)))
            {
                writer.WriteLine(string.Join(",", CsvHeaders.Select(EscapeCsv)));

                foreach (Dictionary<string, string> row in rows)
                {
                    var values = CsvHeaders.Select(header =>
                    {
                        TryGetValue(row, header, out string value);
                        return EscapeCsv(value ?? string.Empty);
                    });
                    writer.WriteLine(string.Join(",", values));
                }
            }
        }

        private static bool TryGetValue(
            Dictionary<string, string> row,
            string key,
            out string value)
        {
            if (row.TryGetValue(key, out value))
                return true;

            foreach (KeyValuePair<string, string> item in row)
            {
                if (item.Key.Equals(key, StringComparison.OrdinalIgnoreCase))
                {
                    value = item.Value;
                    return true;
                }
            }

            value = string.Empty;
            return false;
        }

        private static string EscapeCsv(string value)
        {
            string safeValue = value ?? string.Empty;
            if (safeValue.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
                return "\"" + safeValue.Replace("\"", "\"\"") + "\"";
            return safeValue;
        }

        private static void ArchiveFile(string sourcePath, string archiveFolder)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string baseName = Path.GetFileNameWithoutExtension(sourcePath);
            string extension = Path.GetExtension(sourcePath);
            string destination = Path.Combine(archiveFolder, baseName + "_" + timestamp + extension);
            int suffix = 1;

            while (System.IO.File.Exists(destination))
            {
                destination = Path.Combine(
                    archiveFolder,
                    baseName + "_" + timestamp + "_" + suffix + extension);
                suffix++;
            }

            System.IO.File.Move(sourcePath, destination);
        }
    }
}
