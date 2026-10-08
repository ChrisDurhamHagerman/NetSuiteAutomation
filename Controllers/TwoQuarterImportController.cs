using Microsoft.AspNetCore.Mvc;
using NetSuiteAutomation.Services;
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
        private const string IncomingFolder = @"C:\ADSK-Automation\2Qtrs\Incoming";
        private const string RunnerPath = @"C:\ADSK-Automation\Release\AccessMacroRunner.exe";

        private readonly LogService _log;
        private readonly BackgroundTaskQueue _backgroundTasks;

        public TwoQuarterImportController(LogService log, BackgroundTaskQueue backgroundTasks)
        {
            _log = log;
            _backgroundTasks = backgroundTasks;
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

                Directory.CreateDirectory(IncomingFolder);

                temporaryPath = Path.Combine(
                    IncomingFolder,
                    ".twoqtrs_" + Guid.NewGuid().ToString("N") + ".tmp");

                WriteCsv(jsonData, temporaryPath);

                string jobId = Guid.NewGuid().ToString("N");
                string importPath = Path.Combine(
                    IncomingFolder,
                    "JMHAutodeskSalesOrders2QtrsAllSyncResults_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                    + "_" + jobId.Substring(0, 8) + ".csv");
                System.IO.File.Move(temporaryPath, importPath);
                temporaryPath = string.Empty;

                if (!System.IO.File.Exists(RunnerPath))
                    throw new FileNotFoundException("AccessMacroRunner executable not found.", RunnerPath);

                _backgroundTasks.Enqueue(async (services, cancellationToken) =>
                {
                    var log = services.GetRequiredService<LogService>();
                    var processRunner = services.GetRequiredService<ChildProcessRunner>();
                    try
                    {
                        var result = await processRunner.RunAsync(
                            RunnerPath,
                            new[] { "twoqtrs", importPath },
                            cancellationToken);

                        if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                            log.Log("Two-quarter runner output: " + result.StandardOutput.Trim());
                        if (!string.IsNullOrWhiteSpace(result.StandardError))
                            log.Log("Two-quarter runner error: " + result.StandardError.Trim());
                        if (result.ExitCode != 0)
                            throw new InvalidOperationException("Two-quarter runner exited with code " + result.ExitCode + ".");

                        log.Log("Two-quarter runner completed job " + jobId + ".");
                    }
                    catch (Exception ex)
                    {
                        log.Log("Two-quarter runner failed job " + jobId + ": " + ex.Message);
                        throw;
                    }
                });

                _log.Log(
                    "Two-quarter import accepted. Saved " + jsonData.Count
                    + " rows to " + importPath + " and queued AccessMacroRunner job " + jobId + ".");

                return Accepted(new
                {
                    message = "Two-quarter data received. Access import and report processing queued.",
                    jobId,
                    status = "queued",
                    rowCount = jsonData.Count,
                    filePath = importPath
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

    }
}
