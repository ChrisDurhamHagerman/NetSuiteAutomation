using Microsoft.AspNetCore.Mvc;
using NetSuiteAutomation.Services;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NetSuiteAutomation.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CreditSafeController : ControllerBase
    {
        private readonly LogService _log;
        private readonly BackgroundTaskQueue _backgroundTasks;

        private static readonly string ExportRoot = @"C:\ADSK-Automation\CreditSafe Imports\NetSuiteExports";
        private static readonly string ArchiveRoot = Path.Combine(ExportRoot, "Old Imports");
        private static readonly string CreditSafeExe = @"C:\ADSK-Automation\CreditSafe\CreditSafeController.exe";
        private static readonly string CustomerFile = "CustomerSync.csv";
        private static readonly string InvoiceFile = "InvoiceDSOSync.csv";

        public CreditSafeController(LogService log, BackgroundTaskQueue backgroundTasks)
        {
            _log = log;
            _backgroundTasks = backgroundTasks;
        }

        public class ImportPayload
        {
            public string customerCsv { get; set; }
            public string invoiceCsv { get; set; }
        }

        [HttpPost("import")]
        public IActionResult Import([FromBody] ImportPayload payload)
        {
            _log.Log("Received CreditSafe import payload.");

            if (payload == null)
            {
                _log.Log("Invalid request: payload is null.");
                return BadRequest("Payload is required.");
            }
            if (string.IsNullOrWhiteSpace(payload.customerCsv))
            {
                _log.Log("Invalid request: customerCsv is empty.");
                return BadRequest("customerCsv is required.");
            }
            if (string.IsNullOrWhiteSpace(payload.invoiceCsv))
            {
                _log.Log("Invalid request: invoiceCsv is empty.");
                return BadRequest("invoiceCsv is required.");
            }

            try
            {
                Directory.CreateDirectory(ExportRoot);
                Directory.CreateDirectory(ArchiveRoot);

                var customerPath = Path.Combine(ExportRoot, CustomerFile);
                var invoicePath = Path.Combine(ExportRoot, InvoiceFile);

                SaveWithArchive(customerPath, payload.customerCsv);
                SaveWithArchive(invoicePath, payload.invoiceCsv);

                _log.Log("CreditSafe NetSuite export files saved successfully.");

                // Respond to NetSuite quickly
                var response = Ok(new
                {
                    status = "ok",
                    customerPath,
                    invoicePath
                });

                // Kick off the downstream processor in the background
                _backgroundTasks.Enqueue(async (services, cancellationToken) =>
                {
                    var log = services.GetRequiredService<LogService>();
                    var processRunner = services.GetRequiredService<ChildProcessRunner>();
                    try
                    {
                        var result = await processRunner.RunAsync(CreditSafeExe, cancellationToken: cancellationToken);

                        if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                            log.Log($"CreditSafeController.exe output: {result.StandardOutput.Trim()}");
                        if (!string.IsNullOrWhiteSpace(result.StandardError))
                            log.Log($"CreditSafeController.exe error: {result.StandardError.Trim()}");

                        log.Log($"CreditSafeController.exe exited with code {result.ExitCode}.");
                        if (result.ExitCode != 0)
                            throw new InvalidOperationException("CreditSafe runner exited with code " + result.ExitCode + ".");
                    }
                    catch (Exception ex)
                    {
                        log.Log($"Error launching CreditSafeController.exe: {ex.Message}");
                        throw;
                    }
                });

                return response;
            }
            catch (Exception ex)
            {
                _log.Log("Error during CreditSafe import pipeline: " + ex.Message);
                return StatusCode(500, "An error occurred while processing the data.");
            }
        }

        private void SaveWithArchive(string path, string csv)
        {
            if (System.IO.File.Exists(path))
            {
                string archived = Path.Combine(
                    ArchiveRoot,
                    $"{Path.GetFileNameWithoutExtension(path)}_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
                System.IO.File.Move(path, archived);
                _log.Log($"Archived {Path.GetFileName(path)} to {archived}");
            }

            System.IO.File.WriteAllText(path, csv, Encoding.UTF8);
            _log.Log($"Saved {Path.GetFileName(path)} ({csv.Length} bytes).");
        }
    }
}
