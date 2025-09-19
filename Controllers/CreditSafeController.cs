using Microsoft.AspNetCore.Mvc;
using NetSuiteAutomation.Services;
using System;
using System.Diagnostics;
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

        private static readonly string ExportRoot = @"C:\ADSK-Automation\CreditSafe Imports\NetSuiteExports";
        private static readonly string ArchiveRoot = Path.Combine(ExportRoot, "Old Imports");
        private static readonly string CreditSafeExe = @"C:\ADSK-Automation\CreditSafe\CreditSafeController.exe";
        private static readonly string CustomerFile = "CustomerSync.csv";
        private static readonly string InvoiceFile = "InvoiceDSOSync.csv";

        public CreditSafeController(LogService log)
        {
            _log = log;
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
                Task.Run(() =>
                {
                    try
                    {
                        if (!System.IO.File.Exists(CreditSafeExe))
                        {
                            _log.Log($"CreditSafeController.exe not found at: {CreditSafeExe}. Skipping launch.");
                            return;
                        }

                        var psi = new ProcessStartInfo
                        {
                            FileName = CreditSafeExe,
                            Arguments = "", // to be defined when you add CLI args
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true
                        };

                        using var p = Process.Start(psi);
                        string stdOut = p.StandardOutput.ReadToEnd();
                        string stdErr = p.StandardError.ReadToEnd();
                        p.WaitForExit();

                        if (!string.IsNullOrWhiteSpace(stdOut))
                            _log.Log($"CreditSafeController.exe output: {stdOut.Trim()}");
                        if (!string.IsNullOrWhiteSpace(stdErr))
                            _log.Log($"CreditSafeController.exe error: {stdErr.Trim()}");

                        _log.Log($"CreditSafeController.exe exited with code {p.ExitCode}.");
                    }
                    catch (Exception ex)
                    {
                        _log.Log($"Error launching CreditSafeController.exe: {ex.Message}");
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
