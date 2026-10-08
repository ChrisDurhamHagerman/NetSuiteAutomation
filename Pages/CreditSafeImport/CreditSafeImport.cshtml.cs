using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using System;
using System.IO;
using System.Threading.Tasks;
using NetSuiteAutomation.Services;
using System.Diagnostics;

namespace NetSuiteAutomation.Pages
{
    public class CreditSafeImportModel : PageModel
    {
        private readonly BackgroundTaskQueue _backgroundTasks;

        public CreditSafeImportModel(BackgroundTaskQueue backgroundTasks)
        {
            _backgroundTasks = backgroundTasks;
        }

        private readonly string _importFolder = @"C:\ADSK-Automation\CreditSafe Imports";
        private readonly string _logFilePath = @"C:\ADSK-Automation\automation_log.txt";
        private readonly string _oldImportsFolder = @"C:\ADSK-Automation\CreditSafe Imports\Old Imports";

        [BindProperty]
        public IFormFile FileUpload { get; set; }

        public string Message { get; set; }

        public void OnGet()
        {
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (FileUpload == null || FileUpload.Length == 0)
            {
                Message = "Please select an Excel file to upload.";
                return Page();
            }

            // Optional: enforce Excel file types
            var ext = Path.GetExtension(FileUpload.FileName);
            if (!string.Equals(ext, ".xlsx", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(ext, ".xls", StringComparison.OrdinalIgnoreCase))
            {
                Message = "Please upload an Excel file (.xlsx or .xls).";
                return Page();
            }

            try
            {
                LogMessage($"Upload initiated for file: {FileUpload.FileName}");

                // Ensure folders exist
                if (!Directory.Exists(_importFolder))
                {
                    Directory.CreateDirectory(_importFolder);
                    LogMessage($"Created folder: {_importFolder}");
                }

                if (!Directory.Exists(_oldImportsFolder))
                {
                    Directory.CreateDirectory(_oldImportsFolder);
                    LogMessage($"Created folder: {_oldImportsFolder}");
                }

                // Move existing files to Old Imports (timestamped)
                var existingFiles = Directory.GetFiles(_importFolder, "*.*", SearchOption.TopDirectoryOnly);
                foreach (var existingFile in existingFiles)
                {
                    // Skip subfolder guard (shouldn't match with TopDirectoryOnly, but keep parity with prior code)
                    if (Path.GetDirectoryName(existingFile)?.Equals(_oldImportsFolder, StringComparison.OrdinalIgnoreCase) == true)
                        continue;

                    try
                    {
                        var fileName = Path.GetFileName(existingFile);
                        var destinationPath = Path.Combine(
                            _oldImportsFolder,
                            $"{Path.GetFileNameWithoutExtension(fileName)}_{DateTime.Now:yyyyMMdd_HHmmss}{Path.GetExtension(fileName)}");

                        System.IO.File.Move(existingFile, destinationPath);
                        LogMessage($"Moved old file to: {destinationPath}");
                    }
                    catch (Exception moveEx)
                    {
                        LogMessage($"Failed to move '{existingFile}': {moveEx.Message}");
                    }
                }

                // Save new file
                var newFilePath = Path.Combine(_importFolder, FileUpload.FileName);
                using (var stream = new FileStream(newFilePath, FileMode.Create))
                {
                    await FileUpload.CopyToAsync(stream);
                }

                LogMessage($"New file saved successfully: {newFilePath}");

                // Launch CreditSafeLogicController in background (non-blocking)
                EnqueueCreditSafeLogicController(newFilePath);

                Message = $"File '{FileUpload.FileName}' uploaded. Processing has started.";
            }
            catch (Exception ex)
            {
                LogMessage($"Error uploading file: {ex.Message}");
                Message = $"Error uploading file: {ex.Message}";
            }

            return Page();
        }

        private void EnqueueCreditSafeLogicController(string excelFilePath)
        {
            _backgroundTasks.Enqueue(async (services, cancellationToken) =>
            {
                var log = services.GetRequiredService<LogService>();
                var processRunner = services.GetRequiredService<ChildProcessRunner>();
                string exePath = @"C:\ADSK-Automation\CreditSafe\CreditSafeLogicController.exe";
                var result = await processRunner.RunAsync(exePath, new[] { excelFilePath }, cancellationToken);

                if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                    log.Log("CreditSafeLogicController output: " + result.StandardOutput.Trim());
                if (!string.IsNullOrWhiteSpace(result.StandardError))
                    log.Log("CreditSafeLogicController error: " + result.StandardError.Trim());

                log.Log("CreditSafeLogicController exited with code " + result.ExitCode + ".");
                if (result.ExitCode != 0)
                    throw new InvalidOperationException("CreditSafeLogicController exited with code " + result.ExitCode + ".");
            });
        }

        private void LogMessage(string message)
        {
            try
            {
                var dir = Path.GetDirectoryName(_logFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}";
                System.IO.File.AppendAllText(_logFilePath, logEntry);
            }
            catch
            {
                // Intentionally swallow logging failures to avoid breaking UX
            }
        }
    }
}
