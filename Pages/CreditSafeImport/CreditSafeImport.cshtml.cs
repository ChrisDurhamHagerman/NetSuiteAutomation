using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using System;
using System.IO;
using System.Threading.Tasks;

namespace NetSuiteAutomation.Pages
{
    public class CreditSafeImportModel : PageModel
    {
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
                CallCreditSafeLogicController(newFilePath);

                Message = $"File '{FileUpload.FileName}' uploaded. Processing has started.";
            }
            catch (Exception ex)
            {
                LogMessage($"Error uploading file: {ex.Message}");
                Message = $"Error uploading file: {ex.Message}";
            }

            return Page();
        }

        private void CallCreditSafeLogicController(string excelFilePath)
        {
            try
            {
                string exePath = @"C:\ADSK-Automation\CreditSafe\CreditSafeLogicController.exe";
                if (!System.IO.File.Exists(exePath))
                {
                    LogMessage("CreditSafeLogicController.exe not found in C:\\ADSK-Automation\\CreditSafe, skipping logic run.");
                    return;
                }

                var processInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = $"\"{excelFilePath}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(exePath)
                };

                // Fire-and-forget: run in the background so the page returns immediately
                _ = Task.Run(() =>
                {
                    try
                    {
                        using (var process = System.Diagnostics.Process.Start(processInfo))
                        {
                            string output = process.StandardOutput.ReadToEnd();
                            string error = process.StandardError.ReadToEnd();
                            process.WaitForExit();

                            if (!string.IsNullOrWhiteSpace(output))
                                LogMessage($"CreditSafeLogicController Output: {output.Trim()}");
                            if (!string.IsNullOrWhiteSpace(error))
                                LogMessage($"CreditSafeLogicController Error: {error.Trim()}");

                            LogMessage($"CreditSafeLogicController exited with code {process.ExitCode}.");
                        }
                    }
                    catch (Exception runEx)
                    {
                        LogMessage($"Error running CreditSafeLogicController: {runEx.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                LogMessage($"Error preparing to call CreditSafeLogicController: {ex.Message}");
            }
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
