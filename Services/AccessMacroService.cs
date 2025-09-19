using System;
using System.Diagnostics;
using System.IO;

namespace NetSuiteAutomation.Services
{
    public class AccessMacroService
    {
        private readonly string _logFolder = @"C:\ADSK-Automation\Logs";
        private readonly string _exePath = @"C:\ADSK-Automation\Release\AccessMacroRunner.exe";

        public void RunAccessMacroAndExport()
        {
            string logFilePath = Path.Combine(_logFolder, "AccessMacroIssues.txt");

            try
            {
                Directory.CreateDirectory(_logFolder);
                Log(logFilePath, "🚀 Launching AccessMacroRunner.exe...");

                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = _exePath,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using (Process process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.OutputDataReceived += (sender, args) =>
                    {
                        if (!string.IsNullOrEmpty(args.Data))
                            Log(logFilePath, $"[OUT] {args.Data}");
                    };
                    process.ErrorDataReceived += (sender, args) =>
                    {
                        if (!string.IsNullOrEmpty(args.Data))
                            Log(logFilePath, $"[ERR] {args.Data}");
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    process.WaitForExit();

                    Log(logFilePath, $"✅ AccessMacroRunner.exe completed with exit code {process.ExitCode}.");
                }
            }
            catch (Exception ex)
            {
                Log(logFilePath, $"❌ Error launching AccessMacroRunner.exe: {ex.Message}");
            }
        }

        private void Log(string logPath, string message)
        {
            File.AppendAllText(logPath, $"{DateTime.Now} - {message}{Environment.NewLine}");
        }
    }
}
