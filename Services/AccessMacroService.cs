using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace NetSuiteAutomation.Services
{
    public class AccessMacroService
    {
        private readonly string _logFolder = @"C:\ADSK-Automation\Logs";
        private readonly string _exePath = @"C:\ADSK-Automation\Release\AccessMacroRunner.exe";

        private readonly ChildProcessRunner _processRunner;

        public AccessMacroService(ChildProcessRunner processRunner)
        {
            _processRunner = processRunner;
        }

        public async Task RunAccessMacroAndExportAsync(CancellationToken cancellationToken)
        {
            string logFilePath = Path.Combine(_logFolder, "AccessMacroIssues.txt");

            try
            {
                Directory.CreateDirectory(_logFolder);
                Log(logFilePath, "🚀 Launching AccessMacroRunner.exe...");

                var result = await _processRunner.RunAsync(_exePath, cancellationToken: cancellationToken);
                if (!string.IsNullOrWhiteSpace(result.StandardOutput))
                    Log(logFilePath, "[OUT] " + result.StandardOutput.Trim());
                if (!string.IsNullOrWhiteSpace(result.StandardError))
                    Log(logFilePath, "[ERR] " + result.StandardError.Trim());

                Log(logFilePath, $"✅ AccessMacroRunner.exe completed with exit code {result.ExitCode}.");
                if (result.ExitCode != 0)
                    throw new InvalidOperationException("AccessMacroRunner exited with code " + result.ExitCode + ".");
            }
            catch (Exception ex)
            {
                Log(logFilePath, $"❌ Error launching AccessMacroRunner.exe: {ex.Message}");
                throw;
            }
        }

        private void Log(string logPath, string message)
        {
            File.AppendAllText(logPath, $"{DateTime.Now} - {message}{Environment.NewLine}");
        }
    }
}
