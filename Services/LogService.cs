using System;
using System.IO;
using System.Linq;

namespace NetSuiteAutomation.Services
{
    public class LogService
    {
        private readonly string _logFile = @"C:\ADSK-Automation\Logs\automation_log.txt";

        public LogService()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_logFile));
        }

        public void Log(string message)
        {
            string entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
            File.AppendAllText(_logFile, entry + Environment.NewLine);
        }

        public string[] GetLogLines(int maxLines = 1000)
        {
            if (!File.Exists(_logFile))
                return new[] { "No logs yet." };

            var lines = File.ReadAllLines(_logFile);
            return lines.Reverse().Take(maxLines).ToArray();
        }
    }
}
