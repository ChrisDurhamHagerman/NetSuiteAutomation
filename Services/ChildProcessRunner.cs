using System.Diagnostics;

namespace NetSuiteAutomation.Services;

public sealed record ChildProcessResult(int ExitCode, string StandardOutput, string StandardError);

public sealed class ChildProcessRunner
{
    public async Task<ChildProcessResult> RunAsync(
        string executablePath,
        IEnumerable<string>? arguments = null,
        CancellationToken cancellationToken = default)
    {
        if (!File.Exists(executablePath))
            throw new FileNotFoundException("Child process executable was not found.", executablePath);

        var startInfo = new ProcessStartInfo
        {
            FileName = executablePath,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            WorkingDirectory = Path.GetDirectoryName(executablePath) ?? AppContext.BaseDirectory
        };

        if (arguments != null)
        {
            foreach (string argument in arguments)
                startInfo.ArgumentList.Add(argument);
        }

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Child process could not be started: " + executablePath);

        Task<string> outputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        Task<string> errorTask = process.StandardError.ReadToEndAsync(cancellationToken);

        try
        {
            await process.WaitForExitAsync(cancellationToken);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            try
            {
                if (!process.HasExited)
                    process.Kill(entireProcessTree: true);
            }
            catch { }

            throw;
        }

        return new ChildProcessResult(
            process.ExitCode,
            await outputTask,
            await errorTask);
    }
}
