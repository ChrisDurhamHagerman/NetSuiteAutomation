using System.Threading.Channels;

namespace NetSuiteAutomation.Services;

public sealed class BackgroundTaskQueue
{
    private readonly Channel<Func<IServiceProvider, CancellationToken, Task>> _tasks =
        Channel.CreateUnbounded<Func<IServiceProvider, CancellationToken, Task>>(new UnboundedChannelOptions
        {
            SingleReader = true,
            SingleWriter = false
        });

    public void Enqueue(Func<IServiceProvider, CancellationToken, Task> task)
    {
        ArgumentNullException.ThrowIfNull(task);
        if (!_tasks.Writer.TryWrite(task))
            throw new InvalidOperationException("The background task queue is no longer accepting work.");
    }

    internal ValueTask<Func<IServiceProvider, CancellationToken, Task>> DequeueAsync(CancellationToken cancellationToken) =>
        _tasks.Reader.ReadAsync(cancellationToken);
}
