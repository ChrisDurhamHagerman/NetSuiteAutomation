namespace NetSuiteAutomation.Services;

public sealed class BackgroundTaskWorker : BackgroundService
{
    private readonly BackgroundTaskQueue _queue;
    private readonly IServiceScopeFactory _scopeFactory;
    private readonly ILogger<BackgroundTaskWorker> _logger;

    public BackgroundTaskWorker(
        BackgroundTaskQueue queue,
        IServiceScopeFactory scopeFactory,
        ILogger<BackgroundTaskWorker> logger)
    {
        _queue = queue;
        _scopeFactory = scopeFactory;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            Func<IServiceProvider, CancellationToken, Task> task;
            try
            {
                task = await _queue.DequeueAsync(stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }

            try
            {
                using var scope = _scopeFactory.CreateScope();
                await task(scope.ServiceProvider, stoppingToken);
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                _logger.LogInformation("A queued automation task was cancelled during application shutdown.");
                break;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "A queued automation task failed outside its task handler.");
            }
        }
    }
}
