namespace AfsPdfComparison.Services;

internal sealed record AiProviderAttempt<T>(
    string Name,
    Func<CancellationToken, Task<T>> ExecuteAsync);

internal sealed record AiFailoverOutcome<T>(
    IReadOnlyList<T> CompletedResults,
    IReadOnlyList<T> ReadyResults);

internal static class AiFailoverExecutor
{
    public static async Task<AiFailoverOutcome<T>> ExecuteAsync<T>(
        IEnumerable<AiProviderAttempt<T>> attempts,
        Func<T, string?> getStatus,
        int readyResultTarget,
        TimeSpan waitAfterFirstReady,
        CancellationToken cancellationToken = default)
    {
        var providers = attempts.ToList();
        if (providers.Count == 0)
            return new AiFailoverOutcome<T>(Array.Empty<T>(), Array.Empty<T>());

        using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var pending = providers.ToDictionary(
            attempt => attempt.Name,
            attempt => attempt.ExecuteAsync(linkedCts.Token));

        var completed = new List<T>(providers.Count);
        var ready = new List<T>(providers.Count);
        bool firstReadySeen = false;
        Task? graceWindow = null;

        while (pending.Count > 0)
        {
            var waitSet = pending.Values.Cast<Task>().ToList();
            if (graceWindow != null)
                waitSet.Add(graceWindow);

            Task finished = await Task.WhenAny(waitSet);
            if (graceWindow != null && ReferenceEquals(finished, graceWindow))
                break;

            var finishedKeys = pending
                .Where(entry => entry.Value.IsCompleted)
                .Select(entry => entry.Key)
                .ToList();

            foreach (string key in finishedKeys)
            {
                Task<T> task = pending[key];
                pending.Remove(key);

                T result = await task;
                completed.Add(result);

                if (!string.Equals(getStatus(result), "ready", StringComparison.OrdinalIgnoreCase))
                    continue;

                ready.Add(result);

                if (!firstReadySeen)
                {
                    firstReadySeen = true;

                    if (waitAfterFirstReady <= TimeSpan.Zero || ready.Count >= readyResultTarget)
                    {
                        linkedCts.Cancel();
                        return new AiFailoverOutcome<T>(completed, ready);
                    }

                    graceWindow = Task.Delay(waitAfterFirstReady);
                }
                else if (ready.Count >= readyResultTarget)
                {
                    linkedCts.Cancel();
                    return new AiFailoverOutcome<T>(completed, ready);
                }
            }
        }

        linkedCts.Cancel();
        return new AiFailoverOutcome<T>(completed, ready);
    }
}
