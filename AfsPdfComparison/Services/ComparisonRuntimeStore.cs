using System.Collections.Concurrent;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

/// <summary>
/// Keeps the full comparison result in server memory for the active session so
/// the app does not need to serialize heavy diff payloads into ISession.
/// </summary>
public class ComparisonRuntimeStore
{
    private readonly ConcurrentDictionary<string, ComparisonResult> _comparisons = new();

    public void Set(string sessionId, ComparisonResult comparison)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            return;

        _comparisons[sessionId] = comparison;
    }

    public ComparisonResult? Get(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            return null;

        return _comparisons.TryGetValue(sessionId, out var comparison)
            ? comparison
            : null;
    }

    public void Remove(string sessionId)
    {
        if (string.IsNullOrWhiteSpace(sessionId))
            return;

        _comparisons.TryRemove(sessionId, out _);
    }
}
