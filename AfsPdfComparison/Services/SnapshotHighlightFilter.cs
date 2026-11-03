using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

public static class SnapshotHighlightFilter
{
    private static readonly HashSet<string> HighlightStatuses = new(StringComparer.OrdinalIgnoreCase)
    {
        "changed",
        "added",
        "removed",
    };

    private static readonly HashSet<string> SuppressStatuses = new(StringComparer.OrdinalIgnoreCase)
    {
        "same",
        "same-cross-page",
        "same-different-page",
    };

    public static List<string> Afs1HighlightLines(PageDiffResult pageDiff) =>
        pageDiff.Diffs
            .Where(d => ShouldHighlight(d.Status) && !string.IsNullOrWhiteSpace(d.Line1))
            .Select(d => d.Line1)
            .ToList();

    public static List<string> Afs2HighlightLines(PageDiffResult pageDiff) =>
        pageDiff.Diffs
            .Where(d => ShouldHighlight(d.Status) && !string.IsNullOrWhiteSpace(d.Line2))
            .Select(d => d.Line2)
            .ToList();

    public static bool ShouldHighlight(string status) =>
        HighlightStatuses.Contains(status ?? string.Empty);

    public static bool IsSuppressed(string status) =>
        SuppressStatuses.Contains(status ?? string.Empty);

    public static (int Same, int Changed, int Added, int Removed) SummaryCounts(PageDiffResult pageDiff)
    {
        int same = pageDiff.Diffs.Count(d => !ShouldHighlight(d.Status));
        int changed = pageDiff.Diffs.Count(d => d.Status == "changed");
        int added = pageDiff.Diffs.Count(d => d.Status == "added");
        int removed = pageDiff.Diffs.Count(d => d.Status == "removed");
        return (same, changed, added, removed);
    }
}
