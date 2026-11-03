using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

/// <summary>
/// Orchestrates the full smart comparison pipeline:
/// extract -> semantic blocks -> smart diff.
/// </summary>
public class SmartComparisonOrchestrator
{
    private readonly SmartDiffEngine _engine = new();

    /// <summary>
    /// Compare two extracted PdfReports using semantic block comparison.
    /// Returns a structured list of meaningful changes.
    /// </summary>
    public SmartComparisonResult Compare(PdfReport report1, PdfReport report2)
    {
        var blocks1 = BuildAllBlocks(report1);
        var blocks2 = BuildAllBlocks(report2);

        var diffs = _engine.Compare(blocks1, blocks2);

        var materialNumeric = diffs
            .Where(d => d.ChangeType == SmartChangeType.NumericChange && d.IsMaterial)
            .ToList();
        var narrativeRewrites = diffs
            .Where(d => d.ChangeType == SmartChangeType.NarrativeRewrite)
            .ToList();
        var lineChanges = diffs
            .Where(d => d.ChangeType == SmartChangeType.LineChanged)
            .ToList();
        var added = diffs
            .Where(d => d.ChangeType == SmartChangeType.BlockAdded)
            .ToList();
        var removed = diffs
            .Where(d => d.ChangeType == SmartChangeType.BlockRemoved)
            .ToList();

        return new SmartComparisonResult
        {
            Report1Label = report1.Label,
            Report2Label = report2.Label,
            AllDiffs = diffs,
            MaterialNumeric = materialNumeric,
            NarrativeRewrites = narrativeRewrites,
            LineChanges = lineChanges,
            Added = added,
            Removed = removed,
            TotalBlocks1 = blocks1.Count,
            TotalBlocks2 = blocks2.Count,
            ComparedAt = DateTime.UtcNow
        };
    }

    private static List<SemanticBlock> BuildAllBlocks(PdfReport report)
    {
        var pageResults = new List<SemanticBlock>[report.Pages.Count];
        System.Threading.Tasks.Parallel.For(0, report.Pages.Count, i =>
            pageResults[i] = SemanticBlockBuilder.Build(report.Pages[i], i + 1));
        return pageResults.SelectMany(b => b).ToList();
    }
}

public class SmartComparisonResult
{
    public string Report1Label { get; set; } = "";
    public string Report2Label { get; set; } = "";
    public List<SmartDiff> AllDiffs { get; set; } = new();
    public List<SmartDiff> MaterialNumeric { get; set; } = new();
    public List<SmartDiff> NarrativeRewrites { get; set; } = new();
    public List<SmartDiff> LineChanges { get; set; } = new();
    public List<SmartDiff> Added { get; set; } = new();
    public List<SmartDiff> Removed { get; set; } = new();
    public int TotalBlocks1 { get; set; }
    public int TotalBlocks2 { get; set; }
    public DateTime ComparedAt { get; set; }

    public int TotalChanges => AllDiffs.Count;
    public int MaterialCount => MaterialNumeric.Count + NarrativeRewrites.Count;
}
