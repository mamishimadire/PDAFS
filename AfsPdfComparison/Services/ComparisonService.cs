// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3                                                               ║
// ║                                                                              ║
// ║  SERVICE — ComparisonService                                                 ║
// ║  Orchestrates extraction, alignment, and comparison into one result.         ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • Python equivalent: _build_comparison() in the notebook                  ║
// ║   • Calls PageAlignmentService and LineComparatorService in sequence         ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Diagnostics;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 4 · COMPARISON SERVICE
//
// Responsibility: glue code that calls the three sub-services and assembles
// a ComparisonResult.  Keeps the controller thin.
//
// Processing order (mirrors Python _build_comparison):
//   1. Build smart page alignment map
//   2. Run full-document line diff
//   3. Run number comparison
//   4. Run page-by-page aligned diff
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Orchestrates the full AFS comparison pipeline and returns a
/// <see cref="ComparisonResult"/>.
/// Registered as a singleton in DI.
/// </summary>
public class ComparisonService
{
    private readonly PageAlignmentService  _aligner;
    private readonly LineComparatorService _comparator;
    private readonly SmartComparisonOrchestrator _smartOrchestrator;
    private readonly DocumentIntelligenceService _documentIntelligence;
    public ComparisonService(
        PageAlignmentService  aligner,
        LineComparatorService comparator,
        SmartComparisonOrchestrator smartOrchestrator,
        DocumentIntelligenceService documentIntelligence)
    {
        _aligner              = aligner;
        _comparator           = comparator;
        _smartOrchestrator    = smartOrchestrator;
        _documentIntelligence = documentIntelligence;
    }

    /// <summary>
    /// Runs the full comparison pipeline: page alignment → full-document diff →
    /// number diff → page-level diff.
    /// </summary>
    /// <param name="report1">Baseline AFS (first uploaded).</param>
    /// <param name="report2">Comparand AFS (second uploaded).</param>
    /// <returns>Populated <see cref="ComparisonResult"/>.</returns>
    public async Task<ComparisonResult> BuildComparisonAsync(PdfReport report1, PdfReport report2, CancellationToken cancellationToken = default)
    {
        var sw = Stopwatch.StartNew();
        // ── 4.1  Smart page alignment ──────────────────────────────────────
        // Builds a map from AFS 1 pages to best-matching AFS 2 pages.
        // Pages that find no match (score < 0.30) are left as −1.
        var fullAlignment = _aligner.BuildAlignment(report1.Pages, report2.Pages);
        Console.WriteLine($"[Compare] Alignment completed in {sw.ElapsedMilliseconds} ms");

        // ── 4.2  Full-document line diff ───────────────────────────────────
        // Compares the concatenated full text of both documents.
        // This is no longer used for the primary text diff counts.
        // Overall text differences are built from the aligned page pairs below
        // so that moved content is compared by meaning, not by raw page order.

        // ── 4.3  Number comparison ─────────────────────────────────────────
        // Jaccard similarity on the numeric token sets of both full texts.
        var numCmp    = _comparator.CompareNumbers(report1.FullText, report2.FullText);
        Console.WriteLine($"[Compare] Number comparison completed in {sw.ElapsedMilliseconds} ms");

        // ── 4.4  Page-aligned diff ─────────────────────────────────────────
        // Detailed line diff per matched page pair, used for the visual snapshot viewer.
        var pageDiffs = _comparator.CompareAlignedPages(report1, report2, fullAlignment);
        var fullDiff  = pageDiffs
            .Where(p => !p.SecondaryMatch)
            .SelectMany(p => p.Diffs)
            .ToList();
        Console.WriteLine($"[Compare] Page diffs completed in {sw.ElapsedMilliseconds} ms");

        // ── 4.5  Count aggregation ─────────────────────────────────────────
        var counts = fullDiff
            .GroupBy(d => d.Status is "same" or "same-cross-page" ? "same" : d.Status)
            .ToDictionary(g => g.Key, g => g.Count());
        SmartComparisonResult? smartResult = null;
        DocumentComparisonReport? documentResult = null;

        try
        {
            smartResult = _smartOrchestrator.Compare(report1, report2);
            Console.WriteLine($"[Compare] Smart comparison completed in {sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Compare] Smart comparison failed: {ex.Message}");
        }

        try
        {
            documentResult = _documentIntelligence.Compare(report1, report2);
            Console.WriteLine($"[Compare] Document intelligence completed in {sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Compare] Document intelligence failed: {ex.Message}");
        }

        var result = new ComparisonResult
        {
            Report1   = report1,
            Report2   = report2,
            FullDiff  = fullDiff,
            Counts    = counts,
            NumCmp    = numCmp,
            PageDiffs = pageDiffs,
            Alignment = fullAlignment,
            SmartComparison = smartResult,
            DocumentIntelligence = documentResult,
            ComparedAt = DateTime.UtcNow,
        };

        Console.WriteLine($"[Compare] Deterministic comparison completed in {sw.ElapsedMilliseconds} ms");
        // AI analysis runs separately via /api/ai-analyze — does not block this response
        return result;
    }

}
