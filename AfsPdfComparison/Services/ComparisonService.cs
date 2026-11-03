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
    private readonly AutoReviewEngine _autoReview;
    private readonly TableDiffService _tableDiff;

    public ComparisonService(
        PageAlignmentService  aligner,
        LineComparatorService comparator,
        SmartComparisonOrchestrator smartOrchestrator,
        DocumentIntelligenceService documentIntelligence,
        AutoReviewEngine autoReview,
        TableDiffService tableDiff)
    {
        _aligner              = aligner;
        _comparator           = comparator;
        _smartOrchestrator    = smartOrchestrator;
        _documentIntelligence = documentIntelligence;
        _autoReview           = autoReview;
        _tableDiff            = tableDiff;
    }

    /// <summary>
    /// Runs the full comparison pipeline: page alignment → full-document diff →
    /// number diff → page-level diff.
    /// </summary>
    /// <param name="report1">Baseline AFS (first uploaded).</param>
    /// <param name="report2">Comparand AFS (second uploaded).</param>
    /// <param name="includeDeferredInsights">
    /// When true, also computes the heavier semantic section analyses.
    /// The Step 4 UI does not need those to render, so the fast path leaves them out.
    /// </param>
    /// <returns>Populated <see cref="ComparisonResult"/>.</returns>
    public async Task<ComparisonResult> BuildComparisonAsync(
        PdfReport report1,
        PdfReport report2,
        CancellationToken cancellationToken = default,
        bool includeDeferredInsights = false)
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

        // ── 4.4b  Cross-page move detection ───────────────────────────────
        // Reclassifies removed+added pairs that match as "moved" to reduce noise.
        DetectMovedContent(fullDiff, pageDiffs);
        Console.WriteLine($"[Compare] Move detection completed in {sw.ElapsedMilliseconds} ms");

        // ── 4.5  Count aggregation ─────────────────────────────────────────
        var counts = fullDiff
            .GroupBy(d => d.Status is "same" or "same-cross-page" ? "same" : d.Status)
            .ToDictionary(g => g.Key, g => g.Count());
        SmartComparisonResult? smartResult = null;
        DocumentComparisonReport? documentResult = null;

        if (includeDeferredInsights)
        {
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
        }
        else
        {
            Console.WriteLine($"[Compare] Deferred semantic analysis skipped; total pipeline {sw.ElapsedMilliseconds} ms");
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

        // ── 4.6  Auto-review engine (deterministic — no AI required) ───────
        try
        {
            var (autoFindings, labeledChanges) = _autoReview.GenerateFindings(result);
            result.AutoReview            = autoFindings;
            result.LabeledNumberChanges  = labeledChanges;
            Console.WriteLine($"[Compare] Auto-review generated {autoFindings.Count} finding(s) in {sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Compare] Auto-review failed (non-fatal): {ex.Message}");
        }

        // ── 4.7  Table cell-level diff (#3) ───────────────────────────────
        try
        {
            result.TableCellDiffs = _tableDiff.Diff(result);
            Console.WriteLine($"[Compare] Table diff: {result.TableCellDiffs.Count} cell change(s) in {sw.ElapsedMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Compare] Table diff failed (non-fatal): {ex.Message}");
        }

        Console.WriteLine($"[Compare] Core comparison completed in {sw.ElapsedMilliseconds} ms");
        // AI analysis runs separately via /api/ai-analyze — does not block this response
        return result;
    }

    // ── Move detection ────────────────────────────────────────────────────────
    private static void DetectMovedContent(List<LineDiffResult> fullDiff, List<PageDiffResult> pageDiffs)
    {
        // Only operate on non-trivially short lines
        var removed = fullDiff
            .Where(d => d.Status == "removed" && (d.Line1?.Length ?? 0) >= 30)
            .ToList();
        var added = fullDiff
            .Where(d => d.Status == "added" && (d.Line2?.Length ?? 0) >= 30)
            .ToList();

        if (removed.Count == 0 || added.Count == 0) return;

        // O(n²) but with short-circuit — only matters for large docs
        // Cap at 200 candidates each to avoid performance issues
        var removedSample = removed.Take(200).ToList();
        var addedSample   = added.Take(200).ToList();

        var reclassifiedAdded = new HashSet<LineDiffResult>(ReferenceEqualityComparer.Instance);

        foreach (var rem in removedSample)
        {
            string norm1 = Normalise(rem.Line1 ?? "");
            if (norm1.Length < 20) continue;

            foreach (var add in addedSample)
            {
                if (reclassifiedAdded.Contains(add)) continue;
                string norm2 = Normalise(add.Line2 ?? "");
                if (norm2.Length < 20) continue;

                double sim = PageAlignmentService.SequenceSimilarity(norm1, norm2);
                if (sim >= 0.94)
                {
                    rem.Status = "moved";
                    add.Status = "moved";
                    reclassifiedAdded.Add(add);
                    break;
                }
            }
        }

        // Recompute per-page issue counts after reclassification
        foreach (var pd in pageDiffs)
        {
            pd.Changed = pd.Diffs.Count(d => d.Status == "changed");
            pd.Added   = pd.Diffs.Count(d => d.Status == "added");
            pd.Removed = pd.Diffs.Count(d => d.Status == "removed");
            pd.Same    = pd.Diffs.Count(d => d.Status == "same" || d.Status == "same-cross-page");
        }
    }

    private static string Normalise(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        return System.Text.RegularExpressions.Regex.Replace(s.ToLowerInvariant().Trim(), @"\s+", " ");
    }

}
