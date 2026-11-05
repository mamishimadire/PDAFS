// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3.2                                                             ║
// ║                                                                              ║
// ║  SERVICE — LineComparatorService                                             ║
// ║  Core diff engine: line comparison, word diff, number extraction.            ║
// ║                                                                              ║
// ║  Content-aware status checks distinguish edits from extraction formatting.  ║
// ║  ROOT CAUSE: pdfplumber/PdfPig wraps long paragraphs at different column    ║
// ║  widths per PDF. The sentence:                                               ║
// ║    "used. The separate financial statements...IAS 27 Consolidated"           ║
// ║  in AFS1 continues in AFS2 as:                                               ║
// ║    "used. The separate financial statements...IAS 27 Consolidated            ║
// ║     and Separate Financial Statements."                                      ║
// ║  These score 0.87–0.91 — above SIM_RELATED(0.55) so they are matched,      ║
// ║  but below SIM_EXACT(0.94) so they were falsely flagged "changed".          ║
// ║  Wrap suppression requires an ordered prefix, not loose word overlap.       ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • Ratcliff/Obershelp string similarity — see PageAlignmentService         ║
// ║   • Python equivalent: LineComparator._determine_status() in notebook       ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 3 · LINE COMPARATOR SERVICE
//
// Responsibility: Given two text strings (page or full-document),
// produce a list of LineDiffResult objects describing every line's status.
//
// Algorithm overview:
//   1. Split both texts into non-noise lines.
//   2. Build an inverted-index (bigram + first-token) on AFS 1 lines.
//   3. For each AFS 2 line:
//      a. Retrieve candidate AFS 1 lines from the index (fast approximate lookup).
//      b. Score each candidate with Ratcliff/Obershelp + numeric bonus/penalty.
//      c. Accept the highest-scoring candidate if score ≥ SIM_RELATED (0.55).
//      d. Determine final status using the three-gate equality check.
//   4. AFS 1 lines that were never matched are flagged as "removed".
//
// Equality is based on normalized content. Similarity scores pair candidate
// lines; they do not by themselves prove the content is unchanged.
//
// Reference: Python LineComparator class in the notebook v4.3.2
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Performs line-level and number-level comparison between two AFS texts.
/// Registered as a singleton in DI.
/// </summary>
public class LineComparatorService
{
    // ── Similarity thresholds ────────────────────────────────────────────────
    private const double SimExact   = 0.96;  // raised from 0.94 — reduces false "changed" on near-identical paragraphs
    private const double SimRelated = 0.55;
    private const double GapPenalty = 0.24;
    private const int MaxAlignmentCells = 250_000;

    private const double LengthRatioThreshold = 0.55;

    private const int NeighbourContextLines = 8;

    // Regex: strip page-number stamps before number extraction.
    // Matches lines that are PURELY a page stamp (e.g. "- 40 -", "Page 3 of 10").
    // Reference: Python _PAGE_NUM_RE
    private static readonly Regex _pageStampRe = new(
        @"(?m)^[-–—\s]*\(?\s*\d{1,4}\s*\)?[-–—\s]*$",
        RegexOptions.Compiled | RegexOptions.Multiline);

    // #2 — Header/footer line detector.
    // Matches lines that appear as page decoration rather than audit content:
    //   "Page 3 of 10", "3 of 10", "(continued)", "...continued", copyright notices,
    //   confidentiality notices, standalone years (4-digit), and running-header
    //   short ALL-CAPS fragments (≤3 words, no digits — e.g. repeated title bars).
    private static readonly Regex _headerFooterRe = new(
        @"^(?:page\s+\d+\s+of\s+\d+|p\.\s*\d+|\d{1,4}\s+of\s+\d{1,4}|" +
        @"\(?continued\.?\)?|\.{2,}continued|continued\.{2,}|" +
        @"copyright\s*©|©\s*\d{4}|all\s+rights?\s+reserved|" +
        @"confidential(?:ity)?|for\s+official\s+use\s+only|" +
        @"unaudited|draft\s+only|prepared\s+by|" +
        @"(?:19|20)\d{2})$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // #10 — Number-format normalization: "R 1 234 567" and "R1,234,567" and "1234567" are the same.
    private static readonly Regex _rCurrencyRe   = new(@"\br\s*(?=\d)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex _saThousandsRe = new(@"(?<=\d) (?=\d{3}\b)", RegexOptions.Compiled);
    private static readonly Regex _commaDecimalRe = new(@"(\d),(\d\d?)(?!\d)", RegexOptions.Compiled);

    // Regex: extract numeric tokens (integers and decimals).
    // Handles South African space-thousands-separator format: "6 835" → 6835.
    // Only allows a SINGLE space between digit groups of exactly 3 digits
    // (i.e. real thousands separator, not arbitrary whitespace between numbers).
    private static readonly Regex _numTokenRe = new(
        @"\b\d{1,3}(?:[,\.]?\d{3})*(?:\.\d+)?\b|\b\d+ \d{3}(?:\.\d+)?\b", RegexOptions.Compiled);

    // Static compiled pattern for ExtractNumbers — must NOT be created inside the method
    // because RegexOptions.Compiled triggers IL compilation on every instantiation,
    // adding seconds of overhead when ExtractNumbers is called thousands of times per comparison.
    private static readonly Regex _extNumRe = new(
        @"\((\d{1,3}(?:[\s,]\d{3})*(?:\.\d+)?)\)" +
        @"|\b(\d{1,3}(?:[,\.]?\d{3})*(?:\.\d+)?)\b" +
        @"|\b(\d+\s\d{3}(?:\s\d{3})*(?:\.\d+)?)\b",
        RegexOptions.Compiled);
    private static readonly Regex _noteMarkerRe = new(
        @"^(?:note\s+)?\d+(?:\.\d+)*\.?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex _sectionHeadingLineRe = new(
        @"^(?:note\s+)?\d+(?:\.\d+)*\.?\s+[a-z]",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 3.1 · COMPARE LINES (main entry point)
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Compares every line in <paramref name="text2"/> against
    /// <paramref name="text1"/> and returns a diff result per line.
    /// Also appends "removed" entries for any AFS 1 lines that were not matched.
    /// </summary>
    public List<LineDiffResult> CompareLines(string text1, string text2)
    {
        var lines1 = SplitLines(text1);
        var lines2 = SplitLines(text2);
        long cells = (long)lines1.Count * Math.Max(1, lines2.Count);
        if (cells <= MaxAlignmentCells)
            return CompareLinesOrdered(lines1, lines2);

        return CompareLinesGreedy(lines1, lines2);
    }

    private List<LineDiffResult> CompareLinesGreedy(List<string> lines1, List<string> lines2)
    {
        var index1 = BuildIndex(lines1);   // inverted index over AFS 1

        var used1   = new HashSet<int>();
        var results = new List<LineDiffResult>(lines2.Count + lines1.Count);

        // ── Match each AFS 2 line ─────────────────────────────────────────
        for (int i2 = 0; i2 < lines2.Count; i2++)
        {
            string l2      = lines2[i2];
            var    cands   = Candidates(l2, index1, used1, lines1.Count);
            int    bestIdx = -1;
            double bestSc  = 0.0;

            foreach (int ci in cands)
            {
                double s = Score(l2, lines1[ci]);
                if (s > bestSc) { bestSc = s; bestIdx = ci; }
            }

            if (bestIdx >= 0 && bestSc >= SimRelated)
            {
                used1.Add(bestIdx);
                string l1     = lines1[bestIdx];
                string status = DetermineStatus(l1, l2);
                var    n1set  = ExtractNumbers(l1);
                var    n2set  = ExtractNumbers(l2);
                var    numDiff = status == "changed"
                    ? SymmetricExcept(n1set, n2set).OrderBy(x => x).ToList()
                    : new List<string>();

                results.Add(new LineDiffResult
                {
                    Status     = status,
                    Line1      = l1,
                    Line2      = l2,
                    Similarity = Math.Round(bestSc, 4),
                    NumDiff    = numDiff,
                    WordDiff   = status == "changed" ? WordDiff(l1, l2) : new(),
                });
            }
            else
            {
                if (IsIgnorableStandaloneLine(l2))
                    continue;

                results.Add(new LineDiffResult
                {
                    Status = "added", Line1 = "", Line2 = l2,
                    Similarity = 0.0,
                });
            }
        }

        // ── AFS 1 lines with no match → "removed" ─────────────────────────
        for (int i1 = 0; i1 < lines1.Count; i1++)
        {
            if (!used1.Contains(i1))
            {
                if (IsIgnorableStandaloneLine(lines1[i1]))
                    continue;

                results.Add(new LineDiffResult
                {
                    Status = "removed", Line1 = lines1[i1], Line2 = "",
                    Similarity = 0.0,
                });
            }
        }

        return results;
    }

    private List<LineDiffResult> CompareLinesOrdered(List<string> lines1, List<string> lines2)
    {
        int n1 = lines1.Count;
        int n2 = lines2.Count;
        var results = new List<LineDiffResult>(n1 + n2);

        if (n1 == 0 && n2 == 0) return results;

        var sim = new double[n1, n2];
        for (int i = 0; i < n1; i++)
            for (int j = 0; j < n2; j++)
                sim[i, j] = Score(lines1[i], lines2[j]);

        var dp = new double[n1 + 1, n2 + 1];
        var move = new byte[n1 + 1, n2 + 1]; // 0=up(remove), 1=left(add), 2=diag(match)

        for (int i = 1; i <= n1; i++)
        {
            dp[i, 0] = dp[i - 1, 0] - GapPenalty;
            move[i, 0] = 0;
        }

        for (int j = 1; j <= n2; j++)
        {
            dp[0, j] = dp[0, j - 1] - GapPenalty;
            move[0, j] = 1;
        }

        for (int i = 1; i <= n1; i++)
        {
            for (int j = 1; j <= n2; j++)
            {
                double best = double.NegativeInfinity;
                byte bestMove = 0;

                double s = sim[i - 1, j - 1];
                if (s >= SimRelated)
                {
                    best = dp[i - 1, j - 1] + s;
                    bestMove = 2;
                }

                double up = dp[i - 1, j] - GapPenalty;
                if (up > best)
                {
                    best = up;
                    bestMove = 0;
                }

                double left = dp[i, j - 1] - GapPenalty;
                if (left > best)
                {
                    best = left;
                    bestMove = 1;
                }

                dp[i, j] = best;
                move[i, j] = bestMove;
            }
        }

        var aligned = new List<(int? I1, int? I2, double Sim)>(n1 + n2);
        int ti = n1, tj = n2;
        while (ti > 0 || tj > 0)
        {
            if (ti > 0 && tj > 0 && move[ti, tj] == 2 && sim[ti - 1, tj - 1] >= SimRelated)
            {
                aligned.Add((ti - 1, tj - 1, sim[ti - 1, tj - 1]));
                ti--;
                tj--;
            }
            else if (tj > 0 && (ti == 0 || move[ti, tj] == 1))
            {
                aligned.Add((null, tj - 1, 0.0));
                tj--;
            }
            else
            {
                aligned.Add((ti - 1, null, 0.0));
                ti--;
            }
        }

        aligned.Reverse();

        foreach (var step in aligned)
        {
            if (step.I1.HasValue && step.I2.HasValue)
            {
                string l1 = lines1[step.I1.Value];
                string l2 = lines2[step.I2.Value];
                string status = DetermineStatus(l1, l2);
                var n1set = ExtractNumbers(l1);
                var n2set = ExtractNumbers(l2);
                var numDiff = status == "changed"
                    ? SymmetricExcept(n1set, n2set).OrderBy(x => x).ToList()
                    : new List<string>();

                results.Add(new LineDiffResult
                {
                    Status = status,
                    Line1 = l1,
                    Line2 = l2,
                    Similarity = Math.Round(step.Sim, 4),
                    NumDiff = numDiff,
                    WordDiff = status == "changed" ? WordDiff(l1, l2) : new(),
                });
            }
            else if (step.I2.HasValue)
            {
                string l2 = lines2[step.I2.Value];
                if (IsIgnorableStandaloneLine(l2))
                    continue;

                results.Add(new LineDiffResult
                {
                    Status = "added",
                    Line1 = "",
                    Line2 = l2,
                    Similarity = 0.0,
                });
            }
            else if (step.I1.HasValue)
            {
                string l1 = lines1[step.I1.Value];
                if (IsIgnorableStandaloneLine(l1))
                    continue;

                results.Add(new LineDiffResult
                {
                    Status = "removed",
                    Line1 = l1,
                    Line2 = "",
                    Similarity = 0.0,
                });
            }
        }

        return results;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 3.2 · COMPARE NUMBERS
    //
    // Extracts all numeric tokens from both full-document texts and computes
    // a Jaccard similarity on the resulting sets.
    // Reference: Python LineComparator.compare_numbers()
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Extracts and compares numeric tokens from two full-document texts.
    /// </summary>
    public NumberComparisonResult CompareNumbers(string text1, string text2)
    {
        var n1 = ExtractNumbers(text1);
        var n2 = ExtractNumbers(text2);

        var inBoth   = n1.Intersect(n2).OrderBy(x => double.Parse(x, CultureInfo.InvariantCulture)).ToList();
        var only1    = n1.Except(n2)   .OrderBy(x => double.Parse(x, CultureInfo.InvariantCulture)).ToList();
        var only2    = n2.Except(n1)   .OrderBy(x => double.Parse(x, CultureInfo.InvariantCulture)).ToList();
        var allUnion = n1.Union(n2);
        int unionCnt = allUnion.Count();
        double sim   = unionCnt == 0 ? 100.0
                       : Math.Round((double)inBoth.Count / unionCnt * 100, 1);

        return new NumberComparisonResult
        {
            InBoth       = inBoth,
            OnlyInAfs1   = only1,
            OnlyInAfs2   = only2,
            CountAfs1    = n1.Count,
            CountAfs2    = n2.Count,
            SimilarityPct = sim,
        };
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 3.3 · PAGE-ALIGNED COMPARISON
    //
    // Compare page pairs identified by the PageAlignmentService.
    // Produces one PageDiffResult per aligned pair plus one per unmatched page.
    // Reference: Python LineComparator.compare_pages_aligned()
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Runs a line-by-line comparison for each aligned page pair.
    /// </summary>
    public List<PageDiffResult> CompareAlignedPages(
        PdfReport report1,
        PdfReport report2,
        List<AlignmentEntry> alignment)
    {
        var matchedEntries = alignment.Where(a => a.I2 >= 0).ToList();
        var alignedI1   = new HashSet<int>(matchedEntries.Select(e => e.I1));
        var alignedI2   = new HashSet<int>(matchedEntries.Select(e => e.I2));

        // Each page pair's diff is fully independent — compute all in parallel.
        var matchedResults = new PageDiffResult[matchedEntries.Count];
        System.Threading.Tasks.Parallel.For(0, matchedEntries.Count, idx =>
        {
            var entry = matchedEntries[idx];
            int i1 = entry.I1, i2 = entry.I2;
            string context1 = BuildNeighbourContext(report1.Pages, i1, NeighbourContextLines);
            string context2 = BuildNeighbourContext(report2.Pages, i2, NeighbourContextLines);
            var diff = CompareLinesGreedy(SplitLines(context1), SplitLines(context2));
            diff = SuppressBleedDiffsV2(diff, report1, report2, i1, i2);
            diff = MergeParagraphReflowArtefacts(diff);

            int same  = diff.Count(d => d.Status == "same" || d.Status == "same-cross-page");
            int chg   = diff.Count(d => d.Status == "changed");
            int add   = diff.Count(d => d.Status == "added");
            int rem   = diff.Count(d => d.Status == "removed");
            int total = Math.Max(diff.Count, 1);

            matchedResults[idx] = new PageDiffResult
            {
                PageAfs1   = i1 + 1,
                PageAfs2   = i2 + 1,
                AlignSim   = entry.Sim,
                MatchConfidence = entry.Sim,
                StructuralChangeType = entry.MatchType == "relocated" ? "section-relocation" : "none",
                Same       = same,
                Changed    = chg,
                Added      = add,
                Removed    = rem,
                TotalLines = diff.Count,
                PctSame    = Math.Round((double)same / total * 100, 1),
                Diffs      = diff,
            };
        });

        // Preserve AFS 1 page order and assign sequential pair indices.
        var results = matchedResults.OrderBy(r => r.PageAfs1).ToList();
        MarkSystematicPaginationShift(results);
        int pairIndex = 0;
        foreach (var r in results) r.PairIndex = pairIndex++;

        for (int i1 = 0; i1 < report1.Pages.Count; i1++)
        {
            if (alignedI1.Contains(i1)) continue;

            var diff = BuildUnmatchedDiffs(report1.Pages[i1], "removed");

            results.Add(new PageDiffResult
            {
                PairIndex  = pairIndex++,
                PageAfs1   = i1 + 1,
                PageAfs2   = null,
                AlignSim   = 0.0,
                Removed    = diff.Count,
                TotalLines = diff.Count,
                PctSame    = 0.0,
                Diffs      = diff,
                Unmatched  = true,
            });
        }

        for (int i2 = 0; i2 < report2.Pages.Count; i2++)
        {
            if (alignedI2.Contains(i2)) continue;

            var diff = BuildUnmatchedDiffs(report2.Pages[i2], "added");

            results.Add(new PageDiffResult
            {
                PairIndex  = pairIndex++,
                PageAfs1   = null,
                PageAfs2   = i2 + 1,
                AlignSim   = 0.0,
                Added      = diff.Count,
                TotalLines = diff.Count,
                PctSame    = 0.0,
                Diffs      = diff,
                Unmatched  = true,
            });
        }

        return results;
    }

    private static void MarkSystematicPaginationShift(List<PageDiffResult> matchedPages)
    {
        var offsets = matchedPages
            .Where(page => page.PageAfs1.HasValue && page.PageAfs2.HasValue &&
                           page.StructuralChangeType == "none")
            .GroupBy(page => page.PageAfs2!.Value - page.PageAfs1!.Value)
            .Select(group => new { Offset = group.Key, Count = group.Count() })
            .OrderByDescending(group => group.Count)
            .ThenBy(group => Math.Abs(group.Offset))
            .FirstOrDefault();

        if (offsets == null || offsets.Offset == 0 || offsets.Count < 2 ||
            offsets.Count < Math.Ceiling(matchedPages.Count * 0.60))
            return;

        foreach (var page in matchedPages.Where(page =>
                     page.PageAfs1.HasValue && page.PageAfs2.HasValue &&
                     page.StructuralChangeType == "none" &&
                     page.PageAfs2.Value - page.PageAfs1.Value == offsets.Offset))
        {
            page.StructuralChangeType = "pagination-shift";
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 3-INTERNAL · HELPERS
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Post-processing pass that detects paragraph-reflow false positives.
    ///
    /// When two PDFs use different font sizes or column widths, the same paragraph
    /// text is broken into lines at different points. The line-by-line comparison
    /// marks these as "changed" because individual lines differ. But when consecutive
    /// changed line pairs are joined and their pure alphanumeric content compared,
    /// they are identical — a reflow artefact, not a real change.
    ///
    /// Safety guards:
    ///   • Only merges narrative runs (content must contain letters — not number-only rows).
    ///   • Does NOT merge when numbers differ (financial table values are preserved).
    ///   • Requires ≥ 20 chars of pure content to avoid merging trivial noise.
    ///   • Checks runs of 2–6 consecutive changed pairs.
    /// </summary>
    private static List<LineDiffResult> MergeParagraphReflowArtefacts(List<LineDiffResult> diffs)
    {
        var result = new List<LineDiffResult>(diffs.Count);
        int i = 0;

        while (i < diffs.Count)
        {
            var d = diffs[i];

            // Only start a merge attempt on a "changed" pair with both sides filled.
            if (d.Status != "changed" ||
                string.IsNullOrEmpty(d.Line1) ||
                string.IsNullOrEmpty(d.Line2))
            {
                result.Add(d);
                i++;
                continue;
            }

            // Scan ahead up to 12 positions, tolerating up to 2 "same" lines
            // between "changed" lines — font-size reflow produces mixed runs.
            int runEnd      = i;
            int sameSkipped = 0;
            for (int j = i + 1; j < Math.Min(i + 12, diffs.Count); j++)
            {
                var next = diffs[j];
                if (next.Status == "changed" &&
                    !string.IsNullOrEmpty(next.Line1) &&
                    !string.IsNullOrEmpty(next.Line2))
                {
                    runEnd = j;
                    sameSkipped = 0;
                }
                else if ((next.Status == "same" || next.Status == "same-cross-page")
                         && sameSkipped < 2)
                {
                    sameSkipped++;   // allow a short gap without breaking the run
                }
                else
                {
                    break;
                }
            }

            if (runEnd > i)
            {
                // Include everything from i to runEnd (changed + skipped sames).
                var run = diffs.GetRange(i, runEnd - i + 1);
                string joinedL1 = string.Join(" ", run
                    .Where(r => !string.IsNullOrEmpty(r.Line1)).Select(r => r.Line1));
                string joinedL2 = string.Join(" ", run
                    .Where(r => !string.IsNullOrEmpty(r.Line2)).Select(r => r.Line2));

                // Strip everything except word chars (letters + digits + underscore).
                string pure1 = Regex.Replace(Normalise(joinedL1), @"[^\w]", "");
                string pure2 = Regex.Replace(Normalise(joinedL2), @"[^\w]", "");

                // Must be substantially alphabetic (not a financial number block)
                // and the joined content must be identical.
                bool hasAlpha   = pure1.Any(char.IsLetter);
                bool longEnough = pure1.Length >= 20;
                bool matches    = pure1 == pure2;

                if (hasAlpha && longEnough && matches)
                {
                    foreach (var rd in run)
                        result.Add(new LineDiffResult
                        {
                            Status     = "same",
                            Line1      = rd.Line1,
                            Line2      = rd.Line2,
                            Similarity = 1.0,
                            NumDiff    = new List<string>(),
                            WordDiff   = new List<WordDiffToken>(),
                        });
                    i = runEnd + 1;
                    continue;
                }
            }

            result.Add(d);
            i++;
        }

        return result;
    }

    private static string BuildNeighbourContext(
        List<string> pages,
        int pageIndex,
        int fringeLines)
    {
        var sb = new StringBuilder();

        if (pageIndex > 0)
        {
            var prevLines = SplitLines(pages[pageIndex - 1]);
            foreach (var line in prevLines.TakeLast(fringeLines))
                sb.AppendLine(line);
        }

        sb.Append(pages[pageIndex]);

        if (pageIndex < pages.Count - 1)
        {
            var nextLines = SplitLines(pages[pageIndex + 1]);
            foreach (var line in nextLines.Take(fringeLines))
                sb.AppendLine(line);
        }

        return sb.ToString();
    }

    private static List<LineDiffResult> SuppressBleedDiffs(
        List<LineDiffResult> diffs,
        PdfReport report1,
        PdfReport report2,
        int pageIdx1,
        int pageIdx2)
    {
        var pool2 = BuildAdjacentLinePool(report2.Pages, pageIdx2, neighbourRadius: 2)
            .Select(BuildBleedCandidate)
            .Where(c => c.Norm.Length >= 6)
            .ToList();

        var pool1 = BuildAdjacentLinePool(report1.Pages, pageIdx1, neighbourRadius: 2)
            .Select(BuildBleedCandidate)
            .Where(c => c.Norm.Length >= 6)
            .ToList();

        var pool2Norm = pool2.Select(c => c.Norm).Where(s => s.Length >= 15).ToHashSet(StringComparer.Ordinal);
        var pool2Canon = pool2.Select(c => c.Canon).Where(s => s.Length >= 15).ToHashSet(StringComparer.Ordinal);
        var pool2Flat = pool2.Select(c => c.Flat).Where(s => s.Length >= 15).ToHashSet(StringComparer.Ordinal);
        var pool1Norm = pool1.Select(c => c.Norm).Where(s => s.Length >= 15).ToHashSet(StringComparer.Ordinal);
        var pool1Canon = pool1.Select(c => c.Canon).Where(s => s.Length >= 15).ToHashSet(StringComparer.Ordinal);
        var pool1Flat = pool1.Select(c => c.Flat).Where(s => s.Length >= 15).ToHashSet(StringComparer.Ordinal);

        foreach (var diff in diffs)
        {
            if (diff.Status == "removed" && !string.IsNullOrWhiteSpace(diff.Line1))
            {
                string norm = Normalise(diff.Line1);
                if (norm.Length < 15)
                    continue;
                string canon = Canonical(diff.Line1);
                string flat = WhitespaceAgnosticCanonical(diff.Line1);
                if (pool2Norm.Contains(norm) || pool2Canon.Contains(canon) || pool2Flat.Contains(flat))
                {
                    diff.Status = "same-cross-page";
                    diff.CrossPageNote = "Content found on adjacent AFS2 page — page layout differs between drafts";
                }
            }
            else if (diff.Status == "added" && !string.IsNullOrWhiteSpace(diff.Line2))
            {
                string norm = Normalise(diff.Line2);
                if (norm.Length < 15)
                    continue;
                string canon = Canonical(diff.Line2);
                string flat = WhitespaceAgnosticCanonical(diff.Line2);
                if (pool1Norm.Contains(norm) || pool1Canon.Contains(canon) || pool1Flat.Contains(flat))
                {
                    diff.Status = "same-cross-page";
                    diff.CrossPageNote = "Content found on adjacent AFS1 page — page layout differs between drafts";
                }
            }
        }

        return diffs;
    }

    private static List<string> BuildAdjacentLinePool(
        List<string> pages,
        int centerIdx,
        int neighbourRadius,
        int excludedPageIdx = -1)
    {
        var pool = new List<string>();

        for (int delta = -neighbourRadius; delta <= neighbourRadius; delta++)
        {
            int idx = centerIdx + delta;
            if (idx < 0 || idx >= pages.Count || idx == excludedPageIdx)
                continue;

            pool.AddRange(SplitLines(pages[idx]));
        }

        return pool;
    }

    private sealed class BleedCandidate
    {
        public string Norm { get; init; } = string.Empty;
        public string Canon { get; init; } = string.Empty;
        public string Flat { get; init; } = string.Empty;
    }

    private static BleedCandidate BuildBleedCandidate(string line)
    {
        string norm = Normalise(line);
        return new BleedCandidate
        {
            Norm = norm,
            Canon = Canonical(line),
            Flat = WhitespaceAgnosticCanonical(line),
        };
    }

    private static List<LineDiffResult> SuppressBleedDiffsV2(
        List<LineDiffResult> diffs,
        PdfReport report1,
        PdfReport report2,
        int pageIdx1,
        int pageIdx2)
    {
        var pool2 = BuildAdjacentLinePool(
                report2.Pages,
                pageIdx2,
                neighbourRadius: 2,
                excludedPageIdx: pageIdx2)
            .Select(BuildBleedCandidate)
            .Where(c => c.Norm.Length >= 6)
            .ToList();
        var pool1 = BuildAdjacentLinePool(
                report1.Pages,
                pageIdx1,
                neighbourRadius: 2,
                excludedPageIdx: pageIdx1)
            .Select(BuildBleedCandidate)
            .Where(c => c.Norm.Length >= 6)
            .ToList();

        foreach (var diff in diffs)
        {
            if (diff.Status == "removed" && !string.IsNullOrWhiteSpace(diff.Line1))
            {
                if (LooksLikeCrossPageEquivalent(diff.Line1, pool2))
                {
                    diff.Status = "same-cross-page";
                    diff.CrossPageNote = "Content found on adjacent AFS2 page - page layout differs between drafts";
                }
            }
            else if (diff.Status == "added" && !string.IsNullOrWhiteSpace(diff.Line2))
            {
                if (LooksLikeCrossPageEquivalent(diff.Line2, pool1))
                {
                    diff.Status = "same-cross-page";
                    diff.CrossPageNote = "Content found on adjacent AFS1 page - page layout differs between drafts";
                }
            }
        }

        return diffs;
    }

    private static bool LooksLikeCrossPageEquivalent(string line, List<BleedCandidate> pool)
    {
        string norm = Normalise(line);
        if (norm.Length < 6)
            return false;

        string canon = Canonical(line);
        return pool.Any(candidate =>
            candidate.Norm == norm ||
            candidate.Canon == canon);
    }

    // Split text into non-noise lines.
    // Paragraph reconstruction: only merge a line whose first char is lowercase
    // onto the preceding line — these are genuine mid-sentence wraps (e.g. the
    // word "the" at the end of a column causes the next word to spill onto a new
    // line starting with a lowercase letter).  Lines beginning with a capital
    // letter are always new sentences, headings, or structural items and must
    // NOT be merged — merging them produced false "same" verdicts by joining
    // e.g. "Onerous contracts" with the following provision paragraph.
    // Consecutive identical lines (caused by hyperlink duplication in PdfPig)
    // are deduplicated.
    private static List<string> SplitLines(string text)
    {
        var raw = (text ?? "").Split('\n')
                              .Select(l => l.Trim())
                              .Where(l => !PageAlignmentService.IsNoise(l))
                              .ToList();

        if (raw.Count == 0) return raw;

        var merged = new List<string>(raw.Count) { raw[0] };
        for (int i = 1; i < raw.Count; i++)
        {
            string prev   = merged[^1];
            string curr   = raw[i];
            char   last   = prev[^1];
            bool prevEnds = last is '.' or '!' or '?' or ';' or ':';
            bool currCont = curr.Length > 0 && char.IsLower(curr[0]);

            // Dedup: skip exact consecutive duplicate (hyperlink artefact)
            if (string.Equals(prev, curr, StringComparison.OrdinalIgnoreCase))
                continue;

            // Hyperlink sub-word dedup (adjacent pass):
            // PdfPig extracts a hyperlinked word as BOTH part of the containing
            // sentence (merged above) AND as an extra standalone annotation block.
            // Rule: if a short line (≤ 3 words) already appears verbatim as a
            // phrase inside the most recent merged line, discard it.
            // Space-padding both strings ensures whole-word matching without regex.
            if (curr.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length <= 3)
            {
                string nm = Normalise(curr);
                string np = Normalise(prev);
                if (!string.IsNullOrEmpty(nm) && (" " + np + " ").Contains(" " + nm + " "))
                    continue;
            }

            if (!prevEnds && currCont && prev.Length >= 15)
                merged[^1] = prev + " " + curr;
            else
                merged.Add(curr);
        }

        // Second-pass hyperlink artefact dedup (order-independent).
        // Some PDFs store hyperlink glyph data at a slightly different Y-coordinate
        // than the surrounding text, causing the Y-band grouper to emit the
        // hyperlinked word (e.g. "users") as a SEPARATE line that appears
        // BEFORE the containing sentence in the sorted order.  The adjacent-pass
        // above only fires when the orphan immediately follows the sentence, so
        // it misses the reversed order.  This second pass scans ALL lines:
        // any ≤3-word line that is a whole-word substring of ANY other line in
        // the same page is an orphan annotation duplicate and is removed.
        // Pre-compute normalised forms once to avoid O(n²) repeated Normalise calls.
        var mergedNormsDedup = merged.Select(Normalise).ToArray();
        var deduped = new List<string>(merged.Count);
        for (int k = 0; k < merged.Count; k++)
        {
            string line = merged[k];
            int wc = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
            if (wc > 0 && wc <= 3)
            {
                string nm = mergedNormsDedup[k];
                bool orphan = false;
                for (int j = 0; j < merged.Count; j++)
                {
                    if (j == k) continue;
                    if ((" " + mergedNormsDedup[j] + " ").Contains(" " + nm + " "))
                    {
                        orphan = true;
                        break;
                    }
                }
                if (orphan) continue;
            }
            deduped.Add(line);
        }
        return deduped
            .Where(line => !IsIgnorableStandaloneLine(line))
            .Where(line => !IsHeaderOrFooter(line))
            .ToList();
    }

    // #2 — Returns true for lines that are pure page decoration (header/footer artefacts).
    // Conservative: only strips clear-cut patterns so real content is never lost.
    private static bool IsHeaderOrFooter(string line)
    {
        string norm = Normalise(line);
        if (string.IsNullOrWhiteSpace(norm) || norm.Length > 120)
            return false;
        // Already handled by _pageStampRe / IsIgnorableStandaloneLine — belt-and-braces
        if (_headerFooterRe.IsMatch(norm))
            return true;
        // Short ALL-CAPS line with no digits and ≤5 words: likely a running header/title bar
        // but NOT a financial statement heading (those have lowercase words too)
        if (norm.Length <= 60 && norm == norm.ToUpperInvariant() &&
            !norm.Any(char.IsDigit) &&
            norm.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length <= 5 &&
            norm.All(c => char.IsLetter(c) || char.IsWhiteSpace(c) || c == '-' || c == '/' || c == '&'))
        {
            // Allowlist: short ALL-CAPS that are valid section headings
            // (we preserve these because auditors need them as section markers)
            return false;
        }
        return false;
    }

    private PageDiffResult? TryBuildSecondaryMatchFromAfs1(
        int pageAfs1Index,
        PdfReport report1,
        PdfReport report2,
        int pairIndex)
    {
        var sourceLines = SplitLines(report1.Pages[pageAfs1Index]);
        if (sourceLines.Count < 2)
            return null;

        PageDiffResult? best = null;
        double bestScore = 0.0;

        for (int i2 = 0; i2 < report2.Pages.Count; i2++)
        {
            var candidateLines = SplitLines(report2.Pages[i2]);
            var rawDiff = CompareLines(report1.Pages[pageAfs1Index], report2.Pages[i2]);
            var diff = rawDiff.Where(d => d.Status != "added").ToList();
            if (diff.Count == 0)
                continue;

            int same = diff.Count(d => d.Status == "same");
            int changed = diff.Count(d => d.Status == "changed");
            int removed = diff.Count(d => d.Status == "removed");
            int shared = same + changed;
            double score = (double)shared / Math.Max(1, diff.Count);
            double sourceCoverage = (double)shared / Math.Max(1, sourceLines.Count);
            int headingOverlap = SharedSectionHeadingCount(sourceLines, candidateLines);
            bool defaultMatch = shared >= 2 && score >= 0.72 && (same > 0 || changed >= 2);
            bool splitPageMatch =
                headingOverlap > 0 &&
                shared >= 6 &&
                sourceCoverage >= 0.60 &&
                (same >= 4 || changed >= 2);

            if (!defaultMatch && !splitPageMatch)
                continue;

            double candidateScore = defaultMatch
                ? score
                : Math.Max(score, sourceCoverage + Math.Min(0.10, headingOverlap * 0.05));

            if (candidateScore > bestScore)
            {
                bestScore = candidateScore;
                best = new PageDiffResult
                {
                    PairIndex = pairIndex,
                    PageAfs1 = pageAfs1Index + 1,
                    PageAfs2 = i2 + 1,
                    AlignSim = Math.Round(candidateScore, 3),
                    Same = same,
                    Changed = changed,
                    Added = 0,
                    Removed = removed,
                    TotalLines = diff.Count,
                    PctSame = Math.Round((double)same / Math.Max(1, diff.Count) * 100, 1),
                    Diffs = WithCrossPageNote(diff, pageAfs1Index + 1, i2 + 1),
                    SecondaryMatch = true,
                };
            }
        }

        return best;
    }

    private PageDiffResult? TryBuildSecondaryMatchFromAfs2(
        int pageAfs2Index,
        PdfReport report1,
        PdfReport report2,
        int pairIndex)
    {
        var sourceLines = SplitLines(report2.Pages[pageAfs2Index]);
        if (sourceLines.Count < 2)
            return null;

        PageDiffResult? best = null;
        double bestScore = 0.0;

        for (int i1 = 0; i1 < report1.Pages.Count; i1++)
        {
            var candidateLines = SplitLines(report1.Pages[i1]);
            var rawDiff = CompareLines(report1.Pages[i1], report2.Pages[pageAfs2Index]);
            var diff = rawDiff.Where(d => d.Status != "removed").ToList();
            if (diff.Count == 0)
                continue;

            int same = diff.Count(d => d.Status == "same");
            int changed = diff.Count(d => d.Status == "changed");
            int added = diff.Count(d => d.Status == "added");
            int shared = same + changed;
            double score = (double)shared / Math.Max(1, diff.Count);
            double sourceCoverage = (double)shared / Math.Max(1, sourceLines.Count);
            int headingOverlap = SharedSectionHeadingCount(sourceLines, candidateLines);
            bool defaultMatch = shared >= 2 && score >= 0.72 && (same > 0 || changed >= 2);
            bool splitPageMatch =
                headingOverlap > 0 &&
                shared >= 6 &&
                sourceCoverage >= 0.60 &&
                (same >= 4 || changed >= 2);

            if (!defaultMatch && !splitPageMatch)
                continue;

            double candidateScore = defaultMatch
                ? score
                : Math.Max(score, sourceCoverage + Math.Min(0.10, headingOverlap * 0.05));

            if (candidateScore > bestScore)
            {
                bestScore = candidateScore;
                best = new PageDiffResult
                {
                    PairIndex = pairIndex,
                    PageAfs1 = i1 + 1,
                    PageAfs2 = pageAfs2Index + 1,
                    AlignSim = Math.Round(candidateScore, 3),
                    Same = same,
                    Changed = changed,
                    Added = added,
                    Removed = 0,
                    TotalLines = diff.Count,
                    PctSame = Math.Round((double)same / Math.Max(1, diff.Count) * 100, 1),
                    Diffs = WithCrossPageNote(diff, i1 + 1, pageAfs2Index + 1),
                    SecondaryMatch = true,
                };
            }
        }

        return best;
    }

    private List<LineDiffResult> BuildUnmatchedDiffs(
        string text,
        string status,
        PdfReport? otherReport = null,
        int currentPageIdx = -1)
    {
        var lines = SplitLines(text);
        var diffs = new List<LineDiffResult>(lines.Count);

        foreach (var line in lines)
        {
            string? crossPageNote = null;

            if (otherReport != null && line.Length >= 15)
            {
                string lineNorm = Normalise(line);
                for (int pi = 0; pi < otherReport.Pages.Count; pi++)
                {
                    var otherLines = SplitLines(otherReport.Pages[pi]);
                    foreach (var otherLine in otherLines)
                    {
                        double sim = PageAlignmentService.SequenceSimilarity(
                            lineNorm,
                            Normalise(otherLine));

                        if (sim >= 0.90)
                        {
                            string side = status == "removed" ? "AFS2" : "AFS1";
                            int display = pi + 1;
                            int current = currentPageIdx + 1;
                            crossPageNote =
                                $"Same content found on {side} page {display} " +
                                $"(this page: {(status == "removed" ? "AFS1" : "AFS2")} page {current}) " +
                                "— page layout differs between drafts";
                            break;
                        }
                    }

                    if (crossPageNote != null)
                        break;
                }
            }

            diffs.Add(new LineDiffResult
            {
                Status = crossPageNote != null ? "same-different-page" : status,
                Line1 = status == "removed" ? line : string.Empty,
                Line2 = status == "added" ? line : string.Empty,
                Similarity = 0.0,
                CrossPageNote = crossPageNote,
            });
        }

        return diffs;
    }

    private static int SharedSectionHeadingCount(List<string> leftLines, List<string> rightLines)
    {
        var left = leftLines
            .Select(Normalise)
            .Where(line => _sectionHeadingLineRe.IsMatch(line))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (left.Count == 0)
            return 0;

        var right = rightLines
            .Select(Normalise)
            .Where(line => _sectionHeadingLineRe.IsMatch(line))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (right.Count == 0)
            return 0;

        return left.Intersect(right, StringComparer.OrdinalIgnoreCase).Count();
    }

    private static List<LineDiffResult> WithCrossPageNote(
        List<LineDiffResult> diffs,
        int pageAfs1,
        int pageAfs2)
    {
        string note = $"Continuation match: AFS1 p{pageAfs1} is related to AFS2 p{pageAfs2}, so some lines may continue across pages.";
        foreach (var diff in diffs)
            diff.CrossPageNote = note;
        return diffs;
    }

    private static bool IsIgnorableStandaloneLine(string line)
    {
        string norm = Normalise(line);
        if (string.IsNullOrWhiteSpace(norm))
            return true;

        if (_noteMarkerRe.IsMatch(norm))
            return true;

        var words = norm.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 0)
            return true;

        var nums = ExtractNumbers(line);
        if (nums.Count > 0)
            return false;

        if (norm is "-" or "–" or "—")
            return false;

        if (words.Length == 1 && char.IsLower(words[0][0]))
            return true;

        if (words.Length <= 3)
        {
            bool startsLower = char.IsLower(words[0][0]);
            bool possessiveLead = words[0].EndsWith("'s", StringComparison.Ordinal) || words[0].EndsWith("s'", StringComparison.Ordinal);
            bool shortAlphaFragment = words.All(w =>
                w.All(ch => char.IsLetter(ch) || ch == '\'' || ch == '-') &&
                w.Length <= 4);
            bool noSentenceStop = !norm.EndsWith('.') && !norm.EndsWith('!') && !norm.EndsWith('?');

            if ((startsLower || possessiveLead) && shortAlphaFragment && noSentenceStop)
                return true;
        }

        return false;
    }

    private static string ParentheticalCanonical(string s)
    {
        string t = Normalise(s);
        t = Regex.Replace(t, @"\([^)]*\)", " ");
        t = Regex.Replace(t, @"\bcontinued\b", " ");
        t = Regex.Replace(t, @"^(?:note\s+)?\d+(?:\.\d+)*\.?\s*", "");
        t = Regex.Replace(t, @"[^\w\s]", " ");
        t = Regex.Replace(t, @"\s+", " ").Trim();
        return t;
    }

    private static string SemanticSkeleton(string s)
    {
        string t = ParentheticalCanonical(s);
        t = Regex.Replace(t, @"\b(?:a|an|are|as|at|be|been|for|from|in|is|of|on|the|to|was|were)\b", " ");
        t = Regex.Replace(t, @"\s+", " ").Trim();
        return t;
    }

    // Build inverted index: bigram/first-token → set of line indices
    // Enables fast candidate retrieval without an O(n²) full scan.
    // Reference: Python LineComparator._build_index()
    private static Dictionary<string, HashSet<int>> BuildIndex(List<string> lines)
    {
        var idx = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < lines.Count; i++)
        {
            var toks = Normalise(lines[i]).Split(' ', StringSplitOptions.RemoveEmptyEntries);
            // First token
            if (toks.Length > 0)
            {
                if (!idx.ContainsKey(toks[0])) idx[toks[0]] = new();
                idx[toks[0]].Add(i);
            }
            // Bigrams
            for (int j = 0; j < toks.Length - 1; j++)
            {
                string bigram = toks[j] + toks[j + 1];
                if (!idx.ContainsKey(bigram)) idx[bigram] = new();
                idx[bigram].Add(i);
            }
        }
        return idx;
    }

    // Retrieve up to 40 candidate AFS 1 lines for a given AFS 2 needle.
    // Falls back to full index when no bigram / first-token hit is found.
    // Reference: Python LineComparator._candidates()
    private static List<int> Candidates(
        string needle,
        Dictionary<string, HashSet<int>> index,
        HashSet<int> used,
        int totalLines,
        int topK = 40)
    {
        var toks = Normalise(needle).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var hits = new HashSet<int>();

        if (toks.Length > 0 && index.TryGetValue(toks[0], out var h0))
            hits.UnionWith(h0);

        for (int j = 0; j < toks.Length - 1; j++)
        {
            string bigram = toks[j] + toks[j + 1];
            if (index.TryGetValue(bigram, out var hb))
                hits.UnionWith(hb);
        }

        hits.ExceptWith(used);
        if (hits.Count == 0)
            hits = Enumerable.Range(0, totalLines).Except(used).ToHashSet();

        // Sort before Take to get a deterministic, reproducible candidate set.
        return hits.OrderBy(x => x).Take(topK).ToList();
    }

    // Score two lines: base similarity + numeric bonus/penalty.
    // Reference: Python LineComparator._score()
    // Early-exit: if the first (cheapest) similarity already clears SimExact we skip the
    // three heavier canonical-form comparisons; if it is above SimRelated we still check
    // the others because canonical forms can expose additional matches.
    private static double Score(string l1, string l2)
    {
        double sim = PageAlignmentService.SequenceSimilarity(Normalise(l1), Normalise(l2));
        if (sim < SimExact)
        {
            sim = Math.Max(sim, PageAlignmentService.SequenceSimilarity(ParentheticalCanonical(l1), ParentheticalCanonical(l2)));
            if (sim < SimExact)
            {
                sim = Math.Max(sim, PageAlignmentService.SequenceSimilarity(SemanticSkeleton(l1), SemanticSkeleton(l2)));
                if (sim < SimExact)
                    sim = Math.Max(sim, PageAlignmentService.SequenceSimilarity(WhitespaceAgnosticCanonical(l1), WhitespaceAgnosticCanonical(l2)));
            }
        }
        var    nums1  = ExtractNumbers(l1);
        var    nums2  = ExtractNumbers(l2);
        double bonus  = nums1.Count > 0 && nums2.Count > 0 && nums1.SetEquals(nums2) ? 0.04 : 0.0;

        // When numbers differ, CAP similarity below SimExact so the line is always
        // flagged "changed" regardless of how similar the surrounding text is.
        // This is the key fix for lines like "Revenue 5 234 000" vs "Revenue 5 345 000"
        // which score ~0.95 text similarity but have genuinely different amounts.
        double penalty = 0.0;
        if (nums1.Count > 0 && nums2.Count > 0 && !nums1.SetEquals(nums2))
        {
            // Cap at 0.92 (below SimExact 0.96) to ensure "changed" status
            sim     = Math.Min(sim, 0.92);
            penalty = -0.02;   // small additional push down
        }

        return Math.Min(1.0, sim + bonus + penalty);
    }

    /// <summary>
    /// Determines whether a matched line has the same normalized content.
    ///
    /// Whitespace and punctuation normalization prevents formatting-only
    /// highlights. A wrap extension is ignored only when its ordered words
    /// form a prefix; loose word overlap can conceal substantive edits.
    /// </summary>
    private static string DetermineStatus(string l1, string l2)
    {
        if (AuditTextNormalizer.HasDifferentMeaningfulMarkers(l1, l2) ||
            !AuditTextNormalizer.ExtractNormalizedNumbers(l1)
                .SetEquals(AuditTextNormalizer.ExtractNormalizedNumbers(l2)))
        {
            return "changed";
        }

        // Gate 0: Case- and whitespace-insensitive equality.
        // Lines that differ ONLY in capitalisation, spacing, or Unicode whitespace variants
        // are NEVER flagged as changes — font size and spacing must not create diffs.
        if (Normalise(l1) == Normalise(l2))        return "same";

        // Gate 0.5: Explicit spacing/punctuation-only gate.
        // If stripping ALL punctuation AND ALL whitespace from both lines produces
        // identical strings, the difference is purely formatting (spacing, punctuation,
        // font-related PDF extraction artefacts) — not content.
        string pure1 = Regex.Replace(Normalise(l1), @"[^\w\d]", "");
        string pure2 = Regex.Replace(Normalise(l2), @"[^\w\d]", "");
        if (pure1.Length > 0 && pure1 == pure2)   return "same";

        // Gate 1: Canonical string equality (also strips punctuation and number formatting)
        if (Canonical(l1) == Canonical(l2))        return "same";
        if (WhitespaceAgnosticCanonical(l1) == WhitespaceAgnosticCanonical(l2))
            return "same";

        var n1 = ExtractNumbers(l1);
        var n2 = ExtractNumbers(l2);

        // Keep the legacy short-label reorder allowance without allowing a
        // single changed word in a longer disclosure to disappear in a set match.
        var w1set = new HashSet<string>(
            Normalise(l1).Split(' ', StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);
        var w2set = new HashSet<string>(
            Normalise(l2).Split(' ', StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);

        if (w1set.Count is > 0 and <= 4 &&
            w2set.Count is > 0 and <= 4 &&
            w1set.SetEquals(w2set) &&
            n1.SetEquals(n2))
            return "same";

        // A wrap can extend a line, but shared words alone cannot prove that
        // longer disclosures are equivalent. Require the shorter line to be an
        // exact ordered prefix of the longer line.
        bool numbersMatch = n1.Count == 0 && n2.Count == 0 || n1.SetEquals(n2);
        var orderedW1 = Canonical(l1).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var orderedW2 = Canonical(l2).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var shorterWords = orderedW1.Length <= orderedW2.Length ? orderedW1 : orderedW2;
        var longerWords = orderedW1.Length <= orderedW2.Length ? orderedW2 : orderedW1;
        string shorterText = orderedW1.Length <= orderedW2.Length ? Normalise(l1) : Normalise(l2);
        bool isOrderedPrefix = shorterWords.Length >= 4 &&
                               shorterWords.Length <= longerWords.Length &&
                               longerWords.Take(shorterWords.Length).SequenceEqual(
                                   shorterWords,
                                   StringComparer.OrdinalIgnoreCase);
        bool endsWithSentenceStop = shorterText.Length > 0 &&
                                    shorterText[^1] is '.' or '!' or '?';
        double lengthRatio = longerWords.Length == 0
            ? 0
            : (double)shorterWords.Length / longerWords.Length;
        if (numbersMatch && isOrderedPrefix && !endsWithSentenceStop && lengthRatio >= LengthRatioThreshold)
        {
            return "same";
        }

        // A longer extracted line may include a continuation from the next
        // wrapped line. Require a word boundary so a partial-word edit is kept.
        string gn1 = Normalise(l1), gn2 = Normalise(l2);
        if (gn1.Length >= 40 && gn2.Length >= 40)
        {
            string shorter = gn1.Length <= gn2.Length ? gn1 : gn2;
            string longer  = gn1.Length <= gn2.Length ? gn2 : gn1;
            char   shortEnd = shorter[^1];
            bool   endsStop = shortEnd is '.' or '!' or '?';
            bool   endsAtWordBoundary = shorter.Length < longer.Length &&
                (char.IsWhiteSpace(longer[shorter.Length]) || char.IsPunctuation(longer[shorter.Length]));
            if (!endsStop && endsAtWordBoundary &&
                longer.StartsWith(shorter, StringComparison.Ordinal))
                return "same";
        }

        return "changed";
    }

    // Normalise: lowercase, collapse whitespace, strip pure page-number lines.
    // internal static — shared with snapshot matching.
    // v4.3.2 FIX: NFKD decomposition + \p{C} strip eliminates invisible Unicode
    // control chars (zero-width spaces, object-replacement chars, etc.) that
    // PdfPig embeds around hyperlinked words, causing false "changed" verdicts.
    // Reference: Python _normalise()
    internal static string Normalise(string s)
    {
        // 1. NFKD decomposition: normalises typographic variants (ligatures,
        //    fancy quotes, etc.) to their base forms.
        string t = (s ?? "").Normalize(NormalizationForm.FormKD).Trim().ToLowerInvariant();
        // 2. Strip invisible Unicode control chars including zero-width spaces,
        //    soft-hyphens, object-replacement chars that PdfPig injects around
        //    hyperlinked words.  \uFFFD (REPLACEMENT CHARACTER, category So) is
        //    also stripped here — it is NOT in \p{C} but PdfPig emits it when it
        //    cannot decode a glyph from a non-standard font encoding.
        //    \uFFF0-\uFFFD covers the full Unicode Specials block.
        t = Regex.Replace(t, @"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]", "");
        // Collapse ALL Unicode whitespace variants (non-breaking space  , thin/en/em spaces,
        // narrow no-break  , math space  , ideographic 　). PDF column layout
        // uses these between table cells and must not produce false changed verdicts.
        t = Regex.Replace(t, @"[\s  -   　]+", " ").Trim();
        // Normalise ALL dash/minus variants → hyphen-minus so "provisions – 2025" == "provisions - 2025".
        // PDFs from Word use en-dash (–) or em-dash (—) in headings; other versions use plain hyphen.
        t = Regex.Replace(t, "[‐-―−﹘﹣－]", "-");
        // Normalise spacing around hyphens: "provisions  -  2025" → "provisions - 2025"
        t = Regex.Replace(t, @"\s*-\s*", " - ");
        t = Regex.Replace(t, @"^[-–—\s]*\d{1,4}[-–—\s]*$", "").Trim();
        // #10 — Number-format normalization so that "R 1 234 567", "R1,234,567",
        // and "1 234 567" compare as equal. Applied AFTER dash-normalization.
        t = _rCurrencyRe.Replace(t, "");                    // strip "R" currency prefix
        t = _saThousandsRe.Replace(t, "");                  // "1 234" → "1234"
        t = _commaDecimalRe.Replace(t, "$1.$2");            // "12,5" → "12.5" (SA decimal comma)
        t = Regex.Replace(t, @"(?<=\d),(?=\d{3})", "");    // "1,234" → "1234" (thousands comma)
        t = Regex.Replace(t, @"\s+", " ").Trim();
        t = AuditTextNormalizer.NormalizeDecimalFormatting(t);
        return AuditTextNormalizer.NormalizeNamedDates(
            AuditTextNormalizer.NormalizeCurrencyFormatting(t));
    }

    // Canonical: aggressive normalisation for the equality gates.
    // Strips whitespace collapse artefacts, thousands separators, punctuation.
    // Reference: Python LineComparator._canonical()
    private static string Canonical(string s)
    {
        string t = Normalise(s);
        t = Regex.Replace(t, @"[\s\u00a0]+", " ");
        t = Regex.Replace(t, @"(?<=\d)\s+(?=\d)", "");   // "102 575" → "102575"
        t = Regex.Replace(t, @"(?<=\d),(?=\d{3})", "");  // "102,575" → "102575"
        t = Regex.Replace(t, @"[^\w\s]", "");             // strip punctuation
        t = Regex.Replace(t, @"\s+", " ");                    // re-collapse gaps left by dash/punctuation removal
        return t.Trim();
    }

    private static string WhitespaceAgnosticCanonical(string s)
    {
        string t = Canonical(s);
        t = Regex.Replace(t, @"\s+", "");
        return t.Trim();
    }

    // Extract numeric tokens as a HashSet of canonical strings.
    // Page-stamp lines (e.g. "- 40 -") are stripped first to prevent footer
    // numbers creating spurious differences.
    // Reference: Python _extract_numbers()
    public static HashSet<string> ExtractNumbers(string text)
    {
        string cleaned = _pageStampRe.Replace(text ?? "", "");
        var result     = new HashSet<string>();

        foreach (Match m in _extNumRe.Matches(cleaned))
        {
            string raw = m.Groups[1].Success ? m.Groups[1].Value
                : m.Groups[2].Success ? m.Groups[2].Value
                : m.Groups[3].Value;

            string canon = Regex.Replace(raw, @"[\s,]", "");
            if (Regex.IsMatch(canon, @"^\d+(?:\.\d+)?$") && canon.Length >= 2)
                result.Add(canon);
        }
        return result;
    }

    // Symmetric difference as a sorted list (for NumDiff)
    // Reference: Python sorted(n1.symmetric_difference(n2))
    private static List<string> SymmetricExcept(HashSet<string> a, HashSet<string> b)
        => a.Except(b).Concat(b.Except(a)).OrderBy(x => x).ToList();

    private static string NormaliseWordToken(string word)
    {
        string t = (word ?? "")
            .Normalize(NormalizationForm.FormKD)
            .ToLowerInvariant()
            .Trim();
        t = Regex.Replace(t, @"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]", "");
        t = Regex.Replace(t, @"r\s+(?=\d)",            "r");
        t = Regex.Replace(t, @"(?<=\d)\s*\.\s*(?=\d)", ".");
        t = Regex.Replace(t, @"(?<=\d)\s*,\s*(?=\d)",  ",");
        return t;
    }

    // Word-level diff using Ratcliff/Obershelp opcodes.
    // Produces (word, tag) pairs: tag ∈ {"same", "added", "removed"}.
    // Reference: Python LineComparator._word_diff() — s1.split() splits on ALL whitespace.
    private static List<WordDiffToken> WordDiff(string s1, string s2)
    {
        // Null-coalesce and split on ALL whitespace (mirrors Python str.split() behaviour).
        var w1  = (s1 ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var w2  = (s2 ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);

        if (w1.Length == w2.Length &&
            w1.Zip(w2).All(p => NormaliseWordToken(p.First) == NormaliseWordToken(p.Second)))
        {
            return w1.Select(w => new WordDiffToken { Word = w, Tag = "same" }).ToList();
        }

        var ops = GetOpcodes(w1, w2);
        var out_ = new List<WordDiffToken>();

        foreach (var (tag, i1, i2, j1, j2) in ops)
        {
            switch (tag)
            {
                case "equal":
                    foreach (var w in w1[i1..i2]) out_.Add(new WordDiffToken { Word = w, Tag = "same" });
                    break;
                case "replace":
                    foreach (var w in w1[i1..i2]) out_.Add(new WordDiffToken { Word = w, Tag = "removed" });
                    foreach (var w in w2[j1..j2]) out_.Add(new WordDiffToken { Word = w, Tag = "added" });
                    break;
                case "delete":
                    foreach (var w in w1[i1..i2]) out_.Add(new WordDiffToken { Word = w, Tag = "removed" });
                    break;
                case "insert":
                    foreach (var w in w2[j1..j2]) out_.Add(new WordDiffToken { Word = w, Tag = "added" });
                    break;
            }
        }
        return out_;
    }

    // ── LCS-based opcode generator for word arrays ────────────────────────
    // Returns (tag, i1, i2, j1, j2) tuples — same semantics as
    // Python difflib.SequenceMatcher.get_opcodes().
    // tag values: "equal" | "replace" | "delete" | "insert"
    //
    // Algorithm (mirrors Python SequenceMatcher behaviour):
    //   1. Build LCS DP table (standard bottom-right → top-left).
    //   2. Single-step traceback: one element per iteration, producing atomic
    //      "equal", "insert", or "delete" ops.
    //      Tie-break: prefer delete over insert (dp[ai,bi+1] strictly > dp[ai+1,bi])
    //      so delete+insert pairs naturally merge into replace — matching Python
    //      SequenceMatcher's output order (removed words first, added words second).
    //   3. Merge consecutive same-type atomic ops into spans.
    //   4. Merge adjacent delete+insert → replace.
    //
    // Word equality: case-insensitive, whitespace-trimmed — per requirement that
    // capitalisation and spacing differences are never flagged as word changes.
    private static List<(string Tag, int I1, int I2, int J1, int J2)> GetOpcodes(
        string[] a, string[] b)
    {
        // Normalise a single word token for comparison.
        // Mirrors the user's NormalizeText() principle: normalise for comparison,
        // display the original.  Applied inside GetOpcodes only — the original tokens
        // are still returned in WordDiffToken so the UI shows unmodified text.
        //   "TOTAL"   == "Total"    (case)
        //   "R 700"   == "R700"     (currency prefix spacing)
        //   "70 .08"  == "70.08"    (decimal spacing)
        //   "1 , 000" == "1,000"    (comma spacing in numbers)
        static string Nw(string w)
        {
            string t = w.Trim().ToLowerInvariant();
            t = Regex.Replace(t, @"r\s+(?=\d)",            "r");  // R 70 → r70
            t = Regex.Replace(t, @"(?<=\d)\s*\.\s*(?=\d)", ".");  // 70 . 08 → 70.08
            t = Regex.Replace(t, @"(?<=\d)\s*,\s*(?=\d)",  ",");  // 1 , 000 → 1,000
            return t;
        }
        static bool Eq(string x, string y) => NormaliseWordToken(x) == NormaliseWordToken(y);

        int m = a.Length, n = b.Length;

        // ── 1. Build LCS DP table ─────────────────────────────────────────
        var dp = new int[m + 1, n + 1];
        for (int i = m - 1; i >= 0; i--)
            for (int j = n - 1; j >= 0; j--)
                dp[i, j] = Eq(a[i], b[j])
                    ? dp[i + 1, j + 1] + 1
                    : Math.Max(dp[i + 1, j], dp[i, j + 1]);

        // ── 2. Single-step traceback ──────────────────────────────────────
        var raw = new List<(string, int, int, int, int)>(m + n);
        int ai = 0, bi = 0;
        while (ai < m || bi < n)
        {
            if (ai < m && bi < n && Eq(a[ai], b[bi]))
            {
                raw.Add(("equal",  ai, ai + 1, bi, bi + 1));
                ai++; bi++;
            }
            else if (bi < n && (ai >= m || dp[ai, bi + 1] > dp[ai + 1, bi]))
            {
                // Strictly better to advance b → insert
                raw.Add(("insert", ai, ai, bi, bi + 1));
                bi++;
            }
            else
            {
                // Advance a (delete wins ties so delete precedes insert → merges to replace)
                raw.Add(("delete", ai, ai + 1, bi, bi));
                ai++;
            }
        }

        // ── 3. Merge consecutive same-type spans ──────────────────────────
        var merged = new List<(string, int, int, int, int)>(raw.Count);
        int k = 0;
        while (k < raw.Count)
        {
            var (tag, i1, i2, j1, j2) = raw[k++];
            while (k < raw.Count && raw[k].Item1 == tag)
            {
                i2 = raw[k].Item3;   // extend end of a-range
                j2 = raw[k].Item5;   // extend end of b-range
                k++;
            }
            merged.Add((tag, i1, i2, j1, j2));
        }

        // ── 4. Merge delete+insert → replace ─────────────────────────────
        var result = new List<(string, int, int, int, int)>(merged.Count);
        k = 0;
        while (k < merged.Count)
        {
            if (k + 1 < merged.Count
                && merged[k].Item1 == "delete"
                && merged[k + 1].Item1 == "insert")
            {
                result.Add(("replace",
                    merged[k].Item2, merged[k].Item3,
                    merged[k + 1].Item4, merged[k + 1].Item5));
                k += 2;
            }
            else
            {
                result.Add(merged[k++]);
            }
        }
        return result;
    }
}
