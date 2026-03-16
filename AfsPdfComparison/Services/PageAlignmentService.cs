// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3                                                               ║
// ║                                                                              ║
// ║  SERVICE — PageAlignmentService                                              ║
// ║  Matches pages from two PDFs by content similarity, not page number.        ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • Jaccard similarity: |A ∩ B| / |A ∪ B| — word-set overlap              ║
// ║   • Ratcliff/Obershelp similarity: used for anchor-text structural match    ║
// ║     (C# equivalent of Python difflib.SequenceMatcher.ratio())               ║
// ║   • Python equivalent: build_page_alignment() in the notebook               ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text.RegularExpressions;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 2 · SMART PAGE ALIGNMENT SERVICE
//
// Problem: Two versions of an AFS may have different page counts because pages
// were inserted, removed, or split.  Naïve page-number matching (page 1 ↔ page 1)
// produces false changes.
//
// Solution: Build a "content fingerprint" per page — the first 3 non-noise lines
// (anchor) plus the set of words longer than 3 characters.  Then score every
// (i, j) pair within a sliding window using a weighted average of:
//   • Structural similarity: Ratcliff/Obershelp on the anchor strings
//   • Vocabulary similarity: Jaccard on the word sets
//
// Dynamic programming / greedy assignment ensures each AFS 2 page is used at
// most once.  Pages in AFS 1 with no good match (score < 0.30) are reported as
// unmatched so the auditor can see them explicitly.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Matches pages in AFS 1 to pages in AFS 2 by content similarity.
/// Registered as a singleton in DI.
/// </summary>
public class PageAlignmentService
{
    // Noise filters — same as Python _is_noise()
    private static readonly Regex _noiseShort   = new(@"^.{0,3}$",           RegexOptions.Compiled);
    private static readonly Regex _noiseDashes  = new(@"^[\s\-_=|\.]+$",     RegexOptions.Compiled);
    private static readonly Regex _noisePageNum = new(
        @"^(?:[-–—\s]*\(?\s*\d{1,4}\s*\)?[-–—\s]*|page\s+\d{1,4}(?:\s+of\s+\d{1,4})?)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // Minimum similarity threshold to accept a match (mirrors Python best_s = 0.30)
    private const double _minSim  = 0.30;

    // Sliding window: only compare pages within ±8 positions of each other
    // to keep complexity O(n × window) rather than O(n²).
    private const int _window = 8;

    /// <summary>
    /// Builds a page alignment map between two PdfReports.
    /// </summary>
    /// <returns>
    /// List of tuples (i1, i2, sim) where i1 is a 0-based AFS 1 page index,
    /// i2 is the matched 0-based AFS 2 page index (−1 if unmatched), and
    /// sim is the similarity score 0–1.
    /// </returns>
    public List<AlignmentEntry> BuildAlignment(
        List<string> pages1, List<string> pages2)
    {
        // ── 2.1  Build fingerprints for every page ─────────────────────────
        var fps1 = pages1.Select(PageFingerprint).ToList();
        var fps2 = pages2.Select(PageFingerprint).ToList();
        int n1 = fps1.Count, n2 = fps2.Count;

        // ── 2.2  Compute similarity matrix (sparse, within window) ─────────
        //   sim[i][j] stored only when |i-j| ≤ window to save memory.
        var sim = new double[n1, n2];
        for (int i = 0; i < n1; i++)
        {
            int jStart = Math.Max(0, i - _window);
            int jEnd   = Math.Min(n2, i + _window + 1);
            for (int j = jStart; j < jEnd; j++)
                sim[i, j] = PageSim(fps1[i], fps2[j]);
        }

        // ── 2.3  Greedy best-match assignment ──────────────────────────────
        // For each AFS 1 page (in order) find the highest-scoring unmatched
        // AFS 2 page within the window.
        var used2   = new HashSet<int>();
        var result  = new List<AlignmentEntry>(n1);

        for (int i = 0; i < n1; i++)
        {
            int    bestJ = -1;
            double bestS = _minSim;
            int jStart   = Math.Max(0, i - _window);
            int jEnd     = Math.Min(n2, i + _window + 1);

            for (int j = jStart; j < jEnd; j++)
            {
                if (!used2.Contains(j) && sim[i, j] > bestS)
                {
                    bestS = sim[i, j];
                    bestJ = j;
                }
            }

            if (bestJ >= 0)
                used2.Add(bestJ);

            result.Add(new AlignmentEntry { I1 = i, I2 = bestJ, Sim = bestJ >= 0 ? Math.Round(bestS, 3) : 0.0 });
        }

        return result;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 2-INTERNAL · PAGE FINGERPRINT
    //
    // A fingerprint consists of two parts:
    //   anchor  — first 3 meaningful lines joined and lowercased, capped at 120 chars
    //   words   — HashSet of words longer than 3 characters from the page
    //
    // The anchor captures the structural header of the page (e.g. note heading).
    // The word set captures vocabulary distribution.
    // ─────────────────────────────────────────────────────────────────────────
    private static (string Anchor, HashSet<string> Words) PageFingerprint(string text)
    {
        var cleanLines = text.Split('\n')
            .Select(l => l.Trim())
            .Where(l => !IsNoise(l) && l.Length > 8)
            .Take(3)
            .ToList();

        string anchor = string.Join(" ", cleanLines).ToLowerInvariant();
        if (anchor.Length > 120) anchor = anchor[..120];

        var words = Regex.Matches(text.ToLowerInvariant(), @"[a-zA-Z]{4,}")
            .Select(m => m.Value)
            .ToHashSet();

        return (anchor, words);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 2-INTERNAL · PAGE SIMILARITY
    //
    // Weighted combination:
    //   0.50 × structural similarity (Ratcliff/Obershelp on anchor strings)
    //   0.50 × Jaccard similarity (word-set intersection / union)
    //
    // Reference: Python _page_sim() in the notebook.
    // ─────────────────────────────────────────────────────────────────────────
    private static double PageSim(
        (string Anchor, HashSet<string> Words) fp1,
        (string Anchor, HashSet<string> Words) fp2)
    {
        double structural = SequenceSimilarity(fp1.Anchor, fp2.Anchor);
        double jaccard    = fp1.Words.Count + fp2.Words.Count == 0
                            ? 0.0
                            : (double)fp1.Words.Intersect(fp2.Words).Count()
                              / fp1.Words.Union(fp2.Words).Count();

        return 0.5 * structural + 0.5 * jaccard;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 2-INTERNAL · RATCLIFF/OBERSHELP STRING SIMILARITY
    //
    // C# implementation of the algorithm behind Python's difflib.SequenceMatcher.
    // Finds the longest common substring recursively.
    // Time complexity: O(n²) for short strings (anchors ≤ 120 chars).
    // Reference: Ratcliff, J.W. & Metzener, D.E. (1988) "Pattern Matching: The
    //            Gestalt Approach", Dr Dobb's Journal, July 1988.
    // ─────────────────────────────────────────────────────────────────────────
    public static double SequenceSimilarity(string a, string b)
    {
        if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b)) return 1.0;
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0.0;
        int matches = CountMatches(a, 0, a.Length, b, 0, b.Length);
        return 2.0 * matches / (a.Length + b.Length);
    }

    private static int CountMatches(string a, int a0, int a1, string b, int b0, int b1)
    {
        int best = 0, bestI = a0, bestJ = b0;
        for (int i = a0; i < a1; i++)
        {
            for (int j = b0; j < b1; j++)
            {
                int k = 0;
                while (i + k < a1 && j + k < b1 && a[i + k] == b[j + k])
                    k++;
                if (k > best) { best = k; bestI = i; bestJ = j; }
            }
        }
        if (best == 0) return 0;
        int total = best;
        if (bestI > a0 && bestJ > b0)
            total += CountMatches(a, a0, bestI, b, b0, bestJ);
        if (bestI + best < a1 && bestJ + best < b1)
            total += CountMatches(a, bestI + best, a1, b, bestJ + best, b1);
        return total;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 2-INTERNAL · NOISE DETECTION
    //
    // Lines considered "noise" are suppressed before fingerprinting and before
    // line comparison to avoid false positives.
    // Noise categories (mirrors Python _is_noise()):
    //   • Lines shorter than 4 characters
    //   • Lines consisting only of whitespace, dashes, underscores, pipes, dots
    //   • Lines that are purely a page-number stamp such as "- 40 -" or "Page 3 of 10"
    // ─────────────────────────────────────────────────────────────────────────
    public static bool IsNoise(string line)
    {
        if (string.IsNullOrWhiteSpace(line) || line.Length < 4) return true;
        if (_noiseDashes.IsMatch(line))  return true;
        if (_noisePageNum.IsMatch(line)) return true;
        return false;
    }
}
