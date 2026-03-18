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
// ║  v4.3.2 FIX — Gate 5: Word-overlap line-wrap artefact gate                  ║
// ║  ROOT CAUSE: pdfplumber/PdfPig wraps long paragraphs at different column    ║
// ║  widths per PDF. The sentence:                                               ║
// ║    "used. The separate financial statements...IAS 27 Consolidated"           ║
// ║  in AFS1 continues in AFS2 as:                                               ║
// ║    "used. The separate financial statements...IAS 27 Consolidated            ║
// ║     and Separate Financial Statements."                                      ║
// ║  These score 0.87–0.91 — above SIM_RELATED(0.55) so they are matched,      ║
// ║  but below SIM_EXACT(0.94) so they were falsely flagged "changed".          ║
// ║  Gate 5 fix: if >=85% of shorter line's words appear in the longer line     ║
// ║  AND shorter is >=55% the length of longer → line-wrap artefact → "same"    ║
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
// Five-gate equality (v4.3.2):
//   Gate 1 — Canonical string equality (ignores whitespace, punctuation, number fmt)
//   Gate 2 — Similarity score ≥ SIM_EXACT (0.94)
//   Gate 3 — All numbers match AND score ≥ 0.80
//   Gate 4 — Normalised text equality
//   Gate 5 — Word-overlap ≥ 0.85 AND length ratio ≥ 0.55 (line-wrap artefact)
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
    private const double SimExact   = 0.94;
    private const double SimRelated = 0.55;

    // Gate 5 thresholds (word-overlap line-wrap artefact gate)
    private const double WordOverlapThreshold = 0.85;
    private const double LengthRatioThreshold = 0.55;

    // Regex: strip page-number stamps before number extraction.
    // Matches lines that are PURELY a page stamp (e.g. "- 40 -", "Page 3 of 10").
    // Reference: Python _PAGE_NUM_RE
    private static readonly Regex _pageStampRe = new(
        @"(?m)^[-–—\s]*\(?\s*\d{1,4}\s*\)?[-–—\s]*$",
        RegexOptions.Compiled | RegexOptions.Multiline);

    // Regex: extract numeric tokens (integers and decimals).
    // Handles South African space-thousands-separator format: "6 835" → 6835.
    // Only allows a SINGLE space between digit groups of exactly 3 digits
    // (i.e. real thousands separator, not arbitrary whitespace between numbers).
    private static readonly Regex _numTokenRe = new(
        @"\b\d{1,3}(?:[,\.]?\d{3})*(?:\.\d+)?\b|\b\d+ \d{3}(?:\.\d+)?\b", RegexOptions.Compiled);

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
        var index1 = BuildIndex(lines1);   // inverted index over AFS 1

        var used1   = new HashSet<int>();
        var used2   = new HashSet<int>();
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
                used2.Add(i2);
                string l1     = lines1[bestIdx];
                string status = DetermineStatus(l1, l2, bestSc);
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
                results.Add(new LineDiffResult
                {
                    Status = "removed", Line1 = lines1[i1], Line2 = "",
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
        var results     = new List<PageDiffResult>();
        var alignedI1   = new HashSet<int>();
        int pairIndex   = 0;

        foreach (var entry in alignment.Where(a => a.I2 >= 0))
        {
            int i1 = entry.I1, i2 = entry.I2;
            alignedI1.Add(i1);
            var diff  = CompareLines(report1.Pages[i1], report2.Pages[i2]);
            int same  = diff.Count(d => d.Status == "same");
            int chg   = diff.Count(d => d.Status == "changed");
            int add   = diff.Count(d => d.Status == "added");
            int rem   = diff.Count(d => d.Status == "removed");
            int total = Math.Max(diff.Count, 1);

            results.Add(new PageDiffResult
            {
                PairIndex  = pairIndex++,
                PageAfs1   = i1 + 1,
                PageAfs2   = i2 + 1,
                AlignSim   = entry.Sim,
                Same       = same,
                Changed    = chg,
                Added      = add,
                Removed    = rem,
                TotalLines = diff.Count,
                PctSame    = Math.Round((double)same / total * 100, 1),
                Diffs      = diff,
            });
        }

        for (int i1 = 0; i1 < report1.Pages.Count; i1++)
        {
            if (alignedI1.Contains(i1)) continue;
            int removedCount = report1.Pages[i1]
                .Split('\n')
                .Count(l => !PageAlignmentService.IsNoise(l.Trim()));

            results.Add(new PageDiffResult
            {
                PairIndex  = pairIndex++,
                PageAfs1   = i1 + 1,
                PageAfs2   = null,
                AlignSim   = 0.0,
                Removed    = removedCount,
                TotalLines = 0,
                PctSame    = 0.0,
                Diffs      = new(),
                Unmatched  = true,
            });
        }

        return results;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 3-INTERNAL · HELPERS
    // ─────────────────────────────────────────────────────────────────────────

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

            if (!prevEnds && currCont && prev.Length >= 15)
                merged[^1] = prev + " " + curr;
            else
                merged.Add(curr);
        }

        return merged;
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
    private static double Score(string l1, string l2)
    {
        double sim    = PageAlignmentService.SequenceSimilarity(Normalise(l1), Normalise(l2));
        var    nums1  = ExtractNumbers(l1);
        var    nums2  = ExtractNumbers(l2);
        double bonus  = nums1.Count > 0 && nums2.Count > 0 && nums1.SetEquals(nums2) ?  0.04 : 0.0;
        double penalty= nums1.Count > 0 && nums2.Count > 0 && !nums1.SetEquals(nums2) ? -0.06 : 0.0;
        return Math.Min(1.0, sim + bonus + penalty);
    }

    /// <summary>
    /// Five-gate equality determination. (v4.3.2)
    ///
    /// Gate 1 — Canonical string equality
    /// Gate 2 — Similarity score ≥ SIM_EXACT (0.94)
    /// Gate 3 — All numbers identical AND score ≥ 0.80
    /// Gate 4 — Normalised text equality
    /// Gate 5 — Word-overlap line-wrap artefact gate (NEW in v4.3.2)
    ///
    /// Gate 5 rationale:
    /// PDF text extraction wraps long paragraphs at the column width of each
    /// individual PDF. Two PDFs with slightly different margins produce the
    /// same sentence broken at different points, e.g.:
    ///   AFS1: "...requirements of IAS 27 Consolidated"         (line ends here)
    ///   AFS2: "...requirements of IAS 27 Consolidated and Separate Financial..."
    /// These score ~0.87–0.91 — matched but below SIM_EXACT — causing a false
    /// "changed" flag and incorrect yellow highlight on the snapshot.
    /// Gate 5: if ≥85% of the shorter line's word tokens appear in the longer
    /// line AND the shorter is ≥55% the length of the longer, this is a
    /// line-wrap artefact, not a genuine content change → return "same".
    /// </summary>
    private static string DetermineStatus(string l1, string l2, double rawScore)
    {
        // Gate 1: Canonical string equality
        if (Canonical(l1) == Canonical(l2))        return "same";

        // Gate 2: Similarity score above exact threshold
        if (rawScore >= SimExact)                  return "same";

        // Gate 3: Numbers identical at lower score
        var n1 = ExtractNumbers(l1);
        var n2 = ExtractNumbers(l2);
        if (n1.SetEquals(n2) && rawScore >= 0.80)  return "same";

        // Gate 4: Normalised text equality
        if (Normalise(l1) == Normalise(l2))        return "same";

        // Gate 5: Word-overlap line-wrap artefact gate
        var w1 = new HashSet<string>(
            Normalise(l1).Split(' ', StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);
        var w2 = new HashSet<string>(
            Normalise(l2).Split(' ', StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);

        // Gate 5 is directional: only fires when l2 (AFS2) has at least as many
        // words as l1 (AFS1).  Rationale: line-wrap artefacts occur when AFS1
        // wraps a sentence EARLIER than AFS2 (AFS2's column is wider), so l2 is
        // LONGER.  If l2 is SHORTER, the missing words in AFS2 represent real
        // content removals (e.g. a number like "11 639" present in AFS1 but
        // absent in AFS2) and must NOT be silenced — they are genuine changes.
        //
        // NUMBER GUARD: if the lines contain any numeric tokens and those sets
        // differ even by one digit (e.g. R217 vs R218), this is a genuine content
        // change — Gate 5 must NOT suppress it regardless of word-overlap.
        //
        // SHORT-LINE GUARD: a 1–3 word line (e.g. single hyperlink word "users")
        // must not be silenced by matching against a long paragraph that happens
        // to contain that word.
        bool numbersMatch = n1.Count == 0 && n2.Count == 0 || n1.SetEquals(n2);
        if (w1.Count > 0 && w2.Count > 0 && w2.Count >= w1.Count
            && numbersMatch
            && w1.Count >= 4)   // short-line guard: skip for 1–3 word lines
        {
            int    intersection = w1.Intersect(w2).Count();
            int    minLen       = w1.Count;   // l1 is shorter or equal
            int    maxLen       = w2.Count;   // l2 is longer or equal
            double overlap  = (double)intersection / minLen;
            double lenRatio = (double)minLen / maxLen;

            if (overlap >= WordOverlapThreshold && lenRatio >= LengthRatioThreshold)
                return "same";
        }

        // Gate 6: Paragraph-wrap prefix gate.
        // When one PDF's column is wider, its extracted line contains the entire
        // text of the narrower PDF's line PLUS the next wrapped continuation.
        // If the shorter normalised line (≥ 40 chars) is a LEADING SUBSTRING of
        // the longer one AND the shorter line does NOT end in a sentence-terminal
        // character (proving it is a mid-sentence fragment, not a truncated
        // sentence), treat this pair as identical — same text, different wrap.
        string gn1 = Normalise(l1), gn2 = Normalise(l2);
        if (gn1.Length >= 40 && gn2.Length >= 40)
        {
            string shorter = gn1.Length <= gn2.Length ? gn1 : gn2;
            string longer  = gn1.Length <= gn2.Length ? gn2 : gn1;
            char   shortEnd = shorter[^1];
            bool   endsStop = shortEnd is '.' or '!' or '?';
            if (!endsStop && longer.StartsWith(shorter, StringComparison.Ordinal))
                return "same";
        }

        return "changed";
    }

    // Normalise: lowercase, collapse whitespace, strip pure page-number lines.
    // internal static — also used by Gate 4 / Gate 5 in DetermineStatus.
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
        t = Regex.Replace(t, @"\s+", " ").Trim();
        t = Regex.Replace(t, @"^[-–—\s]*\d{1,4}[-–—\s]*$", "").Trim();
        return t;
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
        foreach (Match m in _numTokenRe.Matches(cleaned))
        {
            string canon = Regex.Replace(m.Value, @"[\s,]", "");
            if (Regex.IsMatch(canon, @"^\d+(?:\.\d+)?$"))
                result.Add(canon);
        }
        return result;
    }

    // Symmetric difference as a sorted list (for NumDiff)
    // Reference: Python sorted(n1.symmetric_difference(n2))
    private static List<string> SymmetricExcept(HashSet<string> a, HashSet<string> b)
        => a.Except(b).Concat(b.Except(a)).OrderBy(x => x).ToList();

    // Word-level diff using Ratcliff/Obershelp opcodes.
    // Produces (word, tag) pairs: tag ∈ {"same", "added", "removed"}.
    // Reference: Python LineComparator._word_diff() — s1.split() splits on ALL whitespace.
    private static List<WordDiffToken> WordDiff(string s1, string s2)
    {
        // Null-coalesce and split on ALL whitespace (mirrors Python str.split() behaviour).
        var w1  = (s1 ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var w2  = (s2 ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
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
    private static List<(string Tag, int I1, int I2, int J1, int J2)> GetOpcodes(
        string[] a, string[] b)
    {
        int m = a.Length, n = b.Length;

        // ── 1. Build LCS DP table ─────────────────────────────────────────
        var dp = new int[m + 1, n + 1];
        for (int i = m - 1; i >= 0; i--)
            for (int j = n - 1; j >= 0; j--)
                dp[i, j] = a[i] == b[j]
                    ? dp[i + 1, j + 1] + 1
                    : Math.Max(dp[i + 1, j], dp[i, j + 1]);

        // ── 2. Single-step traceback ──────────────────────────────────────
        var raw = new List<(string, int, int, int, int)>(m + n);
        int ai = 0, bi = 0;
        while (ai < m || bi < n)
        {
            if (ai < m && bi < n && a[ai] == b[bi])
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
