// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  SERVICE — LineComparisonService  v4.3.2                                     ║
// ║  Line-diff engine using F23.StringSimilarity.NormalizedLevenshtein for      ║
// ║  the matching score, paired with TextNormaliser for canonicalisation.        ║
// ║                                                                              ║
// ║  v4.3.2 FIX — Gate 5: Word-overlap line-wrap artefact gate                  ║
// ║  ROOT CAUSE: PDF text extraction wraps long paragraphs at different column  ║
// ║  widths per file. The same sentence broken differently scores 0.87-0.91 —   ║
// ║  above SIM_RELATED (matched) but below SIM_EXACT (flagged "changed").       ║
// ║  Gate 5: word overlap ≥ 85% AND length ratio ≥ 55% → "same"                ║
// ║                                                                              ║
// ║  NOTE: This service is used by ComparisonController (MVC route).            ║
// ║  The main pipeline uses LineComparatorService (Ratcliff-Obershelp).         ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using F23.StringSimilarity;

namespace AfsPdfComparison.Services
{
    public enum DiffStatus { Same, Changed, Added, Removed }

    public class LineDiff
    {
        public DiffStatus   Status     { get; set; }
        public string       Line1      { get; set; } = "";
        public string       Line2      { get; set; } = "";
        public double       Similarity { get; set; }
        public List<string> NumericDiff { get; set; } = new();
    }

    public class ComparisonCounts
    {
        public int Same    { get; set; }
        public int Changed { get; set; }
        public int Added   { get; set; }
        public int Removed { get; set; }
        public int Total   => Same + Changed + Added + Removed;
        public double PctSame => Total == 0 ? 100.0
            : Math.Round(Same * 100.0 / Total, 1);
    }

    public class LineComparisonService
    {
        // ── Similarity thresholds ────────────────────────────────────────────────
        private const double SIM_EXACT   = 0.94;
        private const double SIM_RELATED = 0.55;

        // Gate 5 thresholds (word-overlap line-wrap artefact gate)
        private const double WORD_OVERLAP_THRESHOLD = 0.85;
        private const double LENGTH_RATIO_THRESHOLD = 0.55;

        private readonly NormalizedLevenshtein _lev = new();

        private List<string> SplitLines(string text) =>
            (text ?? "").Split('\n')
                .Select(l => l.Trim())
                .Where(l => !TextNormaliser.IsNoise(l))
                .ToList();

        // Build inverted index: first-token and bigrams → set of line indices
        private Dictionary<string, HashSet<int>> BuildIndex(List<string> lines)
        {
            var idx = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < lines.Count; i++)
            {
                var tokens = TextNormaliser.Canonicalise(lines[i])
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries);
                // First unigram
                if (tokens.Length > 0)
                {
                    if (!idx.ContainsKey(tokens[0])) idx[tokens[0]] = new();
                    idx[tokens[0]].Add(i);
                }
                // Bigrams
                for (int j = 0; j < tokens.Length - 1; j++)
                {
                    var key = tokens[j] + tokens[j + 1];
                    if (!idx.ContainsKey(key)) idx[key] = new();
                    idx[key].Add(i);
                }
            }
            return idx;
        }

        // Similarity score: NormalizedLevenshtein on canonicalised text,
        // plus numeric bonus/penalty
        private double Score(string a, string b)
        {
            var norm_a = TextNormaliser.Canonicalise(a);
            var norm_b = TextNormaliser.Canonicalise(b);
            // NormalizedLevenshtein.Distance() returns 0 (identical) to 1 (completely different)
            double sim    = 1.0 - _lev.Distance(norm_a, norm_b);
            var    nums_a = TextNormaliser.ExtractNumbers(a);
            var    nums_b = TextNormaliser.ExtractNumbers(b);
            if (nums_a.Count > 0 && nums_b.Count > 0)
            {
                if (nums_a.SetEquals(nums_b)) sim = Math.Min(1.0, sim + 0.04);
                else                          sim = Math.Max(0.0, sim - 0.06);
            }
            return sim;
        }

        /// <summary>
        /// Determines whether two matched lines are "same" or "changed".
        ///
        /// Gate 1 — Canonical string equality
        /// Gate 2 — Similarity score ≥ SIM_EXACT (0.94)
        /// Gate 3 — All numbers identical AND score ≥ 0.80
        /// Gate 4 — Normalised text equality (catches soft-hyphen / space variants)
        /// Gate 5 — Word-overlap line-wrap artefact gate (v4.3.2)
        ///
        /// Gate 5 eliminates false highlights caused by PDF column-width differences.
        /// Example: pdfplumber extracts the same paragraph as:
        ///   AFS1: "...requirements of IAS 27 Consolidated"
        ///   AFS2: "...requirements of IAS 27 Consolidated and Separate Financial..."
        /// These score 0.87-0.91 and were falsely flagged as "changed".
        /// Gate 5: if ≥85% of the shorter line's words appear in the longer
        /// AND shorter ≥55% the length of longer → line-wrap artefact → "same".
        /// </summary>
        private static DiffStatus DetermineStatus(string l1, string l2, double rawScore)
        {
            // Gate 1: Canonical string equality
            if (TextNormaliser.Canonicalise(l1) == TextNormaliser.Canonicalise(l2))
                return DiffStatus.Same;

            // Gate 2: Similarity score above exact threshold
            if (rawScore >= SIM_EXACT)
                return DiffStatus.Same;

            // Gate 3: Numbers identical at lower score
            var n1 = TextNormaliser.ExtractNumbers(l1);
            var n2 = TextNormaliser.ExtractNumbers(l2);
            if (n1.SetEquals(n2) && rawScore >= 0.80)
                return DiffStatus.Same;

            // Gate 4: Normalised text equality
            if (TextNormaliser.Normalise(l1) == TextNormaliser.Normalise(l2))
                return DiffStatus.Same;

            // Gate 5: Word-overlap line-wrap artefact gate
            var w1 = new HashSet<string>(
                TextNormaliser.Normalise(l1).Split(' ', StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase);
            var w2 = new HashSet<string>(
                TextNormaliser.Normalise(l2).Split(' ', StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase);

            if (w1.Count > 0 && w2.Count > 0)
            {
                int    intersection = w1.Intersect(w2).Count();
                int    maxLen       = Math.Max(w1.Count, w2.Count);
                int    minLen       = Math.Min(w1.Count, w2.Count);
                double overlap      = (double)intersection / maxLen;
                double lenRatio     = (double)minLen / maxLen;

                if (overlap >= WORD_OVERLAP_THRESHOLD && lenRatio >= LENGTH_RATIO_THRESHOLD)
                    return DiffStatus.Same;
            }

            return DiffStatus.Changed;
        }

        /// <summary>
        /// Full-document line comparison using inverted-index candidate selection.
        /// </summary>
        public List<LineDiff> CompareLines(string text1, string text2)
        {
            var lines1 = SplitLines(text1);
            var lines2 = SplitLines(text2);
            var idx1   = BuildIndex(lines1);
            var used1  = new HashSet<int>();
            var used2  = new HashSet<int>();
            var results = new List<LineDiff>();

            for (int i2 = 0; i2 < lines2.Count; i2++)
            {
                var l2     = lines2[i2];
                var tokens = TextNormaliser.Canonicalise(l2)
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries);

                // Gather candidates via inverted index
                var candidates = new HashSet<int>();
                if (tokens.Length > 0 && idx1.TryGetValue(tokens[0], out var first))
                    candidates.UnionWith(first);
                for (int j = 0; j < tokens.Length - 1; j++)
                {
                    var key = tokens[j] + tokens[j + 1];
                    if (idx1.TryGetValue(key, out var set))
                        candidates.UnionWith(set);
                }
                candidates.ExceptWith(used1);

                // Fall back to full scan when index yields nothing
                if (candidates.Count == 0)
                    candidates = Enumerable.Range(0, lines1.Count).Except(used1).ToHashSet();

                double bestScore = 0;
                int    bestIdx   = -1;
                foreach (var ci in candidates.Take(40))
                {
                    var s = Score(l2, lines1[ci]);
                    if (s > bestScore) { bestScore = s; bestIdx = ci; }
                }

                if (bestIdx >= 0 && bestScore >= SIM_RELATED)
                {
                    used1.Add(bestIdx); used2.Add(i2);
                    var l1     = lines1[bestIdx];
                    var status = DetermineStatus(l1, l2, bestScore);
                    var n1     = TextNormaliser.ExtractNumbers(l1);
                    var n2     = TextNormaliser.ExtractNumbers(l2);
                    var numDiff = status == DiffStatus.Changed
                        ? n1.Concat(n2).Except(n1.Intersect(n2)).OrderBy(x => x).ToList()
                        : new List<string>();
                    results.Add(new LineDiff
                    {
                        Status     = status,
                        Line1      = l1,
                        Line2      = l2,
                        Similarity = Math.Round(bestScore, 4),
                        NumericDiff = numDiff,
                    });
                }
                else
                {
                    results.Add(new LineDiff
                    {
                        Status = DiffStatus.Added, Line2 = l2, Similarity = 0
                    });
                }
            }

            // Remaining unmatched AFS 1 lines → Removed
            for (int i1 = 0; i1 < lines1.Count; i1++)
            {
                if (!used1.Contains(i1))
                    results.Add(new LineDiff
                    {
                        Status = DiffStatus.Removed, Line1 = lines1[i1], Similarity = 0
                    });
            }
            return results;
        }

        public ComparisonCounts CountResults(List<LineDiff> diffs) => new()
        {
            Same    = diffs.Count(d => d.Status == DiffStatus.Same),
            Changed = diffs.Count(d => d.Status == DiffStatus.Changed),
            Added   = diffs.Count(d => d.Status == DiffStatus.Added),
            Removed = diffs.Count(d => d.Status == DiffStatus.Removed),
        };
    }
}
