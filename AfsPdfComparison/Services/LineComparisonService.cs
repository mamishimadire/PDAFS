// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  SERVICE — LineComparisonService  v4.3.2                                     ║
// ║  Line-diff engine using F23.StringSimilarity.NormalizedLevenshtein for      ║
// ║  the matching score, paired with TextNormaliser for canonicalisation.        ║
// ║                                                                              ║
// ║  Content-aware status checks distinguish edits from extraction formatting.  ║
// ║  ROOT CAUSE: PDF text extraction wraps long paragraphs at different column  ║
// ║  widths per file. The same sentence broken differently scores 0.87-0.91 —   ║
// ║  above SIM_RELATED (matched) but below SIM_EXACT (flagged "changed").       ║
// ║  Wrap suppression requires an ordered prefix, not loose word overlap.       ║
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

        private const double LENGTH_RATIO_THRESHOLD = 0.55;

        private readonly NormalizedLevenshtein _lev = new();

        private List<string> SplitLines(string text)
        {
            var raw = (text ?? "").Split('\n')
                .Select(l => l.Trim())
                .Where(l => !TextNormaliser.IsNoise(l))
                .ToList();

            if (raw.Count == 0) return raw;

            // Merge only lowercase-start continuations; deduplicate consecutive
            // identical lines (hyperlink duplication artefact).
            var merged = new List<string>(raw.Count) { raw[0] };
            for (int i = 1; i < raw.Count; i++)
            {
                string prev   = merged[^1];
                string curr   = raw[i];
                char   last   = prev[^1];
                bool prevEnds = last is '.' or '!' or '?' or ';' or ':';
                bool currCont = curr.Length > 0 && char.IsLower(curr[0]);

                if (string.Equals(prev, curr, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Hyperlink sub-word dedup (adjacent pass, mirrors LineComparatorService):
                // if a short line (≤ 3 words) already appears verbatim as a phrase
                // inside the most recently merged line, it is a hyperlink annotation
                // duplicate extracted by PdfPig → discard it.
                if (curr.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length <= 3)
                {
                    string nm = TextNormaliser.Normalise(curr);
                    string np = TextNormaliser.Normalise(prev);
                    if (!string.IsNullOrEmpty(nm) && (" " + np + " ").Contains(" " + nm + " "))
                        continue;
                }

                if (!prevEnds && currCont && prev.Length >= 15)
                    merged[^1] = prev + " " + curr;
                else
                    merged.Add(curr);
            }

            // Second-pass hyperlink artefact dedup (order-independent).
            // Catches the case where the orphan hyperlink word appears at a
            // slightly different Y-band than the sentence, causing it to be
            // emitted BEFORE the containing sentence in extraction order.
            var deduped = new List<string>(merged.Count);
            for (int k = 0; k < merged.Count; k++)
            {
                string line = merged[k];
                int wc = line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
                if (wc > 0 && wc <= 3)
                {
                    string nm = TextNormaliser.Normalise(line);
                    bool orphan = false;
                    for (int j = 0; j < merged.Count; j++)
                    {
                        if (j == k) continue;
                        if ((" " + TextNormaliser.Normalise(merged[j]) + " ").Contains(" " + nm + " "))
                        {
                            orphan = true;
                            break;
                        }
                    }
                    if (orphan) continue;
                }
                deduped.Add(line);
            }
            return deduped;
        }

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
        /// Formatting normalization and a conservative ordered-prefix check
        /// prevent layout-only differences without hiding changed disclosure text.
        ///
        /// Normalization suppresses formatting artifacts. Similarity scores pair
        /// lines but do not prove their content is unchanged. A wrap is ignored
        /// only when the shorter line is an ordered prefix of the longer one.
        /// </summary>
        private static DiffStatus DetermineStatus(string l1, string l2)
        {
            // Gate 0: Case- and whitespace-insensitive equality.
            // Lines that differ ONLY in capitalisation or spacing are NEVER flagged as
            // changes — this is an absolute requirement. E.g. "TOTAL ASSETS" == "Total Assets".
            string normalized1 = TextNormaliser.Normalise(l1);
            string normalized2 = TextNormaliser.Normalise(l2);
            if (normalized1 == normalized2)
                return DiffStatus.Same;

            // Gate 1: Canonical string equality (also strips punctuation)
            if (TextNormaliser.Canonicalise(normalized1) == TextNormaliser.Canonicalise(normalized2))
                return DiffStatus.Same;

            // Preserve the short reordered-label behavior without accepting a
            // set match for longer text where word order carries meaning.
            var ws1 = new HashSet<string>(
                normalized1.Split(' ', StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase);
            var ws2 = new HashSet<string>(
                normalized2.Split(' ', StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase);
            var n1 = TextNormaliser.ExtractNumbers(l1);
            var n2 = TextNormaliser.ExtractNumbers(l2);

            if (ws1.Count is > 0 and <= 4 &&
                ws2.Count is > 0 and <= 4 &&
                ws1.SetEquals(ws2) &&
                n1.SetEquals(n2))
                return DiffStatus.Same;

            // Treat a line as a wrap only when the shorter line's ordered words
            // are an exact prefix of the longer line. Word overlap alone can hide
            // a real one-word edit in a long disclosure.
            bool numOk = (n1.Count == 0 && n2.Count == 0) || n1.SetEquals(n2);
            var orderedW1 = TextNormaliser.Canonicalise(normalized1)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var orderedW2 = TextNormaliser.Canonicalise(normalized2)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var shorterWords = orderedW1.Length <= orderedW2.Length ? orderedW1 : orderedW2;
            var longerWords = orderedW1.Length <= orderedW2.Length ? orderedW2 : orderedW1;
            string shorterText = orderedW1.Length <= orderedW2.Length
                ? normalized1
                : normalized2;
            bool isOrderedPrefix = shorterWords.Length >= 4 &&
                                   longerWords.Take(shorterWords.Length).SequenceEqual(
                                       shorterWords,
                                       StringComparer.OrdinalIgnoreCase);
            bool endsWithSentenceStop = shorterText.Length > 0 &&
                                        shorterText[^1] is '.' or '!' or '?';
            double lengthRatio = longerWords.Length == 0
                ? 0
                : (double)shorterWords.Length / longerWords.Length;
            if (numOk && isOrderedPrefix && !endsWithSentenceStop &&
                lengthRatio >= LENGTH_RATIO_THRESHOLD)
                return DiffStatus.Same;

            // Keep the longer-line wrap exception on a word boundary only.
            string gn1 = normalized1, gn2 = normalized2;
            if (gn1.Length >= 40 && gn2.Length >= 40)
            {
                string shorter  = gn1.Length <= gn2.Length ? gn1 : gn2;
                string longer   = gn1.Length <= gn2.Length ? gn2 : gn1;
                char   shortEnd = shorter[^1];
                bool   endsStop = shortEnd is '.' or '!' or '?';
                bool   endsAtWordBoundary = shorter.Length < longer.Length &&
                    (char.IsWhiteSpace(longer[shorter.Length]) || char.IsPunctuation(longer[shorter.Length]));
                if (!endsStop && endsAtWordBoundary &&
                    longer.StartsWith(shorter, StringComparison.Ordinal))
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
                    var status = DetermineStatus(l1, l2);
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
