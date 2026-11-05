using System.Text.RegularExpressions;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

public class PageAlignmentService
{
    private const int StubWordThreshold = 20;

    private sealed class PageFingerprintData
    {
        public string Anchor { get; set; } = string.Empty;
        public List<string> Chunks { get; } = new();
        public List<string> Headings { get; } = new();
        public HashSet<string> Words { get; } = new(StringComparer.Ordinal);
        public HashSet<string> SectionTags { get; } = new(StringComparer.OrdinalIgnoreCase);
    }

    private static readonly Regex NoiseDashes = new(@"^[\s\-_=|\.]+$", RegexOptions.Compiled);
    private static readonly Regex NoisePageNum = new(
        @"^(?:[-–—\s]*\(?\s*\d{1,4}\s*\)?[-–—\s]*|page\s+\d{1,4}(?:\s+of\s+\d{1,4})?)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex WordToken = new(@"[a-zA-Z]{4,}", RegexOptions.Compiled);
    private static readonly Regex WhiteSpace = new(@"\s+", RegexOptions.Compiled);
    private static readonly Regex SectionHeading = new(
        @"^(?:note\s+)?\d+(?:\.\d+)*\.?\s+[a-z]",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex AdditionalPageRe = new(
        @"^additional\s+page\s+\d+$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private const double MinSim = 0.30;
    private const double MinRelocationSim = 0.58;
    private const double MinRelocationMargin = 0.07;

    public List<AlignmentEntry> BuildAlignment(List<string> pages1, List<string> pages2)
    {
        var (content1, idx1) = CollapseStubs(pages1);
        var (content2, idx2) = CollapseStubs(pages2);

        var rawAlignment = AlignContentPages(content1, content2);
        MatchConfidentRelocations(content1, content2, rawAlignment);

        var remapped = rawAlignment.Select(a => new AlignmentEntry
        {
            I1 = a.I1 >= 0 && a.I1 < idx1.Count ? idx1[a.I1] : -1,
            I2 = a.I2 >= 0 && a.I2 < idx2.Count ? idx2[a.I2] : -1,
            Sim = a.Sim,
            MatchType = a.MatchType,
        }).ToList();

        var byI1 = remapped.Where(a => a.I1 >= 0).ToDictionary(a => a.I1);
        var result = new List<AlignmentEntry>(pages1.Count);
        for (int i = 0; i < pages1.Count; i++)
        {
            result.Add(byI1.TryGetValue(i, out var e)
                ? e
                : new AlignmentEntry { I1 = i, I2 = -1, Sim = 0.0 });
        }

        return result;
    }

    /// <summary>
    /// The ordered page alignment handles insertions and pagination shifts well, but a
    /// relocated section can cross another section and therefore be left unmatched.
    /// Reconcile only the unmatched pages using the same semantic fingerprint score;
    /// ambiguous candidates stay unmatched for review.
    /// </summary>
    private void MatchConfidentRelocations(
        List<string> pages1,
        List<string> pages2,
        List<AlignmentEntry> alignment)
    {
        if (pages1.Count == 0 || pages2.Count == 0)
            return;

        var fingerprints1 = pages1.Select(PageFingerprint).ToList();
        var fingerprints2 = pages2.Select(PageFingerprint).ToList();
        var candidates = new List<(int I1, int I2, double Score)>();

        for (int i1 = 0; i1 < pages1.Count; i1++)
        for (int i2 = 0; i2 < pages2.Count; i2++)
        {
            double score = PageSim(fingerprints1[i1], fingerprints2[i2]);
            if (score >= MinRelocationSim)
                candidates.Add((i1, i2, score));
        }

        var uniqueCandidates = candidates
            .Where(candidate =>
            {
                double bestForFirst = candidates.Where(other => other.I1 == candidate.I1)
                    .Select(other => other.Score).DefaultIfEmpty(0).Max();
                double bestForSecond = candidates.Where(other => other.I2 == candidate.I2)
                    .Select(other => other.Score).DefaultIfEmpty(0).Max();
                double nextForFirst = candidates.Where(other => other.I1 == candidate.I1 && other.I2 != candidate.I2)
                    .Select(other => other.Score).DefaultIfEmpty(0).Max();
                double nextForSecond = candidates.Where(other => other.I2 == candidate.I2 && other.I1 != candidate.I1)
                    .Select(other => other.Score).DefaultIfEmpty(0).Max();

                return candidate.Score == bestForFirst && candidate.Score == bestForSecond &&
                       candidate.Score - nextForFirst >= MinRelocationMargin &&
                       candidate.Score - nextForSecond >= MinRelocationMargin;
            })
            .OrderByDescending(candidate => candidate.Score)
            .ToList();

        var byFirst = alignment.ToDictionary(entry => entry.I1);
        var bestByFirst = uniqueCandidates.ToDictionary(candidate => candidate.I1);
        foreach (var candidate in uniqueCandidates)
        {
            var current = byFirst[candidate.I1];
            if (current.I2 == candidate.I2)
                continue;

            if (current.I2 >= 0 && candidate.Score - current.Sim < MinRelocationMargin)
                continue;

            var previousOwner = alignment.FirstOrDefault(entry => entry.I2 == candidate.I2);
            if (previousOwner != null && previousOwner.I1 != candidate.I1)
            {
                if (!bestByFirst.TryGetValue(previousOwner.I1, out var replacement) ||
                    replacement.I2 == candidate.I2 ||
                    (previousOwner.I2 >= 0 && replacement.Score - previousOwner.Sim < MinRelocationMargin))
                    continue;

                var replacementOwner = alignment.FirstOrDefault(entry => entry.I2 == replacement.I2);
                if (replacementOwner != null && replacementOwner.I1 != previousOwner.I1 &&
                    replacementOwner.I1 != candidate.I1)
                    continue;

                previousOwner.I2 = replacement.I2;
                previousOwner.Sim = Math.Round(replacement.Score, 3);
                previousOwner.MatchType = "relocated";
            }

            current.I2 = candidate.I2;
            current.Sim = Math.Round(candidate.Score, 3);
            current.MatchType = "relocated";
        }
    }

    private static (List<string> Pages, List<int> OriginalIndices) CollapseStubs(List<string> pages)
    {
        var filtered = new List<string>();
        var indices = new List<int>();
        for (int i = 0; i < pages.Count; i++)
        {
            if (CountMeaningfulWords(pages[i]) >= StubWordThreshold)
            {
                filtered.Add(pages[i]);
                indices.Add(i);
            }
        }

        return (filtered, indices);
    }

    private static int CountMeaningfulWords(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return 0;

        return text
            .Split(new[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries)
            .Count(w => w.Length >= 3 && w.Any(char.IsLetter));
    }

    private List<AlignmentEntry> AlignContentPages(List<string> p1, List<string> p2)
    {
        var fps1 = p1.Select(PageFingerprint).ToList();
        var fps2 = p2.Select(PageFingerprint).ToList();
        int n1 = fps1.Count;
        int n2 = fps2.Count;

        var sim = new double[n1, n2];
        // Each row is independent — compute in parallel across available cores.
        System.Threading.Tasks.Parallel.For(0, n1, i =>
        {
            for (int j = 0; j < n2; j++)
                sim[i, j] = PageSim(fps1[i], fps2[j]);
        });

        var dp = new double[n1 + 1, n2 + 1];
        var move = new byte[n1 + 1, n2 + 1];

        for (int i = 1; i <= n1; i++)
        {
            for (int j = 1; j <= n2; j++)
            {
                double best = dp[i - 1, j];
                byte bestMove = 0;

                if (dp[i, j - 1] > best)
                {
                    best = dp[i, j - 1];
                    bestMove = 1;
                }

                double s = sim[i - 1, j - 1];
                if (s >= MinSim)
                {
                    double diag = dp[i - 1, j - 1] + s;
                    if (diag > best)
                    {
                        best = diag;
                        bestMove = 2;
                    }
                }

                dp[i, j] = best;
                move[i, j] = bestMove;
            }
        }

        var matched = new Dictionary<int, AlignmentEntry>();
        int ti = n1;
        int tj = n2;
        while (ti > 0 && tj > 0)
        {
            switch (move[ti, tj])
            {
                case 2:
                    matched[ti - 1] = new AlignmentEntry
                    {
                        I1 = ti - 1,
                        I2 = tj - 1,
                        Sim = Math.Round(sim[ti - 1, tj - 1], 3),
                    };
                    ti--;
                    tj--;
                    break;
                case 1:
                    tj--;
                    break;
                default:
                    ti--;
                    break;
            }
        }

        var result = new List<AlignmentEntry>(n1);
        for (int i = 0; i < n1; i++)
        {
            result.Add(matched.TryGetValue(i, out var e)
                ? e
                : new AlignmentEntry { I1 = i, I2 = -1, Sim = 0.0 });
        }

        return result;
    }

    private static PageFingerprintData PageFingerprint(string text)
    {
        var lines = ExtractMeaningfulLines(text);
        var fp = new PageFingerprintData();
        fp.Anchor = JoinChunk(lines.Take(5));
        foreach (var chunk in BuildChunks(lines))
        {
            fp.Chunks.Add(chunk);
        }

        foreach (var heading in lines.Where(IsSectionHeading).Take(12))
        {
            fp.Headings.Add(heading);
        }

        foreach (var word in WordToken.Matches(string.Join(" ", lines)).Select(m => m.Value))
        {
            fp.Words.Add(word);
        }

        foreach (var heading in fp.Headings)
        {
            var tag = BuildHeadingSignature(heading);
            if (tag.Length > 0)
            {
                fp.SectionTags.Add(tag);
            }
        }

        var layoutTag = BuildLayoutTag(lines);
        if (layoutTag.Length > 0)
        {
            fp.SectionTags.Add(layoutTag);
        }

        return fp;
    }

    private static List<string> ExtractMeaningfulLines(string text)
    {
        var lines = new List<string>();
        foreach (var raw in (text ?? string.Empty).Split('\n'))
        {
            string line = NormalizeLine(raw);
            if (line.Length < 5 || IsNoise(line))
            {
                continue;
            }

            if (lines.Count > 0 && lines[^1] == line)
            {
                continue;
            }

            lines.Add(line);
        }

        return lines;
    }

    private static string NormalizeLine(string line) =>
        WhiteSpace.Replace((line ?? string.Empty).Trim().ToLowerInvariant(), " ").Trim();

    private static IEnumerable<string> BuildChunks(List<string> lines)
    {
        if (lines.Count == 0)
        {
            yield break;
        }

        var starts = new SortedSet<int> { 0 };
        for (int i = 0; i < lines.Count; i++)
        {
            if (i % 4 == 0)
            {
                starts.Add(i);
            }

            if (IsSectionHeading(lines[i]))
            {
                starts.Add(i);
            }
        }

        var emitted = new HashSet<string>(StringComparer.Ordinal);
        foreach (int start in starts)
        {
            string chunk = JoinChunk(lines.Skip(start).Take(4));
            if (chunk.Length >= 20 && emitted.Add(chunk))
            {
                yield return chunk;
            }
        }
    }

    private static string JoinChunk(IEnumerable<string> lines)
    {
        string chunk = string.Join(" ", lines).Trim();
        return chunk.Length > 260 ? chunk[..260] : chunk;
    }

    private static bool IsSectionHeading(string line)
    {
        if (!SectionHeading.IsMatch(line))
        {
            return false;
        }

        int wordCount = line.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length;
        return wordCount is >= 2 and <= 14 && line.Length <= 140;
    }

    private static string BuildHeadingSignature(string line)
    {
        string norm = NormalizeLine(line);
        if (norm.Length == 0)
        {
            return string.Empty;
        }

        norm = Regex.Replace(norm, @"^(?:note\s+)?\d+(?:\.\d+)*\.?\s*", "");
        var tokens = Regex.Matches(norm, @"[a-z]{3,}").Select(m => m.Value).Take(6).ToList();
        return tokens.Count == 0 ? string.Empty : "heading:" + string.Join("-", tokens);
    }

    private static string BuildLayoutTag(List<string> lines)
    {
        if (lines.Count == 0)
        {
            return string.Empty;
        }

        int numericHeavy = lines.Count(l => Regex.Matches(l, @"\d").Count >= 4);
        int noteHeadings = lines.Count(IsSectionHeading);
        int narrativeLong = lines.Count(l => l.Length >= 90 && !Regex.IsMatch(l, @"\d"));

        if (numericHeavy >= Math.Max(3, lines.Count / 4))
        {
            return "layout:table-heavy";
        }

        if (narrativeLong >= Math.Max(3, lines.Count / 4))
        {
            return "layout:narrative-heavy";
        }

        if (noteHeadings >= 2)
        {
            return "layout:note-structured";
        }

        return "layout:mixed";
    }

    private static double PageSim(PageFingerprintData fp1, PageFingerprintData fp2)
    {
        double startSim = SequenceSimilarity(fp1.Anchor, fp2.Anchor);
        double chunkSim = BestSimilarity(fp1.Chunks, fp2.Chunks);
        double headingSim = BestSimilarity(fp1.Headings, fp2.Headings);
        double tagSim = TagSimilarity(fp1.SectionTags, fp2.SectionTags);

        int intersection = fp1.Words.Intersect(fp2.Words).Count();
        int union = fp1.Words.Union(fp2.Words).Count();
        int minCount = Math.Min(fp1.Words.Count, fp2.Words.Count);
        double containment = minCount == 0 ? 0.0 : (double)intersection / minCount;
        double jaccard = union == 0 ? 0.0 : (double)intersection / union;

        double score =
            0.15 * startSim +
            0.40 * chunkSim +
            0.15 * headingSim +
            0.10 * tagSim +
            0.12 * containment +
            0.08 * jaccard;

        if (chunkSim >= 0.88 && containment >= 0.45)
        {
            score += 0.08;
        }

        if (headingSim >= 0.90 && containment >= 0.40)
        {
            score += 0.05;
        }

        if (tagSim >= 0.99 && fp1.SectionTags.Count > 0)
        {
            score += 0.08;
        }

        return Math.Min(1.0, score);
    }

    private static double TagSimilarity(HashSet<string> left, HashSet<string> right)
    {
        if (left.Count == 0 || right.Count == 0)
        {
            return 0.0;
        }

        int intersection = left.Intersect(right).Count();
        int union = left.Union(right).Count();
        return union == 0 ? 0.0 : (double)intersection / union;
    }

    private static double BestSimilarity(List<string> left, List<string> right)
    {
        if (left.Count == 0 || right.Count == 0)
        {
            return 0.0;
        }

        double best = 0.0;
        foreach (var a in left)
        {
            foreach (var b in right)
            {
                double score = SequenceSimilarity(a, b);
                if (score > best)
                {
                    best = score;
                    if (best >= 0.995)
                    {
                        return best;
                    }
                }
            }
        }

        return best;
    }

    public static double SequenceSimilarity(string a, string b)
    {
        if (string.IsNullOrEmpty(a) && string.IsNullOrEmpty(b))
        {
            return 1.0;
        }

        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
        {
            return 0.0;
        }

        int matches = CountMatches(a, 0, a.Length, b, 0, b.Length);
        return 2.0 * matches / (a.Length + b.Length);
    }

    private static int CountMatches(string a, int a0, int a1, string b, int b0, int b1)
    {
        int best = 0;
        int bestI = a0;
        int bestJ = b0;
        for (int i = a0; i < a1; i++)
        {
            for (int j = b0; j < b1; j++)
            {
                int k = 0;
                while (i + k < a1 && j + k < b1 && a[i + k] == b[j + k])
                {
                    k++;
                }

                if (k > best)
                {
                    best = k;
                    bestI = i;
                    bestJ = j;
                }
            }
        }

        if (best == 0)
        {
            return 0;
        }

        int total = best;
        if (bestI > a0 && bestJ > b0)
        {
            total += CountMatches(a, a0, bestI, b, b0, bestJ);
        }

        if (bestI + best < a1 && bestJ + best < b1)
        {
            total += CountMatches(a, bestI + best, a1, b, bestJ + best, b1);
        }

        return total;
    }

    // Common function/stop words — a line made entirely of these is an extraction artefact.
    private static readonly HashSet<string> _functionWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "a", "an", "the", "and", "or", "but", "nor", "so", "yet",
        "in", "on", "at", "to", "for", "of", "as", "by", "up", "if",
        "is", "it", "its", "be", "do", "no", "not", "was", "are"
    };

    public static bool IsNoise(string line)
    {
        if (string.IsNullOrWhiteSpace(line) || line.Length < 4)
            return true;

        if (NoiseDashes.IsMatch(line))
            return true;

        if (NoisePageNum.IsMatch(line))
            return true;

        if (AdditionalPageRe.IsMatch(line.Trim()))
            return true;

        if (Regex.IsMatch(line, @"^\d{4}(?:\s+\d{4})+$"))
            return true;

        if (Regex.IsMatch(line, @"^(?:r'?0{3,}|r'000)(?:\s+(?:r'?0{3,}|r'000))+$", RegexOptions.IgnoreCase))
            return true;

        // Line whose canonical form (stripped of all punctuation and whitespace) is
        // shorter than 3 characters is not comparable content.
        var tokens = line.Trim().Split(new char[]{' ', '\t', ' '}, StringSplitOptions.RemoveEmptyEntries);
        string canonical = Regex.Replace(string.Join("", tokens), @"[^\w]", "").ToLowerInvariant();
        if (canonical.Length < 3)
            return true;

        // Lines composed entirely of common function/stop words carry no auditable
        // content (e.g. "or as a or", "and the or") and are extraction artefacts.
        if (tokens.Length <= 6 && tokens.All(t => _functionWords.Contains(t.Trim('.',',',';',':'))))
            return true;

        return false;
    }
}
