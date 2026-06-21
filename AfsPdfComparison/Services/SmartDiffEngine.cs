using System.Text.RegularExpressions;

namespace AfsPdfComparison.Services;

public enum SmartChangeType
{
    NumericChange,
    NarrativeRewrite,
    LineChanged,
    BlockAdded,
    BlockRemoved,
    FormattingOnly,
    Same
}

public class SmartDiff
{
    public SmartChangeType ChangeType { get; set; }
    public string Section { get; set; } = "";
    public string NoteKey { get; set; } = "";
    public string Label { get; set; } = "";
    public string OldText { get; set; } = "";
    public string NewText { get; set; } = "";
    public List<string> OldNumbers { get; set; } = new();
    public List<string> NewNumbers { get; set; } = new();
    public int PageAfs1 { get; set; }
    public int PageAfs2 { get; set; }
    public double Similarity { get; set; }
    public bool IsMaterial { get; set; }
    public string ChangedWordsDescription { get; set; } = "";
}

/// <summary>
/// Three-tier semantic diff engine for AFS content.
/// This is a foundation service and is intended to sit above SemanticBlockBuilder.
/// </summary>
public class SmartDiffEngine
{
    private const double ExactThreshold = 0.96;
    private const double RelatedThreshold = 0.50;

    public List<SmartDiff> Compare(
        List<SemanticBlock> blocks1,
        List<SemanticBlock> blocks2)
    {
        var results = new List<SmartDiff>();
        var used1 = new HashSet<int>();
        var used2 = new HashSet<int>();

        MatchNumericBlocks(blocks1, blocks2, used1, used2, results);
        MatchParagraphBlocks(blocks1, blocks2, used1, used2, results);
        MatchRemainingLines(blocks1, blocks2, used1, used2, results);

        for (int i = 0; i < blocks1.Count; i++)
        {
            if (!used1.Contains(i) && !ShouldIgnoreBlock(blocks1[i]))
                results.Add(MakeRemoved(blocks1[i]));
        }

        for (int i = 0; i < blocks2.Count; i++)
        {
            if (!used2.Contains(i) && !ShouldIgnoreBlock(blocks2[i]))
                results.Add(MakeAdded(blocks2[i]));
        }

        return results
            .Where(d => d.ChangeType != SmartChangeType.Same &&
                        d.ChangeType != SmartChangeType.FormattingOnly)
            .OrderBy(d => d.PageAfs1)
            .ThenBy(d => d.NoteKey)
            .ToList();
    }

    private void MatchNumericBlocks(
        List<SemanticBlock> b1,
        List<SemanticBlock> b2,
        HashSet<int> used1,
        HashSet<int> used2,
        List<SmartDiff> results)
    {
        var numericBlocks1 = b1
            .Select((b, i) => (b, i))
            .Where(x => x.b.Type is BlockType.TableRow or BlockType.SingleValue)
            .ToList();

        var numericBlocks2 = b2
            .Select((b, i) => (b, i))
            .Where(x => x.b.Type is BlockType.TableRow or BlockType.SingleValue)
            .ToList();

        foreach (var (nb2, idx2) in numericBlocks2)
        {
            if (used2.Contains(idx2))
                continue;

            var candidates = numericBlocks1
                .Where(x => !used1.Contains(x.i))
                .Where(x => SameContext(x.b, nb2))
                .Select(x => (x.b, x.i, score: LabelSimilarity(x.b.NormText, nb2.NormText)))
                .Where(x => x.score >= RelatedThreshold)
                .OrderByDescending(x => x.score)
                .ToList();

            if (candidates.Count == 0)
            {
                candidates = numericBlocks1
                    .Where(x => !used1.Contains(x.i))
                    .Select(x => (x.b, x.i, score: LabelSimilarity(x.b.NormText, nb2.NormText)))
                    .Where(x => x.score >= 0.80)
                    .OrderByDescending(x => x.score)
                    .Take(3)
                    .ToList();
            }

            if (candidates.Count == 0)
                continue;

            var (best1, bestIdx1, bestScore) = candidates[0];
            used1.Add(bestIdx1);
            used2.Add(idx2);
            results.Add(CompareNumericBlocks(best1, nb2, bestScore));
        }
    }

    private SmartDiff CompareNumericBlocks(SemanticBlock b1, SemanticBlock b2, double labelScore)
    {
        bool numsIdentical = b1.Numbers.SequenceEqual(b2.Numbers) ||
                             new HashSet<string>(b1.Numbers).SetEquals(new HashSet<string>(b2.Numbers));

        string label1 = StripNumbers(b1.NormText);
        string label2 = StripNumbers(b2.NormText);
        bool labelSame = label1 == label2 || LabelSimilarity(label1, label2) >= 0.90;

        if (numsIdentical && labelSame)
            return MakeSame(b1, b2);

        if (numsIdentical && !labelSame)
        {
            string c1 = AggressiveCanon(b1.RawText);
            string c2 = AggressiveCanon(b2.RawText);
            if (c1 == c2)
                return MakeFormattingOnly(b1, b2);
        }

        var diff = new SmartDiff
        {
            ChangeType = SmartChangeType.NumericChange,
            Section = b1.SectionTag.Length > 0 ? b1.SectionTag : b2.SectionTag,
            NoteKey = b1.NoteKey.Length > 0 ? b1.NoteKey : b2.NoteKey,
            Label = ExtractLabel(b1.RawText),
            OldText = b1.RawText,
            NewText = b2.RawText,
            OldNumbers = b1.Numbers,
            NewNumbers = b2.Numbers,
            PageAfs1 = b1.PageNumber,
            PageAfs2 = b2.PageNumber,
            Similarity = labelScore,
            IsMaterial = true
        };

        diff.ChangedWordsDescription = BuildNumericChangeSummary(diff.Label, b1.Numbers, b2.Numbers);
        return diff;
    }

    private void MatchParagraphBlocks(
        List<SemanticBlock> b1,
        List<SemanticBlock> b2,
        HashSet<int> used1,
        HashSet<int> used2,
        List<SmartDiff> results)
    {
        var paras1 = b1.Select((b, i) => (b, i))
            .Where(x => !used1.Contains(x.i) &&
                        x.b.Type == BlockType.Paragraph &&
                        x.b.NormText.Length >= 60)
            .ToList();

        var paras2 = b2.Select((b, i) => (b, i))
            .Where(x => !used2.Contains(x.i) &&
                        x.b.Type == BlockType.Paragraph &&
                        x.b.NormText.Length >= 60)
            .ToList();

        foreach (var (p2, idx2) in paras2)
        {
            if (used2.Contains(idx2))
                continue;

            var candidates = paras1
                .Where(x => !used1.Contains(x.i))
                .Select(x => (x.b, x.i, score: ParagraphTopicScore(x.b.NormText, p2.NormText)))
                .Where(x => x.score >= 0.40)
                .OrderByDescending(x => x.score)
                .Take(3)
                .ToList();

            if (candidates.Count == 0)
                continue;

            var (best1, bestIdx1, topicScore) = candidates[0];
            used1.Add(bestIdx1);
            used2.Add(idx2);
            results.Add(CompareParagraphBlocks(best1, p2, topicScore));
        }
    }

    private SmartDiff CompareParagraphBlocks(SemanticBlock p1, SemanticBlock p2, double topicScore)
    {
        if (p1.NormText == p2.NormText)
            return MakeSame(p1, p2);

        double fullSim = PageAlignmentService.SequenceSimilarity(p1.NormText, p2.NormText);
        if (fullSim >= ExactThreshold)
            return MakeSame(p1, p2);

        if (IsLineWrapArtefact(p1.NormText, p2.NormText))
            return MakeSame(p1, p2);

        var n1 = new HashSet<string>(p1.Numbers);
        var n2 = new HashSet<string>(p2.Numbers);
        if (n1.SetEquals(n2) && fullSim >= 0.85)
            return MakeSame(p1, p2);

        bool isMaterial =
            !n1.SetEquals(n2) ||
            fullSim < 0.72 ||
            p1.NoteKey.Length > 0 ||
            p2.NoteKey.Length > 0;

        return new SmartDiff
        {
            ChangeType = SmartChangeType.NarrativeRewrite,
            Section = p1.SectionTag.Length > 0 ? p1.SectionTag : p2.SectionTag,
            NoteKey = p1.NoteKey.Length > 0 ? p1.NoteKey : p2.NoteKey,
            Label = SummariseSection(p1.SectionTag, p1.NoteKey),
            OldText = p1.RawText,
            NewText = p2.RawText,
            OldNumbers = p1.Numbers,
            NewNumbers = p2.Numbers,
            PageAfs1 = p1.PageNumber,
            PageAfs2 = p2.PageNumber,
            Similarity = Math.Max(topicScore, fullSim),
            IsMaterial = isMaterial,
            ChangedWordsDescription = BuildNarrativeChangeSummary(p1.NormText, p2.NormText, fullSim)
        };
    }

    private void MatchRemainingLines(
        List<SemanticBlock> b1,
        List<SemanticBlock> b2,
        HashSet<int> used1,
        HashSet<int> used2,
        List<SmartDiff> results)
    {
        var rem1 = b1.Select((b, i) => (b, i)).Where(x => !used1.Contains(x.i)).ToList();
        var rem2 = b2.Select((b, i) => (b, i)).Where(x => !used2.Contains(x.i)).ToList();

        var index1 = BuildInvertedIndex(rem1.Select(x => (x.b.NormText, x.i)).ToList());
        var availableIds = rem1.Select(x => x.i).ToHashSet();

        foreach (var (b2block, idx2) in rem2)
        {
            if (used2.Contains(idx2))
                continue;

            var candidates = GetCandidates(b2block.NormText, index1, used1, availableIds);
            double bestScore = 0.0;
            int bestIdx = -1;

            foreach (var ci in candidates.Take(40))
            {
                var match = rem1.FirstOrDefault(x => x.i == ci);
                if (match.b == null)
                    continue;

                double s = SequenceSim(b2block.NormText, match.b.NormText);
                if (s > bestScore)
                {
                    bestScore = s;
                    bestIdx = ci;
                }
            }

            if (bestIdx < 0 || bestScore < RelatedThreshold)
                continue;

            var b1block = b1[bestIdx];
            used1.Add(bestIdx);
            used2.Add(idx2);

            var status = DetermineLineStatus(
                b1block.NormText,
                b2block.NormText,
                bestScore,
                b1block.Numbers,
                b2block.Numbers);

            if (status is SmartChangeType.Same or SmartChangeType.FormattingOnly)
                continue;

            results.Add(new SmartDiff
            {
                ChangeType = status,
                Section = b1block.SectionTag.Length > 0 ? b1block.SectionTag : b2block.SectionTag,
                NoteKey = b1block.NoteKey.Length > 0 ? b1block.NoteKey : b2block.NoteKey,
                Label = ExtractLabel(b1block.RawText),
                OldText = b1block.RawText,
                NewText = b2block.RawText,
                OldNumbers = b1block.Numbers,
                NewNumbers = b2block.Numbers,
                PageAfs1 = b1block.PageNumber,
                PageAfs2 = b2block.PageNumber,
                Similarity = bestScore,
                IsMaterial = b1block.Numbers.Count > 0 || b2block.Numbers.Count > 0,
                ChangedWordsDescription = BuildWordDiffSummary(b1block.NormText, b2block.NormText)
            });
        }
    }

    private static SmartChangeType DetermineLineStatus(
        string n1,
        string n2,
        double score,
        List<string> nums1,
        List<string> nums2)
    {
        if (n1 == n2)
            return SmartChangeType.Same;

        string c1 = AggressiveCanon(n1);
        string c2 = AggressiveCanon(n2);
        if (c1 == c2)
            return SmartChangeType.FormattingOnly;

        if (score >= ExactThreshold)
            return SmartChangeType.Same;

        var ns1 = new HashSet<string>(nums1);
        var ns2 = new HashSet<string>(nums2);
        if (ns1.SetEquals(ns2) && score >= 0.80)
            return SmartChangeType.Same;

        if (IsLineWrapArtefact(n1, n2))
            return SmartChangeType.Same;

        var w1 = new HashSet<string>(
            n1.Split(' ', StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);
        var w2 = new HashSet<string>(
            n2.Split(' ', StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);

        if (w1.SetEquals(w2) && ns1.SetEquals(ns2))
            return SmartChangeType.FormattingOnly;

        return SmartChangeType.LineChanged;
    }

    private static bool SameContext(SemanticBlock b1, SemanticBlock b2)
    {
        if (b1.NoteKey.Length > 0 && b2.NoteKey.Length > 0)
        {
            return b1.NoteKey == b2.NoteKey ||
                   b1.NoteKey.StartsWith(b2.NoteKey) ||
                   b2.NoteKey.StartsWith(b1.NoteKey);
        }

        if (b1.SectionTag.Length > 0 && b2.SectionTag.Length > 0)
            return b1.SectionTag == b2.SectionTag;

        return true;
    }

    private static double LabelSimilarity(string a, string b)
    {
        string la = StripNumbers(a);
        string lb = StripNumbers(b);
        if (la.Length == 0 && lb.Length == 0)
            return 1.0;
        if (la.Length == 0 || lb.Length == 0)
            return 0.0;
        return PageAlignmentService.SequenceSimilarity(la, lb);
    }

    private static double ParagraphTopicScore(string p1, string p2)
    {
        if (p1.Length == 0 || p2.Length == 0)
            return 0.0;

        int len1 = Math.Max(20, (int)(p1.Length * 0.40));
        int len2 = Math.Max(20, (int)(p2.Length * 0.40));
        string open1 = p1[..Math.Min(len1, p1.Length)];
        string open2 = p2[..Math.Min(len2, p2.Length)];
        double openSim = PageAlignmentService.SequenceSimilarity(open1, open2);

        var vocab1 = ContentWords(p1);
        var vocab2 = ContentWords(p2);
        int common = vocab1.Intersect(vocab2).Count();
        int union = vocab1.Union(vocab2).Count();
        double vocabSim = union == 0 ? 0.0 : (double)common / union;

        return openSim * 0.55 + vocabSim * 0.45;
    }

    private static HashSet<string> ContentWords(string text) =>
        new(
            text.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(w => w.Length >= 5 &&
                            w.Any(char.IsLetter) &&
                            !Regex.IsMatch(w, @"^\d+$"))
                .Select(w => Regex.Replace(w.ToLowerInvariant(), @"^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$", ""))
                .Where(w => w.Length >= 5),
            StringComparer.OrdinalIgnoreCase);

    private static bool IsLineWrapArtefact(string n1, string n2)
    {
        var w1 = new HashSet<string>(
            n1.Split(' ', StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);
        var w2 = new HashSet<string>(
            n2.Split(' ', StringSplitOptions.RemoveEmptyEntries),
            StringComparer.OrdinalIgnoreCase);

        if (w1.Count < 4 || w2.Count < 4)
            return false;

        var shorter = w1.Count <= w2.Count ? w1 : w2;
        var longer = w1.Count <= w2.Count ? w2 : w1;
        string sText = w1.Count <= w2.Count ? n1 : n2;
        string lText = w1.Count <= w2.Count ? n2 : n1;

        int hits = shorter.Count(w => longer.Contains(w));
        double overlap = (double)hits / shorter.Count;
        double lenRatio = (double)shorter.Count / longer.Count;

        bool isPrefix = lText.StartsWith(sText, StringComparison.OrdinalIgnoreCase) &&
                        !sText[^1..].Contains('.');

        return (overlap >= 0.85 && lenRatio >= 0.55) || isPrefix;
    }

    private static string StripNumbers(string text) =>
        Regex.Replace(text, @"\b\d[\d\s,\.]*\b", " ").Replace("  ", " ").Trim();

    private static string AggressiveCanon(string text)
    {
        string t = TextNormaliser.Canonicalise(text);
        t = Regex.Replace(t, @"[\(\)\[\]]", "");
        t = Regex.Replace(t, @"\s+", " ").Trim();
        return t;
    }

    private static string ExtractLabel(string rawText)
    {
        var m = Regex.Match(rawText, @"^([^\d\(\-]+)");
        return m.Success ? m.Groups[1].Value.Trim() : rawText.Trim();
    }

    private static string SummariseSection(string sectionTag, string noteKey)
    {
        if (noteKey.Length > 0)
            return $"Note {noteKey}";
        return sectionTag.Replace("-", " ").ToUpperInvariant();
    }

    private static string BuildNumericChangeSummary(string label, List<string> old, List<string> nw)
    {
        if (old.Count == 0 && nw.Count == 0)
            return "Formatting change";

        var added = nw.Except(old).ToList();
        var removed = old.Except(nw).ToList();

        if (removed.Count == 1 && added.Count == 1)
            return $"{label.Trim()}: {removed[0]} -> {added[0]}";

        var parts = new List<string>();
        if (removed.Count > 0)
            parts.Add($"removed: {string.Join(", ", removed)}");
        if (added.Count > 0)
            parts.Add($"added: {string.Join(", ", added)}");
        return string.Join("; ", parts);
    }

    private static string BuildNarrativeChangeSummary(string old, string nw, double sim)
    {
        int pct = (int)((1 - sim) * 100);
        var w1 = old.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var w2 = nw.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < Math.Min(w1.Length, w2.Length); i++)
        {
            if (!string.Equals(w1[i], w2[i], StringComparison.OrdinalIgnoreCase))
                return $"~{pct}% rewritten. First difference near: \"{w1[i]}\" -> \"{w2[i]}\"";
        }

        return $"~{pct}% rewritten";
    }

    private static string BuildWordDiffSummary(string old, string nw)
    {
        var w1 = old.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(w => w.ToLowerInvariant())
            .ToList();
        var w2 = nw.Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(w => w.ToLowerInvariant())
            .ToList();
        var set1 = new HashSet<string>(w1);
        var set2 = new HashSet<string>(w2);
        var removed = set1.Except(set2).Where(w => w.Length > 3).Take(5).ToList();
        var added = set2.Except(set1).Where(w => w.Length > 3).Take(5).ToList();
        var parts = new List<string>();
        if (removed.Count > 0)
            parts.Add($"removed words: {string.Join(", ", removed)}");
        if (added.Count > 0)
            parts.Add($"added words: {string.Join(", ", added)}");
        return parts.Count > 0 ? string.Join("; ", parts) : "minor wording change";
    }

    private static bool ShouldIgnoreBlock(SemanticBlock b) =>
        b.RawText.Length < 5 ||
        (b.Type == BlockType.Other && b.RawText.Length < 10);

    private static SmartDiff MakeSame(SemanticBlock b1, SemanticBlock b2) =>
        new()
        {
            ChangeType = SmartChangeType.Same,
            PageAfs1 = b1.PageNumber,
            PageAfs2 = b2.PageNumber
        };

    private static SmartDiff MakeFormattingOnly(SemanticBlock b1, SemanticBlock b2) =>
        new()
        {
            ChangeType = SmartChangeType.FormattingOnly,
            OldText = b1.RawText,
            NewText = b2.RawText,
            PageAfs1 = b1.PageNumber,
            PageAfs2 = b2.PageNumber
        };

    private static SmartDiff MakeAdded(SemanticBlock b) =>
        new()
        {
            ChangeType = SmartChangeType.BlockAdded,
            Section = b.SectionTag,
            NoteKey = b.NoteKey,
            NewText = b.RawText,
            NewNumbers = b.Numbers,
            PageAfs2 = b.PageNumber,
            IsMaterial = b.Numbers.Count > 0,
            ChangedWordsDescription = "New content added"
        };

    private static SmartDiff MakeRemoved(SemanticBlock b) =>
        new()
        {
            ChangeType = SmartChangeType.BlockRemoved,
            Section = b.SectionTag,
            NoteKey = b.NoteKey,
            OldText = b.RawText,
            OldNumbers = b.Numbers,
            PageAfs1 = b.PageNumber,
            IsMaterial = b.Numbers.Count > 0,
            ChangedWordsDescription = "Content removed"
        };

    private static Dictionary<string, HashSet<int>> BuildInvertedIndex(
        List<(string text, int idx)> lines)
    {
        var index = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (text, i) in lines)
        {
            var tokens = text.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length > 0)
            {
                if (!index.ContainsKey(tokens[0]))
                    index[tokens[0]] = new();
                index[tokens[0]].Add(i);
            }

            for (int j = 0; j < tokens.Length - 1; j++)
            {
                var key = tokens[j] + tokens[j + 1];
                if (!index.ContainsKey(key))
                    index[key] = new();
                index[key].Add(i);
            }
        }

        return index;
    }

    private static List<int> GetCandidates(
        string needle,
        Dictionary<string, HashSet<int>> index,
        HashSet<int> used,
        HashSet<int> availableIds)
    {
        var tokens = needle.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var hits = new HashSet<int>();
        if (tokens.Length > 0 && index.TryGetValue(tokens[0], out var h0))
            hits.UnionWith(h0);

        for (int j = 0; j < tokens.Length - 1; j++)
        {
            var key = tokens[j] + tokens[j + 1];
            if (index.TryGetValue(key, out var hb))
                hits.UnionWith(hb);
        }

        hits.ExceptWith(used);
        if (hits.Count == 0)
            hits = availableIds.Except(used).ToHashSet();

        return hits.OrderBy(x => x).ToList();
    }

    private static double SequenceSim(string a, string b) =>
        PageAlignmentService.SequenceSimilarity(a, b);
}
