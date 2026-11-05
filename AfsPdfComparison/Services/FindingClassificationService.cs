using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

/// <summary>
/// Assigns a reproducible review classification to extracted line changes before
/// a PDF renderer can request coordinates. Structural references and uncertain OCR
/// tokens remain reviewable but are never approved as substantive highlights.
/// </summary>
public static class FindingClassificationService
{
    private sealed record NoteHeading(string Number, string Title);

    private static readonly Regex NoteHeadingPattern = new(
        @"^\s*(?:note\s+)?(?<number>\d+(?:\.\d+)*)[.)]?\s*[:\-]?\s+(?<title>[\p{L}].{2,})$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex NoteReferencePattern = new(
        @"\bnote\s+(?<number>\d+(?:\.\d+)*)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex PageReferencePattern = new(
        @"\bpages?\s+(?<number>\d+(?:\s*[-–—]\s*\d+)?)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex DotLeaderReferencePattern = new(
        @"\.{2,}\s*(?<number>\d{1,4})\s*$", RegexOptions.Compiled);
    private static readonly Regex UncertainOcrPattern = new(
        @"[\uFFFD\u25A1]|\d\s*\?|[?]\d", RegexOptions.Compiled);

    public static void Classify(ComparisonResult comparison)
    {
        var noteMap = BuildNoteMap(comparison.Report1.Pages, comparison.Report2.Pages);
        var pageMap = comparison.PageDiffs
            .Where(page => page.PageAfs1.HasValue && page.PageAfs2.HasValue)
            .GroupBy(page => page.PageAfs1!.Value)
            .Where(group => group.Select(page => page.PageAfs2!.Value).Distinct().Count() == 1)
            .ToDictionary(group => group.Key, group => group.First().PageAfs2!.Value);

        foreach (var page in comparison.PageDiffs)
        foreach (var diff in page.Diffs)
        {
            diff.FindingId = BuildFindingId(page.PairIndex, diff);
            diff.OriginalDocumentId = comparison.Report1.Filename;
            diff.RevisedDocumentId = comparison.Report2.Filename;
            diff.OriginalPageNumber = page.PageAfs1;
            diff.RevisedPageNumber = page.PageAfs2;
            diff.MatchedSectionId = $"pair-{page.PairIndex}";
            diff.OriginalValue = diff.Line1 ?? "";
            diff.RevisedValue = diff.Line2 ?? "";

            if (diff.Status is "same" or "same-cross-page" or "same-different-page")
            {
                diff.Confidence = 1.0;
                diff.SemanticClassification = "no-meaningful-change";
                diff.HighlightEligible = false;
                continue;
            }

            if (UncertainOcrPattern.IsMatch(diff.Line1 ?? "") ||
                UncertainOcrPattern.IsMatch(diff.Line2 ?? ""))
            {
                diff.Confidence = 0.30;
                diff.SemanticClassification = "extraction-error";
                diff.DifferenceType = "uncertain-extraction";
                diff.HighlightEligible = false;
                continue;
            }

            if (IsReferenceOnlyChange(diff, noteMap, pageMap))
            {
                diff.Confidence = 0.94;
                diff.SemanticClassification = "structural-change";
                diff.DifferenceType = "reference-renumbering";
                diff.StructuralChangeType = "reference-renumbering";
                diff.HighlightEligible = false;
                if (page.StructuralChangeType == "none")
                    page.StructuralChangeType = "reference-renumbering";
                continue;
            }

            if (diff.Status == "changed")
            {
                diff.Confidence = diff.Similarity;
                diff.SemanticClassification = diff.Similarity >= 0.62
                    ? "confirmed-change"
                    : "possible-change";
                diff.DifferenceType = AuditTextNormalizer.ExtractNormalizedNumbers(diff.Line1 ?? "")
                    .SetEquals(AuditTextNormalizer.ExtractNormalizedNumbers(diff.Line2 ?? ""))
                    ? "text"
                    : "financial-value-or-reference";
                diff.HighlightEligible = diff.Confidence >= 0.62;
                continue;
            }

            diff.Confidence = Math.Min(0.90, Math.Max(0.65, (diff.Line1 ?? diff.Line2 ?? "").Length / 100.0));
            diff.SemanticClassification = "confirmed-change";
            diff.DifferenceType = diff.Status == "added" ? "added-content" : "removed-content";
            diff.HighlightEligible = true;
        }
    }

    private static Dictionary<string, string> BuildNoteMap(List<string> firstPages, List<string> secondPages)
    {
        var first = ExtractNoteHeadings(firstPages);
        var second = ExtractNoteHeadings(secondPages);
        var candidates = (from oldHeading in first
                          from newHeading in second
                          let score = PageAlignmentService.SequenceSimilarity(
                              NormalizeHeading(oldHeading.Title), NormalizeHeading(newHeading.Title))
                          where score >= 0.84
                          select (Old: oldHeading, New: newHeading, Score: score))
            .ToList();

        return candidates.Where(candidate =>
                candidate.Score - candidates.Where(other => other.Old.Number == candidate.Old.Number && other.New.Number != candidate.New.Number)
                    .Select(other => other.Score).DefaultIfEmpty(0).Max() >= 0.08 &&
                candidate.Score - candidates.Where(other => other.New.Number == candidate.New.Number && other.Old.Number != candidate.Old.Number)
                    .Select(other => other.Score).DefaultIfEmpty(0).Max() >= 0.08)
            .GroupBy(candidate => candidate.Old.Number, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single().New.Number, StringComparer.OrdinalIgnoreCase);
    }

    private static List<NoteHeading> ExtractNoteHeadings(IEnumerable<string> pages)
    {
        var result = new Dictionary<string, NoteHeading>(StringComparer.OrdinalIgnoreCase);
        foreach (string line in pages.SelectMany(page => (page ?? "").Split('\n')))
        {
            var match = NoteHeadingPattern.Match(line.Trim());
            if (!match.Success)
                continue;

            string number = match.Groups["number"].Value;
            string title = match.Groups["title"].Value.Trim();
            if (title.Length < 4 || AuditTextNormalizer.ExtractNormalizedNumbers(title).Count > 1)
                continue;
            result.TryAdd(number, new NoteHeading(number, title));
        }

        return result.Values.ToList();
    }

    private static bool IsReferenceOnlyChange(
        LineDiffResult diff,
        IReadOnlyDictionary<string, string> noteMap,
        IReadOnlyDictionary<int, int> pageMap)
    {
        if (diff.Status != "changed")
            return false;

        string first = diff.Line1 ?? "";
        string second = diff.Line2 ?? "";
        bool hasMappedReference = false;

        var firstNotes = NoteReferencePattern.Matches(first).Select(match => match.Groups["number"].Value).ToList();
        var secondNotes = NoteReferencePattern.Matches(second).Select(match => match.Groups["number"].Value).ToList();
        if (firstNotes.Count != secondNotes.Count)
            return false;
        for (int index = 0; index < firstNotes.Count; index++)
        {
            if (!noteMap.TryGetValue(firstNotes[index], out string? mapped) ||
                !string.Equals(mapped, secondNotes[index], StringComparison.OrdinalIgnoreCase))
                return false;
            hasMappedReference = true;
        }

        var firstPages = PageReferencePattern.Matches(first)
            .SelectMany(match => ExtractPageNumbers(match.Groups["number"].Value)).ToList();
        var secondPages = PageReferencePattern.Matches(second)
            .SelectMany(match => ExtractPageNumbers(match.Groups["number"].Value)).ToList();
        if (firstPages.Count == 0 && secondPages.Count == 0)
        {
            var firstDotLeader = DotLeaderReferencePattern.Match(first);
            var secondDotLeader = DotLeaderReferencePattern.Match(second);
            if (firstDotLeader.Success && secondDotLeader.Success)
            {
                firstPages.Add(int.Parse(firstDotLeader.Groups["number"].Value));
                secondPages.Add(int.Parse(secondDotLeader.Groups["number"].Value));
            }
        }

        if (firstPages.Count != secondPages.Count)
            return false;
        for (int index = 0; index < firstPages.Count; index++)
        {
            if (!pageMap.TryGetValue(firstPages[index], out int mapped) || mapped != secondPages[index])
                return false;
            hasMappedReference = true;
        }

        if (!hasMappedReference)
            return false;

        string firstShape = NormalizeReferenceShape(first);
        string secondShape = NormalizeReferenceShape(second);
        return string.Equals(firstShape, secondShape, StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<int> ExtractPageNumbers(string value) =>
        Regex.Matches(value, @"\d+").Select(match => int.TryParse(match.Value, out int number) ? number : -1)
            .Where(number => number > 0);

    private static string NormalizeReferenceShape(string text)
    {
        string normalized = NoteReferencePattern.Replace(text ?? "", "note #");
        normalized = PageReferencePattern.Replace(normalized, match =>
        {
            string prefix = Regex.Match(match.Value, @"^\s*pages?", RegexOptions.IgnoreCase).Value;
            return prefix + " #";
        });
        normalized = DotLeaderReferencePattern.Replace(normalized, match =>
            match.Value[..match.Groups["number"].Index] + "#");
        return Regex.Replace(normalized, @"\s+", " ").Trim();
    }

    private static string NormalizeHeading(string value) =>
        Regex.Replace((value ?? "").ToLowerInvariant(), @"[^\p{L}\p{N}]+", " ").Trim();

    private static string BuildFindingId(int pairIndex, LineDiffResult diff)
    {
        string text = $"{pairIndex}|{diff.Line1}|{diff.Line2}";
        string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text)))[..16];
        return $"line-{hash}";
    }
}
