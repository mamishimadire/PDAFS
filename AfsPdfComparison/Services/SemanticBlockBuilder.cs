using System.Text.RegularExpressions;

namespace AfsPdfComparison.Services;

public enum BlockType
{
    NoteHeading,
    Paragraph,
    TableRow,
    TableHeader,
    SectionHeading,
    SingleValue,
    DateLine,
    Other
}

public class SemanticBlock
{
    public BlockType Type { get; set; }
    public string RawText { get; set; } = "";
    public string NormText { get; set; } = "";
    public int PageNumber { get; set; }
    public int LineIndex { get; set; }
    public List<string> Numbers { get; set; } = new();
    public string NoteKey { get; set; } = "";
    public string SectionTag { get; set; } = "";
    public bool IsHeader { get; set; }
}

/// <summary>
/// Converts raw extracted page text into typed semantic blocks.
/// Section identity is derived from headings and note numbering found in the
/// document itself rather than from predefined phrase lists.
/// </summary>
public static class SemanticBlockBuilder
{
    private static readonly Regex NotePrefix = new(
        @"^(?:note\s+)?(\d+(?:\.\d+)*\.?)\s+\S",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex AllCaps = new(
        @"^[A-Z][A-Z\s\(\)\/\-]{8,}$",
        RegexOptions.Compiled);

    private static readonly Regex NumPattern = new(
        @"\b\d{1,3}(?:[\s,]\d{3})*(?:\.\d+)?\b|\b\d{4,}\b",
        RegexOptions.Compiled);

    private static readonly Regex DatePattern = new(
        @"\b\d{1,2}\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\w*\s+\d{4}\b" +
        @"|\b\d{4}[-/]\d{2}[-/]\d{2}\b" +
        @"|\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    public static List<SemanticBlock> Build(string pageText, int pageNumber)
    {
        var rawLines = (pageText ?? "")
            .Split('\n')
            .Select(l => l.Trim())
            .Where(l => l.Length >= 3 && !TextNormaliser.IsNoise(l))
            .ToList();

        var blocks = new List<SemanticBlock>();
        string currentNoteKey = "";
        string currentSectionTag = "";

        for (int i = 0; i < rawLines.Count; i++)
        {
            string raw = rawLines[i];
            string norm = TextNormaliser.Normalise(raw);
            if (string.IsNullOrWhiteSpace(norm))
                continue;

            var numbers = ExtractNumberList(raw);
            var type = ClassifyLine(raw, norm, numbers);

            var noteMatch = NotePrefix.Match(raw);
            if (noteMatch.Success)
                currentNoteKey = noteMatch.Groups[1].Value.TrimEnd('.');

            if (type is BlockType.NoteHeading or BlockType.SectionHeading)
                currentSectionTag = BuildSectionSignature(raw, norm, currentNoteKey);
            else if (string.IsNullOrWhiteSpace(currentSectionTag))
                currentSectionTag = BuildContextSignature(norm, numbers, type);

            blocks.Add(new SemanticBlock
            {
                Type = type,
                RawText = raw,
                NormText = norm,
                PageNumber = pageNumber,
                LineIndex = i,
                Numbers = numbers,
                NoteKey = currentNoteKey,
                SectionTag = currentSectionTag,
                IsHeader = type is BlockType.NoteHeading or BlockType.SectionHeading
            });
        }

        return MergeParagraphs(blocks);
    }

    private static string BuildSectionSignature(string raw, string norm, string noteKey)
    {
        string core = Regex.Replace(norm, @"^(?:note\s+)?\d+(?:\.\d+)*\.?\s*", "");
        var tokens = Regex.Matches(core, @"[a-z]{3,}")
            .Select(m => m.Value)
            .Take(6)
            .ToList();

        if (tokens.Count == 0)
            return noteKey.Length > 0 ? $"note:{noteKey}" : string.Empty;

        string headingTag = string.Join("-", tokens);
        return noteKey.Length > 0
            ? $"note:{noteKey}:{headingTag}"
            : $"heading:{headingTag}";
    }

    private static string BuildContextSignature(string norm, List<string> numbers, BlockType type)
    {
        if (type == BlockType.TableRow)
            return numbers.Count >= 3 ? "layout:table-wide" : "layout:table-row";

        if (type == BlockType.Paragraph)
            return "layout:paragraph";

        if (type == BlockType.SingleValue)
            return "layout:single-value";

        if (type == BlockType.DateLine)
            return "layout:date-line";

        return string.Empty;
    }

    private static BlockType ClassifyLine(string raw, string norm, List<string> numbers)
    {
        if (DatePattern.IsMatch(raw))
            return BlockType.DateLine;

        if (NotePrefix.IsMatch(raw))
            return BlockType.NoteHeading;

        if (AllCaps.IsMatch(raw.Trim()) && numbers.Count == 0)
            return BlockType.SectionHeading;

        bool hasLabel = norm.Any(char.IsLetter);
        if (numbers.Count >= 2 && hasLabel)
            return BlockType.TableRow;

        if (numbers.Count == 1 && hasLabel)
            return BlockType.SingleValue;

        if (norm.Length > 80 && numbers.Count == 0)
            return BlockType.Paragraph;

        return BlockType.Other;
    }

    private static List<string> ExtractNumberList(string text)
    {
        var result = new List<string>();
        foreach (Match m in NumPattern.Matches(text ?? ""))
        {
            string canon = Regex.Replace(m.Value, @"[\s,]", "");
            if (Regex.IsMatch(canon, @"^\d+(?:\.\d+)?$") && canon.Length >= 2)
                result.Add(canon);
        }

        return result;
    }

    private static List<SemanticBlock> MergeParagraphs(List<SemanticBlock> blocks)
    {
        if (blocks.Count == 0)
            return blocks;

        var result = new List<SemanticBlock> { blocks[0] };

        for (int i = 1; i < blocks.Count; i++)
        {
            var prev = result[^1];
            var curr = blocks[i];

            bool prevIsPara = prev.Type is BlockType.Paragraph or BlockType.Other;
            bool currCont = curr.RawText.Length > 0 && char.IsLower(curr.RawText[0]);
            bool noNums = curr.Numbers.Count == 0 && prev.Numbers.Count == 0;
            bool prevNoStop = prev.RawText.Length > 0 &&
                              prev.RawText[^1] is not '.' and not '!' and not '?';

            if (prevIsPara && currCont && noNums && prevNoStop && prev.RawText.Length >= 15)
            {
                prev.RawText += " " + curr.RawText;
                prev.NormText += " " + curr.NormText;
                prev.Type = BlockType.Paragraph;
            }
            else
            {
                result.Add(curr);
            }
        }

        return result;
    }
}
