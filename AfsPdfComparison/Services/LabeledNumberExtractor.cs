// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  SERVICE — LabeledNumberExtractor                                            ║
// ║  Extracts labeled financial figures from AFS page text for structured        ║
// ║  comparison and change detection without AI.                                 ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text.RegularExpressions;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

/// <summary>
/// Extracts label→number pairs from financial table rows in AFS page text.
/// Used by AutoReviewEngine to build a structured view of all financial figures.
/// </summary>
public class LabeledNumberExtractor
{
    // Matches: text label + 2+ spaces + numeric content
    // Label: 3–70 chars of text (letters, spaces, punctuation, but not just numbers)
    // Number: parenthetical or plain, with space/comma/dot separators
    private static readonly Regex TableRowRe = new(
        @"^(.{3,70}?)\s{2,}((?:\([\d\s,\.]+\)|[\d][\d\s,\.]{1,})(?:\s+(?:\([\d\s,\.]+\)|[\d][\d\s,\.]+))*)$",
        RegexOptions.Compiled);

    // Extract the FIRST number token from a string (handles parenthetical negatives)
    private static readonly Regex FirstNumRe = new(
        @"(\()([\d][\d\s,\.]*\d)(\))|(\b)([\d][\d\s,]{0,}[\d])(\b)",
        RegexOptions.Compiled);

    public sealed class LabeledNumber
    {
        public string  Label     { get; init; } = "";
        public string  Raw       { get; init; } = "";
        public decimal Value     { get; init; }
        public int     PageIndex { get; init; }  // 0-based
        public int?    PairIndex { get; init; }
    }

    /// <summary>Extract labeled numbers from all pages.</summary>
    public List<LabeledNumber> Extract(List<string> pages, List<int?>? pairIndices = null)
    {
        var result = new List<LabeledNumber>(64);
        for (int p = 0; p < pages.Count; p++)
        {
            int? pairIdx = pairIndices != null && p < pairIndices.Count ? pairIndices[p] : null;
            foreach (var line in pages[p].Split('\n', StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.Trim();
                if (t.Length < 6) continue;

                var m = TableRowRe.Match(t);
                if (!m.Success) continue;

                string label = m.Groups[1].Value.Trim();
                string numStr = m.Groups[2].Value.Trim();

                // Label must have at least 2 alphabetic chars
                if (label.Count(char.IsLetter) < 2) continue;

                // Don't treat all-number labels as table rows
                if (Regex.IsMatch(label, @"^\d[\d\s\.]*$")) continue;

                var (val, raw) = ParseFirstNumber(numStr);
                if (raw.Length == 0) continue;

                result.Add(new LabeledNumber
                {
                    Label     = NormaliseLabel(label),
                    Raw       = raw,
                    Value     = val,
                    PageIndex = p,
                    PairIndex = pairIdx,
                });
            }
        }
        return result;
    }

    /// <summary>
    /// Compare two sets of labeled numbers and return all changes.
    /// Matches by label (case-insensitive, normalised).
    /// </summary>
    public List<LabeledNumberChange> Compare(
        List<LabeledNumber> set1,
        List<LabeledNumber> set2,
        List<PageDiffResult>? pageDiffs = null)
    {
        var dict1 = set1.GroupBy(n => n.Label, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);
        var dict2 = set2.GroupBy(n => n.Label, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.OrdinalIgnoreCase);

        var changes = new List<LabeledNumberChange>();

        foreach (var (label, nums1) in dict1)
        {
            if (!dict2.TryGetValue(label, out var nums2)) continue;

            // Take the first occurrence from each side
            var n1 = nums1[0];
            var n2 = nums2[0];

            if (n1.Value == n2.Value) continue;   // no change

            decimal abs = Math.Abs(n2.Value - n1.Value);
            double pct  = n1.Value == 0 ? (n2.Value == 0 ? 0 : 1.0)
                                        : (double)Math.Abs(n2.Value - n1.Value) / (double)Math.Abs(n1.Value);

            bool signChanged = (n1.Value > 0 && n2.Value < 0) || (n1.Value < 0 && n2.Value > 0);

            string severity;
            if (signChanged || pct >= 0.20 || abs >= 500_000)
                severity = "high";
            else if (pct >= 0.05 || abs >= 50_000)
                severity = "medium";
            else
                severity = "low";

            // Find the pairIndex from pageDiffs if available
            int? pairIdx = null;
            if (pageDiffs != null)
            {
                var pair = pageDiffs.FirstOrDefault(p =>
                    p.PageAfs1 == n1.PageIndex + 1 || p.PageAfs2 == n2.PageIndex + 1);
                pairIdx = pair?.PairIndex;
            }

            changes.Add(new LabeledNumberChange
            {
                Label       = label,
                RawBefore   = n1.Raw,
                RawAfter    = n2.Raw,
                Before      = n1.Value,
                After       = n2.Value,
                PctChange   = pct,
                AbsChange   = abs,
                Severity    = severity,
                PageAfs1    = n1.PageIndex + 1,
                PageAfs2    = n2.PageIndex + 1,
                PairIndex   = pairIdx,
                SignChanged = signChanged,
            });
        }

        return changes.OrderByDescending(c => c.Severity == "high" ? 2 : c.Severity == "medium" ? 1 : 0)
                      .ThenByDescending(c => c.AbsChange)
                      .ToList();
    }

    // ── helpers ──────────────────────────────────────────────────────────────
    private static string NormaliseLabel(string label)
    {
        // Collapse whitespace, lowercase for matching
        return Regex.Replace(label.Trim(), @"\s+", " ").ToLowerInvariant();
    }

    private static (decimal value, string raw) ParseFirstNumber(string numStr)
    {
        string s = numStr.Trim();
        if (s.Length == 0) return (0, "");

        // Check parenthetical negative: (1 234 000)
        bool negative = s.StartsWith('(') && s.EndsWith(')');
        string inner = negative ? s[1..^1].Trim() : s;

        // Take just the first whitespace-delimited token
        int spaceIdx = inner.IndexOf(' ');
        if (spaceIdx > 0)
        {
            // Check if it's a spaced-thousands: "1 234 000" → all digits after spaces
            // If so keep the full thing; otherwise take just the first token
            string rest = inner[(spaceIdx + 1)..];
            if (rest.All(c => char.IsDigit(c) || c == ' ' || c == ','))
            {
                // space-separated thousands — keep full inner
            }
            else
            {
                inner = inner[..spaceIdx];
            }
        }

        // Remove commas and spaces
        string clean = Regex.Replace(inner, @"[\s,]", "");
        if (!decimal.TryParse(clean, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out decimal val))
            return (0, "");

        return (negative ? -val : val, s.Split(' ')[0]);
    }
}
