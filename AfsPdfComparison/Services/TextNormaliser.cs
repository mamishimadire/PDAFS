// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  SERVICE — TextNormaliser                                                    ║
// ║  Centralised text-normalisation utilities shared by PdfExtractionService,   ║
// ║  LineComparisonService, and PageSnapshotService.                             ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • Python notebook v4.3 supplement — TextNormaliser                        ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text.RegularExpressions;

namespace AfsPdfComparison.Services
{
    public static class TextNormaliser
    {
        // Matches pure page-stamp lines: "- 40 -", "(40)", "40 of 120", "page 3 of 5"
        private static readonly Regex PageStampRegex = new(
            @"^[-–—\s]*\(?\s*\d{1,4}\s*\)?[-–—\s]*$|^page\s+\d{1,4}(\s+of\s+\d{1,4})?$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Matches lines that are purely decorative separators
        private static readonly Regex SeparatorRegex = new(
            @"^[\s\-_=|\.]+$",
            RegexOptions.Compiled);

        // Matches numeric tokens including spaced or comma-formatted numbers
        private static readonly Regex NumberTokenRegex = new(
            @"\b\d[\d\s,\.]*\d\b|\b\d\b",
            RegexOptions.Compiled);

        /// <summary>
        /// Returns true if the line should be discarded before comparison.
        /// Filters very short lines, separator lines, and page-stamp lines.
        /// </summary>
        public static bool IsNoise(string line)
        {
            var t = line?.Trim() ?? "";
            if (t.Length < 4) return true;
            if (SeparatorRegex.IsMatch(t)) return true;
            if (PageStampRegex.IsMatch(t)) return true;
            return false;
        }

        /// <summary>
        /// Strips whitespace/punctuation/number-format differences so that
        /// "R 1 234 567" and "R1,234,567" compare as equal.
        /// </summary>
        public static string Canonicalise(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var t = s.Trim().ToLowerInvariant();
            // Collapse all whitespace (including non-breaking space) to single space
            t = Regex.Replace(t, @"[\s\u00a0]+", " ");
            // Remove whitespace between digits: "102 575" → "102575"
            t = Regex.Replace(t, @"(?<=\d)\s+(?=\d)", "");
            // Remove comma thousand-separators: "102,575" → "102575"
            t = Regex.Replace(t, @"(?<=\d),(?=\d{3})", "");
            // Remove punctuation
            t = Regex.Replace(t, @"[^\w\s]", "");
            return t.Trim();
        }

        /// <summary>
        /// Extracts unique canonical numeric tokens from text.
        /// Strips page-stamp lines first so footer numbers are excluded.
        /// </summary>
        public static HashSet<string> ExtractNumbers(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return new();
            // Strip page-stamp lines before scanning for numbers
            var cleaned = string.Join("\n",
                text.Split('\n')
                    .Where(l => !PageStampRegex.IsMatch(l.Trim())));
            var result = new HashSet<string>();
            foreach (Match m in NumberTokenRegex.Matches(cleaned))
            {
                var canon = Regex.Replace(m.Value, @"[\s,]", "");
                if (Regex.IsMatch(canon, @"^\d+(\.\d+)?$"))
                    result.Add(canon);
            }
            return result;
        }
    }
}
