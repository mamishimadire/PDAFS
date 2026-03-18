// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  SERVICE — PdfExtractionService                                              ║
// ║  Byte-array-based PDF extraction using PdfPig + TextNormaliser.             ║
// ║  Provides a PdfExtractionResult compatible with the new                     ║
// ║  LineComparisonService / PageSnapshotService pipeline.                      ║
// ║                                                                              ║
// ║  NOTE: This service is additive alongside the existing PageExtractorService. ║
// ║   The existing PageExtractorService operates on file paths and populates    ║
// ║   PdfReport objects used by the main app pipeline.  This service accepts    ║
// ║   byte arrays and returns the lighter PdfExtractionResult which is used     ║
// ║   by ComparisonController and PageSnapshotService.                          ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text;
using UglyToad.PdfPig;

namespace AfsPdfComparison.Services
{
    public class PdfExtractionResult
    {
        public string   Filename      { get; set; } = "";
        public List<string> Pages     { get; set; } = new();
        public string   FullText      => string.Join("\n", Pages);
        public int      PageCount     => Pages.Count;
        public int      WordCount     { get; set; }
        public HashSet<string> UniqueNumbers { get; set; } = new();
    }

    public class PdfExtractionService
    {
        // Strip invisible Unicode chars (same set as TextNormaliser / PageExtractorService).
        // \uFFF0-\uFFFD covers the Unicode Specials block including \uFFFD (REPLACEMENT
        // CHARACTER, category So, NOT \p{C}), which PdfPig emits for undecodable glyphs.
        private static readonly System.Text.RegularExpressions.Regex _invisRe =
            new(@"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]",
                System.Text.RegularExpressions.RegexOptions.Compiled);

        private static string StripInvisible(string s) =>
            string.IsNullOrEmpty(s) ? s
            : _invisRe.Replace(s.Normalize(NormalizationForm.FormKD), "");

        /// <summary>
        /// Extracts text from a PDF supplied as a byte array.
        /// Filters noise lines, computes word count, and extracts unique numbers.
        /// </summary>
        public PdfExtractionResult Extract(byte[] pdfBytes, string filename)
        {
            var result = new PdfExtractionResult { Filename = filename };
            using var doc = PdfDocument.Open(pdfBytes);
            int totalWords = 0;

            foreach (var page in doc.GetPages())
            {
                double pageH = page.Height;

                // Group words into lines by Y-coordinate band (round to nearest 5 pt)
                // PDF Y is bottom-up; "distance from page top" = pageH - BoundingBox.Top
                // v4.3.2: StripInvisible applied to each word at source.
                var lineTexts = page.GetWords()
                    .GroupBy(w => (int)(Math.Round((pageH - w.BoundingBox.Top) / 5.0) * 5))
                    .OrderBy(g => g.Key) // ascending = top-of-page first
                    .Select(g => string.Join(" ",
                        g.OrderBy(w => w.BoundingBox.Left)
                         .Select(w => StripInvisible(w.Text ?? "").Trim())
                         .Where(t => t.Length > 0)))
                    .Where(l => !TextNormaliser.IsNoise(l))
                    .ToList();

                totalWords += lineTexts.Sum(l =>
                    l.Split(' ', StringSplitOptions.RemoveEmptyEntries).Length);

                result.Pages.Add(string.Join("\n", lineTexts));
            }

            result.WordCount     = totalWords;
            result.UniqueNumbers = TextNormaliser.ExtractNumbers(result.FullText);
            return result;
        }
    }
}
