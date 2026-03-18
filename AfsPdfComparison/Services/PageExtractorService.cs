// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3                                                               ║
// ║                                                                              ║
// ║  SERVICE — PageExtractorService                                              ║
// ║  Reads each uploaded PDF page by page, producing a PdfReport.               ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • PdfPig library (UglyToad.PdfPig) — MIT-licensed C# PDF parser           ║
// ║     https://github.com/UglyToad/PdfPig                                       ║
// ║   • Python equivalent: PageExtractor._try_digital() in the notebook         ║
// ║   • Regex year detection: matches 4-digit years 1900–2099                   ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text;
using System.Text.RegularExpressions;
using AfsPdfComparison.Models;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace AfsPdfComparison.Services;

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 1 · PAGE EXTRACTOR SERVICE
//
// Responsibility: Open a PDF file, extract the text of every page, and build
// a PdfReport object.  Uses PdfPig for digital PDFs (selectable text).
// Scanned / image-only PDFs are detected by checking whether the extracted
// character count is suspiciously low; a warning flag is set in DocType.
//
// Key design decisions:
//  • Pages with fewer than 100 characters total are flagged as potentially scanned.
//  • Financial years (1900–2099) are extracted with a simple regex so the
//    document's primary year can be shown in the summary table.
//  • The temporary PDF path is written to PdfReport.PdfPath so the snapshot
//    renderer can re-open the same file for page image generation.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Extracts text from a PDF using PdfPig and produces a <see cref="PdfReport"/>.
/// One instance is registered as a singleton in DI.
/// </summary>
public class PageExtractorService
{
    // Regex: matches 4-digit years 1900–2099 anywhere in the text.
    // Reference: Python re.findall(r'(?:19|20)\d{2}', full_text)
    private static readonly Regex _yearRegex = new(@"(?:19|20)\d{2}", RegexOptions.Compiled);

    /// <summary>
    /// Extracts all pages from a PDF file at <paramref name="pdfPath"/>.
    /// </summary>
    /// <param name="pdfPath">Absolute path to the temporary PDF on the server.</param>
    /// <param name="filename">Original filename shown in the UI and exported papers.</param>
    /// <returns>A populated <see cref="PdfReport"/> for this document.</returns>
    public PdfReport Extract(string pdfPath, string filename)
    {
        var pages = new List<string>();

        // ── 1.1  Digital extraction via PdfPig ────────────────────────────
        // PdfPig.Open() reads the cross-reference table and page tree.
        // Page.GetWords() returns positionally-sorted word tokens which we
        // join into a single line per page (the library does not preserve
        // the original newline structure, so we reconstruct it by grouping
        // words whose Y-coordinates are within a tolerance band).
        using (var document = PdfDocument.Open(pdfPath))
        {
            foreach (var page in document.GetPages())
            {
                pages.Add(ExtractPageText(page));
            }
        }

        // ── 1.2  Scanned-document detection ───────────────────────────────
        // If the total extracted character count is very low the PDF likely
        // consists of scanned images.  We cannot do OCR in this C# version
        // (that would require Tesseract.NET integration), so we flag the
        // document and let the user know.
        int totalChars = pages.Sum(p => p.Length);
        string docType = totalChars < 100 ? "scanned (no OCR)" : "digital";

        // ── 1.3  Full-document text and year detection ─────────────────────
        string fullText = string.Join("\n", pages);
        var yearsFound  = _yearRegex.Matches(fullText)
                                    .Select(m => int.Parse(m.Value))
                                    .Distinct()
                                    .OrderBy(y => y)
                                    .ToList();

        return new PdfReport
        {
            Filename    = filename,
            DocType     = docType,
            Pages       = pages,
            FullText    = fullText,
            PageCount   = pages.Count,
            // WordCount is a computed property on PdfReport (filters noise lines automatically)
            Years       = yearsFound,
            PrimaryYear = yearsFound.Count > 0 ? yearsFound.Max() : (int?)null,
            PdfPath     = pdfPath,
        };
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 1-INTERNAL · LINE RECONSTRUCTION FROM PDF WORDS
    //
    // PdfPig returns individual Word objects with bounding-box coordinates.
    // Lines are reconstructed by grouping words whose top-Y values round to
    // the same integer bin (bin size = 5 user units ≈ 1.8 pt at 72dpi).
    // This mirrors the Python pdfplumber lines_by_y grouping used in the
    // snapshot highlight routine.
    // ─────────────────────────────────────────────────────────────────────────
    // ── v4.3.2 invisible-char stripping ──────────────────────────────────────
    // PdfPig occasionally embeds Unicode control/format characters (zero-width
    // spaces, object-replacement chars, BOM, …) inside word Text from hyperlinked
    // or annotated glyphs.  Strip them here — at the source — so that the raw
    // PdfReport.Pages text is already clean before any comparison logic runs.
    // \uFFFD (REPLACEMENT CHARACTER, category So, NOT \p{C}) is added explicitly:
    // PdfPig emits it when it cannot decode a glyph from a non-standard encoding.
    // \uFFF0-\uFFFD covers the full Unicode Specials block (includes \uFFFC).
    private static readonly System.Text.RegularExpressions.Regex _invisRe =
        new(@"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]",
            System.Text.RegularExpressions.RegexOptions.Compiled);

    private static string StripInvisible(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        return _invisRe.Replace(s.Normalize(NormalizationForm.FormKD), "");
    }

    private static string ExtractPageText(Page page)
    {
        try
        {
            // Group words by their rounded Y position (line key).
            // PdfPig: origin bottom-left, BoundingBox.Top = distance from page bottom to top of glyph.
            // Convert to "distance from page TOP" to match Python pdfplumber w['top']:
            //   pdfplumber_top = pageHeight - BoundingBox.Top
            // Then round to nearest 5 pt bucket — identical to Python round(w['top']/5)*5.
            // Using Top (not Bottom) avoids false line-splits caused by descender glyphs (g,p,y,j).
            double pageH = page.Height;
            var lineGroups = page.GetWords()
                .GroupBy(w => (int)(Math.Round((pageH - w.BoundingBox.Top) / 5.0) * 5))
                .OrderBy(g => g.Key);   // ascending = top-of-page first in distance-from-top coords

            // ── Step 1: strip invisible chars from each word ───────────────
            var rawLines = new List<string>();
            foreach (var lineGroup in lineGroups)
            {
                var words = lineGroup
                    .OrderBy(w => w.BoundingBox.Left)
                    .Select(w => StripInvisible(w.Text ?? "").Trim())
                    .Where(w => w.Length > 0);
                string line = string.Join(" ", words).Trim();
                if (line.Length > 0)
                    rawLines.Add(line);
            }

            // ── Step 2: paragraph reconstruction ──────────────────────────
            // Merge continuation lines to normalise column-width differences.
            // Rules (applied at SOURCE so both PDFs get identical reconstruction):
            //   • Only merge when the current line starts with a LOWERCASE letter
            //     (genuine mid-sentence wrap).  Capital-letter lines are always
            //     new sentences/headings and must NOT be merged.
            //   • Previous line must be at least 15 chars (not just a label).
            //   • Previous line must not end with a sentence-terminal char (.!?;:).
            //   • Consecutive identical lines (hyperlink duplication) are dropped.
            var mergedLines = new List<string>(rawLines.Count);
            foreach (string line in rawLines)
            {
                if (mergedLines.Count > 0 &&
                    string.Equals(mergedLines[^1], line, StringComparison.OrdinalIgnoreCase))
                    continue;   // deduplicate hyperlink artefacts

                bool merged = false;
                if (mergedLines.Count > 0)
                {
                    string prev   = mergedLines[^1];
                    char   last   = prev.Length > 0 ? prev[^1] : '\0';
                    bool prevEnds = last is '.' or '!' or '?' or ';' or ':';
                    bool currCont = line.Length > 0 && char.IsLower(line[0]);

                    if (!prevEnds && currCont && prev.Length >= 15)
                    {
                        mergedLines[^1] = prev + " " + line;
                        merged = true;
                    }
                }
                if (!merged)
                    mergedLines.Add(line);
            }

            var sb = new StringBuilder();
            foreach (string line in mergedLines)
                // Use "\n" not AppendLine — AppendLine emits "\r\n" on Windows which
                // leaves stray "\r" characters after SplitLines splits on "\n".
                sb.Append(line + "\n");

            return sb.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }
}
