// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  SERVICE — PageSnapshotService  v2.0                                         ║
// ║  Pixel-perfect PDF rasterisation using Docnet.Core (PDFium) + SkiaSharp     ║
// ║  Matches Python notebook v4.3 — _pdf_page_to_b64(highlight_texts=...)       ║
// ║                                                                              ║
// ║  Rendering pipeline (mirrors Python exactly):                                ║
// ║   1. Docnet.Core (PDFium) rasterises the PDF page at 130 DPI               ║
// ║      → produces pixel-accurate BGRA byte array (same as Poppler)            ║
// ║   2. SkiaSharp converts BGRA → SKBitmap                                     ║
// ║   3. PdfPig extracts word bounding boxes for highlight lookup                ║
// ║   4. Lines are grouped by Y-band (same algorithm as Python _pdf_page_to_b64)║
// ║   5. Yellow RGBA(255,220,0,90) fill + RGBA(255,160,0,220) outline drawn     ║
// ║      over matched lines — exact port of Python ImageDraw.rectangle()        ║
// ║   6. Result encoded as PNG Base64                                            ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text;
using System.Text.RegularExpressions;
using AfsPdfComparison.Models;
using Docnet.Core;
using Docnet.Core.Models;
using SkiaSharp;
using UglyToad.PdfPig;

namespace AfsPdfComparison.Services
{
    public class PageSnapshotService
    {
        // Match Python notebook DPI exactly
        private const int DPI = 130;

        // Strip invisible/control chars from PDF word text before highlight matching.
        // Must mirror PageExtractorService.StripInvisible() so the needle (from diff
        // output, already stripped at extraction) matches the lineTxt rebuilt here.
        private static readonly Regex _snapInvisRe =
            new(@"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]",
                RegexOptions.Compiled);

        // Must mirror PageExtractorService.StripInvisible exactly:
        // NFKD-decompose first so Unicode variants of the same glyph map to
        // the same code-points, then strip invisible/control chars.  Without
        // the NFKD step the lineTxt and the needle (produced by extraction)
        // can have different byte representations of the same visual character,
        // causing Contains() comparisons to silently fail.
        private static string SnapStrip(string s) =>
            string.IsNullOrEmpty(s) ? s :
                _snapInvisRe.Replace(s.Normalize(NormalizationForm.FormKD), "");

        // SnapNorm: mirrors LineComparatorService.Normalise() so that the rendered
        // line text and the spec text (which came from extraction → diff pipeline)
        // are compared with identical normalisation.
        // Steps: NFKD → strip invisible → lowercase → collapse whitespace →
        //        strip page-stamp lines → strip currency spacing → decimal spacing.
        private static readonly Regex _snapPageStamp =
            new(@"^[-–—\s]*\(?\s*\d{1,4}\s*\)?[-–—\s]*$", RegexOptions.Compiled);
        private static string SnapNorm(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            string t = SnapStrip(s.ToLowerInvariant());
            t = Regex.Replace(t, @"\s+", " ").Trim();
            t = _snapPageStamp.IsMatch(t) ? "" : t;
            t = Regex.Replace(t, @"\br\s+(?=\d)", "r");
            t = Regex.Replace(t, @"(?<=\d)\s*\.\s*(?=\d)", ".");
            t = Regex.Replace(t, @"(?<=\d)\s*,\s*(?=\d)", ",");
            return t;
        }

        // PDFium (DocLib.Instance) is a process-wide singleton and is NOT thread-safe.
        // Serialize all calls so parallel snapshot requests don't corrupt each other.
        private static readonly SemaphoreSlim _pdfiumLock = new SemaphoreSlim(1, 1);

        /// <summary>
        /// Backward-compatible overload: converts a plain list of line texts to
        /// <see cref="HighlightSpec"/> objects (with empty ChangedWords so entire lines are
        /// highlighted) and calls the main overload. Keeps WorkingPaperExportService and
        /// ComparisonController working without any changes.
        /// </summary>
        public Task<string> RenderPageToBase64(byte[] pdfBytes, int pageIndex,
            List<string>? highlightTexts = null)
        {
            var specs = highlightTexts == null ? null :
                highlightTexts.Select(t => new HighlightSpec { LineText = t ?? "" }).ToList();
            return RenderPageToBase64(pdfBytes, pageIndex, specs);
        }

        /// <summary>
        /// Renders page <paramref name="pageIndex"/> (0-based) from PDF bytes to a
        /// pixel-accurate Base64 PNG with yellow highlight bands over changed lines.
        /// When a <see cref="HighlightSpec"/> has <c>ChangedWords</c> set, only those
        /// individual words are highlighted instead of the full line band.
        /// Mirrors Python v4.3 _pdf_page_to_b64(pdf_path, page_num, dpi=130, highlight_texts=...).
        /// </summary>
        public async Task<string> RenderPageToBase64(byte[] pdfBytes, int pageIndex,
            List<HighlightSpec>? specs = null)
        {
            if (pdfBytes == null || pdfBytes.Length == 0) return "";

            try
            {
                // ── Step 1: Get PDF page dimensions via PdfPig (gives us PDF unit size) ──
                using var pdfDoc = PdfDocument.Open(pdfBytes);
                var allPages = pdfDoc.GetPages().ToList();
                if (pageIndex >= allPages.Count) return "";

                var pdfPage   = allPages[pageIndex];
                double pageWidth  = pdfPage.Width  > 0 ? pdfPage.Width  : 612;
                double pageHeight = pdfPage.Height > 0 ? pdfPage.Height : 792;

                int imgW = (int)Math.Round(pageWidth  * DPI / 72.0);
                int imgH = (int)Math.Round(pageHeight * DPI / 72.0);

                // ── Step 2: Rasterise PDF page using PDFium ──────────────────────────────
                await _pdfiumLock.WaitAsync();
                byte[] rawBytes;
                try
                {
                    var lib = DocLib.Instance;
                    using var docReader = lib.GetDocReader(pdfBytes, new PageDimensions(imgW, imgH));
                    if (pageIndex >= docReader.GetPageCount()) return "";
                    using var pageReader = docReader.GetPageReader(pageIndex);
                    rawBytes = pageReader.GetImage();
                }
                finally
                {
                    _pdfiumLock.Release();
                }

                // ── Step 3: Build SKBitmap from BGRA bytes ──────────────────────────────
                var bmp = new SKBitmap(imgW, imgH, SKColorType.Bgra8888, SKAlphaType.Premul);
                var handle = bmp.GetPixels();
                System.Runtime.InteropServices.Marshal.Copy(rawBytes, 0, handle, rawBytes.Length);

                using var canvas = new SKCanvas(bmp);

                // ── Step 4: Draw highlight bands / word boxes over matched lines ──────────
                if (specs != null && specs.Count > 0)
                {
                    if (pageIndex < allPages.Count)
                    {
                        double sx = imgW / pageWidth;
                        double sy = imgH / pageHeight;

                        var words = pdfPage.GetWords().ToList();
                        var linesByY = words
                            .GroupBy(w =>
                            {
                                double top = pageHeight - w.BoundingBox.Top;
                                return (int)(Math.Round(top / 5.0) * 5);
                            })
                            .ToList();

                        using var fillPaint = new SKPaint
                        {
                            Color = new SKColor(255, 220, 0, 90),
                            Style = SKPaintStyle.Fill,
                        };
                        using var borderPaint = new SKPaint
                        {
                            Color       = new SKColor(255, 160, 0, 220),
                            Style       = SKPaintStyle.Stroke,
                            StrokeWidth = 1f,
                        };

                        foreach (var grp in linesByY)
                        {
                            var lineWords = grp.ToList();
                            var lineTxt = string.Join(" ",
                                lineWords.OrderBy(w => w.BoundingBox.Left)
                                         .Select(w => SnapStrip(w.Text ?? "").Trim())
                                         .Where(w => w.Length > 0)).ToLower();

                            // Normalise the rendered line text the same way the
                            // extraction pipeline normalises it, so spec and line
                            // are compared on equal footing.
                            var lineNorm = SnapNorm(lineTxt);

                            HighlightSpec? matchedSpec = null;
                            foreach (var spec in specs)
                            {
                                var specNorm = SnapNorm((spec.LineText ?? "").Trim());
                                if (specNorm.Length < 4) continue;

                                // ── Strict match strategy ────────────────────
                                // Only highlight this rendered line when the spec
                                // text is demonstrably the SAME line — not just a
                                // word that happens to appear inside a longer sentence.
                                //
                                // Rule 1: exact normalised equality (handles case /
                                //         spacing / invisible-char differences).
                                if (lineNorm == specNorm) { matchedSpec = spec; break; }

                                // Rule 2: prefix match for long lines that wrap
                                // differently across PDFs (spec is the shorter version).
                                // Both must be at least 30 chars to avoid false hits.
                                if (specNorm.Length >= 30 && lineNorm.Length >= 30)
                                {
                                    string shorter = specNorm.Length <= lineNorm.Length ? specNorm : lineNorm;
                                    string longer  = specNorm.Length <= lineNorm.Length ? lineNorm : specNorm;
                                    // Length ratio must be ≥ 0.60 — spec must cover at
                                    // least 60 % of the line to be a genuine prefix match,
                                    // not just a short fragment of a long paragraph.
                                    double ratio = (double)shorter.Length / longer.Length;
                                    if (ratio >= 0.60 && longer.StartsWith(shorter, StringComparison.Ordinal))
                                    { matchedSpec = spec; break; }
                                }

                                // Rule 3: high ordered-token overlap for long lines
                                // (handles minor wording differences in the same sentence).
                                // Only fires when BOTH spec and line have ≥ 6 words,
                                // preventing short headings / single words from matching.
                                var lineTokens = lineNorm.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                                var specTokens = specNorm.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                                if (lineTokens.Length >= 6 && specTokens.Length >= 6)
                                {
                                    var lineSet = lineTokens.ToHashSet(StringComparer.Ordinal);
                                    var specSet = specTokens.ToHashSet(StringComparer.Ordinal);
                                    int common  = lineSet.Intersect(specSet).Count();
                                    int maxWc   = Math.Max(lineSet.Count, specSet.Count);
                                    int minWc   = Math.Min(lineSet.Count, specSet.Count);
                                    double wordOverlap = (double)common / maxWc;
                                    double wordRatio   = (double)minWc  / maxWc;
                                    // Both overlap AND length-ratio must clear high bars
                                    // to ensure this is the SAME sentence, not a paragraph
                                    // that happens to share many common English words.
                                    if (wordOverlap >= 0.88 && wordRatio >= 0.65)
                                    { matchedSpec = spec; break; }
                                }
                            }

                            if (matchedSpec == null) continue;

                            if (matchedSpec.ChangedWords.Count == 0)
                            {
                                // Whole-line band highlight (added / removed lines)
                                double pdfX0   = lineWords.Min(w => w.BoundingBox.Left);
                                double pdfX1   = lineWords.Max(w => w.BoundingBox.Right);
                                double pdfTop  = lineWords.Min(w => pageHeight - w.BoundingBox.Top);
                                double pdfBot  = lineWords.Max(w => pageHeight - w.BoundingBox.Bottom);

                                float pixX0 = (float)(pdfX0 * sx) - 2f;
                                float pixX1 = (float)(pdfX1 * sx) + 2f;
                                float pixY0 = (float)(pdfTop * sy) - 2f;
                                float pixY1 = (float)(pdfBot * sy) + 2f;

                                var rect = new SKRect(pixX0, pixY0, pixX1, pixY1);
                                canvas.DrawRect(rect, fillPaint);
                                canvas.DrawRect(rect, borderPaint);
                            }
                            else
                            {
                                // Per-word highlight: only the specific changed tokens
                                var changedSet = matchedSpec.ChangedWords
                                    .Select(w => SnapStrip(w ?? "").Trim().ToLower())
                                    .Where(w => w.Length >= 1)
                                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                                foreach (var word in lineWords)
                                {
                                    var wt = SnapStrip(word.Text ?? "").Trim().ToLower();
                                    if (wt.Length < 1 || !changedSet.Contains(wt)) continue;

                                    double wx0  = word.BoundingBox.Left;
                                    double wx1  = word.BoundingBox.Right;
                                    double wTop = pageHeight - word.BoundingBox.Top;
                                    double wBot = pageHeight - word.BoundingBox.Bottom;

                                    var wr = new SKRect(
                                        (float)(wx0  * sx) - 2f,
                                        (float)(wTop * sy) - 2f,
                                        (float)(wx1  * sx) + 2f,
                                        (float)(wBot * sy) + 2f);
                                    canvas.DrawRect(wr, fillPaint);
                                    canvas.DrawRect(wr, borderPaint);
                                }
                            }
                        }
                    }
                }

                // ── Step 5: Encode to PNG Base64 ────────────────────────────────────────
                using var image = SKImage.FromBitmap(bmp);
                using var data  = image.Encode(SKEncodedImageFormat.Png, 90);
                return Convert.ToBase64String(data.ToArray());
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PageSnapshotService] Error: {ex.Message}");
                return "";
            }
        }
    }
}
