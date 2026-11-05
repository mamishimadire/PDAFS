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
using UglyToad.PdfPig.Content;

namespace AfsPdfComparison.Services
{
    public class PageSnapshotService
    {
        private sealed class RenderLine
        {
            public List<Word> Words { get; } = new();
            public string Text { get; set; } = "";
            public string NormalizedText => SnapNorm(Text);
        }

        // Render at a higher DPI so the in-app snapshot stays legible even when
        // displayed beside another page in the browser.
        private const int DPI = 220;

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

        private static bool HasNumericChangedWords(HighlightSpec spec) =>
            spec.ChangedWords.Any(w => Regex.IsMatch(w ?? "", @"\d"));

        private static bool ShouldMatchSnapshotSpec(string lineNorm, HighlightSpec spec)
        {
            if (!spec.HighlightEligible || spec.Confidence < 0.62)
                return false;

            var specNorm = SnapNorm((spec.LineText ?? "").Trim());
            if (specNorm.Length < 4)
                return false;

            var lineTokenSet = new HashSet<string>(
                lineNorm.Split(' ', StringSplitOptions.RemoveEmptyEntries),
                StringComparer.OrdinalIgnoreCase);
            var changedNorms = spec.ChangedWords
                .SelectMany(w => SnapNorm(w ?? "")
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var anchorNorms = spec.AnchorWords
                .SelectMany(w => SnapNorm(w ?? "")
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (lineNorm == specNorm)
            {
                if (spec.WholeLineMode)
                    return true;

                if (changedNorms.Count == 0 ||
                    changedNorms.All(changed => lineTokenSet.Contains(changed)))
                    return false;
            }

            bool changedTokenPresent = changedNorms.Any(lineTokenSet.Contains);
            if (changedNorms.Count > 0 && !changedTokenPresent)
                return false;

            int anchorHits = anchorNorms.Count(lineTokenSet.Contains);
            int requiredAnchorHits = Math.Min(2, anchorNorms.Count);
            bool hasAnchorContext = requiredAnchorHits > 0 && anchorHits >= requiredAnchorHits;

            if (specNorm.Length >= 30 && lineNorm.Length >= 30)
            {
                string shorter = specNorm.Length <= lineNorm.Length ? specNorm : lineNorm;
                string longer = specNorm.Length <= lineNorm.Length ? lineNorm : specNorm;
                double ratio = (double)shorter.Length / longer.Length;
                bool orderedPrefix = ratio >= 0.60 &&
                    longer.StartsWith(shorter, StringComparison.Ordinal);
                bool hasChangeContext = hasAnchorContext || spec.WholeLineMode;
                if (orderedPrefix &&
                    hasChangeContext &&
                    (HasNumericChangedWords(spec) || changedTokenPresent ||
                     (spec.ChangedWords.Count == 0 && spec.WholeLineMode)))
                {
                    return true;
                }
            }

            return lineNorm == specNorm && (spec.ChangedWords.Count > 0 || spec.WholeLineMode);
        }

        private static HashSet<string> BuildChangedTokenSet(IEnumerable<string> changedWords)
        {
            var changedSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var cw in changedWords)
            {
                if (string.IsNullOrWhiteSpace(cw))
                    continue;

                string stripped = SnapStrip(cw).Trim().ToLowerInvariant();
                if (stripped.Length == 0)
                    continue;

                changedSet.Add(stripped);

                string clean = Regex.Replace(stripped, @"[^\p{L}\p{N}]", "");
                if (clean.Length > 0)
                    changedSet.Add(clean);

                string numClean = Regex.Replace(stripped, @"\s+", "");
                if (numClean.Length > 0)
                    changedSet.Add(numClean);
            }

            return changedSet;
        }

        private static bool SnapshotWordMatchesChangedSet(string wordText, HashSet<string> changedSet)
        {
            string wRaw = SnapStrip(wordText ?? "").Trim().ToLowerInvariant();
            string wClean = Regex.Replace(wRaw, @"[^\p{L}\p{N}]", "");
            string wNum = Regex.Replace(wRaw, @"\s+", "");

            return (wRaw.Length >= 1 && changedSet.Contains(wRaw)) ||
                   (wClean.Length >= 1 && changedSet.Contains(wClean)) ||
                   (wNum.Length >= 1 && changedSet.Contains(wNum));
        }

        private static string BuildLineText(IEnumerable<Word> words) =>
            string.Join(" ",
                words.Select(w => SnapStrip(w.Text ?? "").Trim())
                     .Where(w => w.Length > 0));

        private static int WordCount(string text) =>
            (text ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;

        private static bool ContainsWholeLine(string haystack, string needle) =>
            !string.IsNullOrWhiteSpace(needle) &&
            (" " + haystack + " ").Contains(" " + needle + " ", StringComparison.Ordinal);

        private static List<RenderLine> BuildRenderableLines(UglyToad.PdfPig.Content.Page pdfPage, double pageHeight)
        {
            var raw = pdfPage.GetWords()
                .GroupBy(w =>
                {
                    double top = pageHeight - w.BoundingBox.Top;
                    return (int)(Math.Round(top / 5.0) * 5);
                })
                .OrderBy(g => g.Key)
                .Select(g =>
                {
                    var words = g.OrderBy(w => w.BoundingBox.Left).ToList();
                    var line = new RenderLine { Text = BuildLineText(words) };
                    line.Words.AddRange(words);
                    return line;
                })
                .Where(l => !string.IsNullOrWhiteSpace(l.Text))
                .ToList();

            if (raw.Count == 0) return raw;

            var merged = new List<RenderLine> { raw[0] };
            for (int i = 1; i < raw.Count; i++)
            {
                var prev = merged[^1];
                var curr = raw[i];

                if (string.Equals(prev.NormalizedText, curr.NormalizedText, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (WordCount(curr.Text) <= 3 &&
                    ContainsWholeLine(prev.NormalizedText, curr.NormalizedText))
                    continue;

                char last = prev.Text[^1];
                bool prevEnds = last is '.' or '!' or '?' or ';' or ':';
                bool currCont = curr.Text.Length > 0 && char.IsLower(curr.Text[0]);

                if (!prevEnds && currCont && prev.Text.Length >= 15)
                {
                    prev.Text += " " + curr.Text;
                    prev.Words.AddRange(curr.Words);
                }
                else
                {
                    merged.Add(curr);
                }
            }

            var deduped = new List<RenderLine>(merged.Count);
            for (int i = 0; i < merged.Count; i++)
            {
                var line = merged[i];
                if (WordCount(line.Text) <= 3)
                {
                    string needle = line.NormalizedText;
                    bool orphan = false;
                    for (int j = 0; j < merged.Count; j++)
                    {
                        if (i == j) continue;
                        if (ContainsWholeLine(merged[j].NormalizedText, needle))
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
                        var renderLines = BuildRenderableLines(pdfPage, pageHeight);

                        // Normal (yellow) — deterministic diff, no AI flag
                        using var fillPaintNormal    = new SKPaint { Color = new SKColor(255, 220, 0, 90),  Style = SKPaintStyle.Fill };
                        using var lineFillNormal     = new SKPaint { Color = new SKColor(255, 220, 0, 42),  Style = SKPaintStyle.Fill };
                        using var borderPaintNormal  = new SKPaint { Color = new SKColor(255, 160, 0, 220), Style = SKPaintStyle.Stroke, StrokeWidth = 1f };
                        // Medium (orange) — AI page-focus flag
                        using var fillPaintMedium    = new SKPaint { Color = new SKColor(255, 140, 0, 105), Style = SKPaintStyle.Fill };
                        using var lineFillMedium     = new SKPaint { Color = new SKColor(255, 140, 0, 48),  Style = SKPaintStyle.Fill };
                        using var borderPaintMedium  = new SKPaint { Color = new SKColor(230, 100, 0, 230), Style = SKPaintStyle.Stroke, StrokeWidth = 1.5f };
                        // High (red) — AI material finding
                        using var fillPaintHigh      = new SKPaint { Color = new SKColor(220, 30, 60, 115),  Style = SKPaintStyle.Fill };
                        using var lineFillHigh       = new SKPaint { Color = new SKColor(220, 30, 60, 50),   Style = SKPaintStyle.Fill };
                        using var borderPaintHigh    = new SKPaint { Color = new SKColor(200, 16, 46, 240),  Style = SKPaintStyle.Stroke, StrokeWidth = 2f };

                        var verifiedSpecs = specs.Where(spec => spec.HighlightEligible && spec.Confidence >= 0.62).ToList();
                        if (verifiedSpecs.Count > 0 && verifiedSpecs.All(spec => spec.SourceCoordinates.Count > 0))
                        {
                            foreach (var spec in verifiedSpecs)
                            {
                                var (fillPaint, _, borderPaint) = spec.Severity switch
                                {
                                    AiHighlightSeverity.High => (fillPaintHigh, lineFillHigh, borderPaintHigh),
                                    AiHighlightSeverity.Medium => (fillPaintMedium, lineFillMedium, borderPaintMedium),
                                    _ => (fillPaintNormal, lineFillNormal, borderPaintNormal),
                                };
                                foreach (var coordinates in spec.SourceCoordinates.Where(rect => rect.PageNumber == pageIndex + 1))
                                {
                                    var rectangle = new SKRect(
                                        (float)(coordinates.X * sx) - 1f,
                                        (float)(coordinates.Y * sy) - 1f,
                                        (float)((coordinates.X + coordinates.Width) * sx) + 1f,
                                        (float)((coordinates.Y + coordinates.Height) * sy) + 1f);
                                    canvas.DrawRect(rectangle, fillPaint);
                                    canvas.DrawRect(rectangle, borderPaint);
                                }
                            }

                            using var verifiedImage = SKImage.FromBitmap(bmp);
                            using var verifiedData = verifiedImage.Encode(SKEncodedImageFormat.Png, 90);
                            return Convert.ToBase64String(verifiedData.ToArray());
                        }

                        foreach (var renderLine in renderLines)
                        {
                            var lineWords = renderLine.Words;
                            var lineNorm = renderLine.NormalizedText;

                            HighlightSpec? matchedSpec = null;
                            foreach (var spec in specs.Where(spec => spec.HighlightEligible && spec.Confidence >= 0.62))
                            {
                                if (ShouldMatchSnapshotSpec(lineNorm, spec))
                                {
                                    matchedSpec = spec;
                                    break;
                                }
                            }

                            if (matchedSpec == null) continue;

                            // Choose paint set based on AI severity
                            var (fillPaint, lineFillPaint, borderPaint) = matchedSpec.Severity switch
                            {
                                AiHighlightSeverity.High   => (fillPaintHigh,   lineFillHigh,   borderPaintHigh),
                                AiHighlightSeverity.Medium => (fillPaintMedium, lineFillMedium, borderPaintMedium),
                                _                          => (fillPaintNormal, lineFillNormal, borderPaintNormal),
                            };

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
                                if (!matchedSpec.ExactRegion)
                                {
                                    double pdfX0   = lineWords.Min(w => w.BoundingBox.Left);
                                    double pdfX1   = lineWords.Max(w => w.BoundingBox.Right);
                                    double pdfTop  = lineWords.Min(w => pageHeight - w.BoundingBox.Top);
                                    double pdfBot  = lineWords.Max(w => pageHeight - w.BoundingBox.Bottom);

                                    float pixX0 = (float)(pdfX0 * sx) - 2f;
                                    float pixX1 = (float)(pdfX1 * sx) + 2f;
                                    float pixY0 = (float)(pdfTop * sy) - 2f;
                                    float pixY1 = (float)(pdfBot * sy) + 2f;

                                    var lineRect = new SKRect(pixX0, pixY0, pixX1, pixY1);
                                    canvas.DrawRect(lineRect, lineFillPaint);
                                }

                                // Per-word highlight: emphasize the specific changed tokens
                                var changedSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                foreach (var cw in matchedSpec.ChangedWords)
                                {
                                    if (string.IsNullOrWhiteSpace(cw)) continue;
                                    string stripped = SnapStrip(cw).Trim().ToLower();
                                    if (stripped.Length == 0) continue;

                                    changedSet.Add(stripped);

                                    string clean = System.Text.RegularExpressions.Regex
                                        .Replace(stripped, @"[^\p{L}\p{N}]", "");
                                    if (clean.Length > 0) changedSet.Add(clean);

                                    string numClean = System.Text.RegularExpressions.Regex
                                        .Replace(stripped, @"\s+", "");
                                    if (numClean.Length > 0) changedSet.Add(numClean);
                                }

                                bool anyWordHit = false;
                                foreach (var word in lineWords)
                                {
                                    string wRaw = SnapStrip(word.Text ?? "").Trim().ToLower();
                                    string wClean = System.Text.RegularExpressions.Regex
                                        .Replace(wRaw, @"[^\p{L}\p{N}]", "");
                                    string wNum = System.Text.RegularExpressions.Regex
                                        .Replace(wRaw, @"\s+", "");

                                    bool hit = (wRaw.Length >= 1 && changedSet.Contains(wRaw)) ||
                                               (wClean.Length >= 1 && changedSet.Contains(wClean)) ||
                                               (wNum.Length >= 1 && changedSet.Contains(wNum));

                                    if (!hit) continue;
                                    anyWordHit = true;

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

                                if (!anyWordHit && !matchedSpec.ExactRegion)
                                {
                                    double pdfX0 = lineWords.Min(w => w.BoundingBox.Left);
                                    double pdfX1 = lineWords.Max(w => w.BoundingBox.Right);
                                    double pdfTop = lineWords.Min(w => pageHeight - w.BoundingBox.Top);
                                    double pdfBot = lineWords.Max(w => pageHeight - w.BoundingBox.Bottom);
                                    var lineRect = new SKRect(
                                        (float)(pdfX0 * sx) - 2f,
                                        (float)(pdfTop * sy) - 2f,
                                        (float)(pdfX1 * sx) + 2f,
                                        (float)(pdfBot * sy) + 2f);
                                    canvas.DrawRect(lineRect, borderPaint);
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
