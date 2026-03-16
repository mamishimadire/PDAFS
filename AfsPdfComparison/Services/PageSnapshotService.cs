// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  SERVICE — PageSnapshotService                                               ║
// ║  Renders a PDF page to a Base64 PNG using PdfPig word positions +           ║
// ║  SkiaSharp for drawing.  Works independently of Poppler (no external         ║
// ║  process required) and is used as a FALLBACK when Poppler is unavailable.   ║
// ║                                                                              ║
// ║  Rendering approach:                                                         ║
// ║   1. PdfPig provides word bounding boxes and text.                          ║
// ║   2. SkiaSharp draws each word as black text at its PDF position.           ║
// ║   3. Yellow highlight bands are drawn over lines containing changed text.   ║
// ║                                                                              ║
// ║  BUG FIX for: "Page snapshots have no yellow highlight bands"               ║
// ║   → When Poppler is not found, AfsController.ApiSnapshot now falls back     ║
// ║     to this service so highlights are always visible.                       ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • Python notebook v4.3 — _pdf_page_to_b64(highlight_texts=...)           ║
// ║   • SkiaSharp 2.88 docs — https://github.com/mono/SkiaSharp                ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using UglyToad.PdfPig;
using SkiaSharp;

namespace AfsPdfComparison.Services
{
    public class PageSnapshotService
    {
        private const int   DPI   = 130;
        private const float SCALE = DPI / 72f;

        /// <summary>
        /// Renders page <paramref name="pageIndex"/> (0-based) from the given PDF bytes
        /// to a Base64-encoded PNG.  Yellow highlight bands are drawn over any line
        /// whose text contains a match from <paramref name="highlightTexts"/>.
        /// </summary>
        public string RenderPageToBase64(byte[] pdfBytes, int pageIndex,
            List<string>? highlightTexts = null)
        {
            if (pdfBytes == null || pdfBytes.Length == 0) return "";

            using var pdfDoc = PdfDocument.Open(pdfBytes);
            var pages = pdfDoc.GetPages().ToList();
            if (pageIndex >= pages.Count) return "";
            var page = pages[pageIndex];

            int imgW = (int)(page.Width  * SCALE);
            int imgH = (int)(page.Height * SCALE);

            using var bmp    = new SKBitmap(imgW, imgH);
            using var canvas = new SKCanvas(bmp);
            canvas.Clear(SKColors.White);

            // ── 1. Draw word tokens as black text at their PDF positions ──────
            using var textPaint = new SKPaint
            {
                Color       = SKColors.Black,
                IsAntialias = true,
                TextSize    = 9f,
            };

            foreach (var word in page.GetWords())
            {
                float x = (float)(word.BoundingBox.Left * SCALE);
                // PDF Y is bottom-up; flip to image top-down and offset by font size
                float y = (float)((page.Height - word.BoundingBox.Top) * SCALE) + 9f;
                canvas.DrawText(word.Text ?? "", x, y, textPaint);
            }

            // ── 2. Draw yellow highlight bands over matched lines ─────────────
            if (highlightTexts != null && highlightTexts.Count > 0)
            {
                // Group words by Y-band (same bucketing as text extraction)
                var lineGroups = page.GetWords()
                    .GroupBy(w =>
                        (int)(Math.Round((page.Height - w.BoundingBox.Top) / 5.0) * 5))
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

                foreach (var grp in lineGroups)
                {
                    var words = grp.ToList();
                    var lineText = string.Join(" ",
                        words.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text ?? ""))
                        .ToLower();

                    // Exact Python port: needle = ht.strip().lower()[:30]
                    // if needle and (needle in line_txt or line_txt[:30] in needle)
                    // Guard: only allow reverse-contains when lineHead30 >= 6 chars to prevent
                    // single-digit page numbers (e.g. "2" from "-2-") matching inside the needle.
                    bool matched = highlightTexts.Any(ht =>
                    {
                        var rawN = (ht ?? "").Trim().ToLower();
                        if (rawN.Length < 4) return false;
                        var needle     = rawN.Length    > 30 ? rawN[..30]       : rawN;
                        var lineHead30 = lineText.Length > 30 ? lineText[..30] : lineText;
                        return lineText.Contains(needle) ||
                               (lineHead30.Length >= 6 && needle.Contains(lineHead30));
                    });

                    if (!matched) continue;

                    // Compute bounding box in PDF space
                    float pdfX0 = (float)words.Min(w => w.BoundingBox.Left)   - 2f;
                    float pdfX1 = (float)words.Max(w => w.BoundingBox.Right)  + 2f;
                    float pdfY0 = (float)words.Min(w => w.BoundingBox.Bottom) - 2f;
                    float pdfY1 = (float)words.Max(w => w.BoundingBox.Top)    + 2f;

                    // Transform to image pixel coordinates (flip Y axis)
                    float pixX0 = pdfX0 * SCALE;
                    float pixX1 = pdfX1 * SCALE;
                    float pixY0 = (float)(page.Height - pdfY1) * SCALE;
                    float pixY1 = (float)(page.Height - pdfY0) * SCALE;

                    var rect = new SKRect(pixX0, pixY0, pixX1, pixY1);
                    canvas.DrawRect(rect, fillPaint);
                    canvas.DrawRect(rect, borderPaint);
                }
            }

            // ── 3. Encode to Base64 PNG ───────────────────────────────────────
            using var image = SKImage.FromBitmap(bmp);
            using var data  = image.Encode(SKEncodedImageFormat.Png, 90);
            return Convert.ToBase64String(data.ToArray());
        }
    }
}
