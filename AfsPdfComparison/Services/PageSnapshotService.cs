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

        // PDFium (DocLib.Instance) is a process-wide singleton and is NOT thread-safe.
        // Serialize all calls so parallel snapshot requests don't corrupt each other.
        private static readonly SemaphoreSlim _pdfiumLock = new SemaphoreSlim(1, 1);

        /// <summary>
        /// Renders page <paramref name="pageIndex"/> (0-based) from PDF bytes to a
        /// pixel-accurate Base64 PNG, with yellow highlight bands over changed lines.
        /// Mirrors Python v4.3 _pdf_page_to_b64(pdf_path, page_num, dpi=130, highlight_texts=...).
        /// </summary>
        public async Task<string> RenderPageToBase64(byte[] pdfBytes, int pageIndex,
            List<string>? highlightTexts = null)
        {
            if (pdfBytes == null || pdfBytes.Length == 0) return "";

            try
            {
                // ── Step 1: Get PDF page dimensions via PdfPig (gives us PDF unit size) ──
                // PdfPig page dimensions are in PDF points (1 pt = 1/72 inch).
                // We convert to pixels at DPI to get the exact output size for PDFium.
                using var pdfDoc = PdfDocument.Open(pdfBytes);
                var allPages = pdfDoc.GetPages().ToList();
                if (pageIndex >= allPages.Count) return "";

                var pdfPage   = allPages[pageIndex];
                double pageWidth  = pdfPage.Width  > 0 ? pdfPage.Width  : 612;
                double pageHeight = pdfPage.Height > 0 ? pdfPage.Height : 792;

                // Pixel dimensions at target DPI — DPI/72 converts PDF points → pixels
                int imgW = (int)Math.Round(pageWidth  * DPI / 72.0);
                int imgH = (int)Math.Round(pageHeight * DPI / 72.0);

                // ── Step 2: Rasterise PDF page using PDFium (pixel-perfect, like Poppler) ──
                // DocLib.Instance is a process-wide singleton — never dispose it.
                // Acquire the semaphore to serialise concurrent requests.
                await _pdfiumLock.WaitAsync();
                byte[] rawBytes;
                try
                {
                    var lib = DocLib.Instance;
                    using var docReader = lib.GetDocReader(pdfBytes, new PageDimensions(imgW, imgH));

                    if (pageIndex >= docReader.GetPageCount()) return "";

                    using var pageReader = docReader.GetPageReader(pageIndex);

                    // PDFium returns raw BGRA bytes — same pixel data Poppler produces
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

                // ── Step 4: Draw yellow highlight bands over matched lines ───────────────
                if (highlightTexts != null && highlightTexts.Count > 0)
                {
                    // PdfPig words already loaded above — reuse pdfPage
                    if (pageIndex < allPages.Count)
                    {

                        // Scale factors: PDF units → image pixels
                        // Mirrors Python: sx = iw / pw,  sy = ih / ph
                        double sx = imgW / pageWidth;
                        double sy = imgH / pageHeight;

                        // Group words into lines by Y-coordinate band
                        // Python: y_key = round(float(w['top']) / 5) * 5
                        // PdfPig 'top' = pageHeight - BoundingBox.Top  (PDF coords are bottom-up)
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
                            // Reconstruct line text (same as Python)
                            var lineTxt = string.Join(" ",
                                lineWords.OrderBy(w => w.BoundingBox.Left)
                                         .Select(w => w.Text ?? "")).ToLower();

                            // Python matching logic — exact port:
                            // needle = ht.strip().lower()[:30]
                            // if needle and (needle in line_txt or line_txt[:30] in needle)
                            bool matched = (highlightTexts ?? new()).Any(ht =>
                            {
                                var rawN = (ht ?? "").Trim().ToLower();
                                if (rawN.Length < 4) return false;
                                var needle     = rawN.Length    > 30 ? rawN[..30]       : rawN;
                                var lineHead30 = lineTxt.Length > 30 ? lineTxt[..30] : lineTxt;
                                return lineTxt.Contains(needle) ||
                                       (lineHead30.Length >= 6 && needle.Contains(lineHead30));
                            });

                            if (!matched) continue;

                            // Compute bounding box in PDF space
                            // Python: x0=min(w['x0'])*sx-2, x1=max(w['x1'])*sx+2
                            //         y0=min(w['top'])*sy-2, y1=max(w['bottom'])*sy+2
                            // PdfPig: x0=BoundingBox.Left, x1=BoundingBox.Right
                            //         top (from top of page) = pageHeight - BoundingBox.Top
                            //         bottom (from top)      = pageHeight - BoundingBox.Bottom
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
                    }
                }

                // ── Step 4: Encode to PNG Base64 ────────────────────────────────────────
                using var image = SKImage.FromBitmap(bmp);
                using var data  = image.Encode(SKEncodedImageFormat.Png, 90);
                return Convert.ToBase64String(data.ToArray());
            }
            catch (Exception ex)
            {
                // Log and return empty — controller will show placeholder
                Console.WriteLine($"[PageSnapshotService] Error: {ex.Message}");
                return "";
            }
        }
    }
}
