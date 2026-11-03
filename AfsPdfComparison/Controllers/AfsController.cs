// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3                                                               ║
// ║                                                                              ║
// ║  CONTROLLER — AfsController                                                  ║
// ║  Main MVC controller: handles all wizard steps, AJAX endpoints, and export. ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • ASP.NET Core MVC routing — Microsoft Docs                               ║
// ║   • IFormFile — multipart/form-data upload handling                         ║
// ║   • Python equivalent: widget event handlers (on_eng, on_add, on_extract,  ║
// ║     on_compare, on_save) in the notebook                                    ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using AfsPdfComparison.Models;
using AfsPdfComparison.Services;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace AfsPdfComparison.Controllers;

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 7 · AFS CONTROLLER
//
// Route map (mirrors the 5 Accordion steps in the Python notebook):
//
//   GET  /            → Step 1 — Engagement Details form
//   POST /engagement  → Step 1 — Save engagement details
//   GET  /upload      → Step 2 — Upload & queue PDFs
//   POST /upload      → Step 2 — Add PDF to queue
//   POST /clear       → Step 2 — Clear all queued PDFs
//   POST /extract     → Step 3 — Extract text from all queued PDFs
//   GET  /compare     → Step 4 — Show comparison UI
//   POST /compare     → Step 4 — Run comparison
//   POST /comment     → Step 4 — Save a single auditor comment (AJAX)
//   GET  /export      → Step 5 — Export options form
//   POST /export      → Step 5 — Generate and save working papers
// ─────────────────────────────────────────────────────────────────────────────

[Route("")]
public class AfsController : Controller
{
    private readonly PageExtractorService              _extractor;
    private readonly ComparisonService                  _comparer;
    private readonly ExportService                      _exporter;
    private readonly SessionStateService                _session;
    private readonly ComparisonRuntimeStore             _comparisonStore;
    private readonly IWebHostEnvironment                _env;
    private readonly PageSnapshotService                _snapshots;
    private readonly WorkingPaperExportService          _wpExporter;
    private readonly MultiAiOrchestrator                 _multiAi;
    private readonly AiContentReaderService              _aiContentReader;
    private readonly AiBackendStatusService              _aiBackendStatus;

    public AfsController(
        PageExtractorService              extractor,
        ComparisonService                 comparer,
        ExportService                     exporter,
        SessionStateService               session,
        ComparisonRuntimeStore            comparisonStore,
        IWebHostEnvironment               env,
        PageSnapshotService               snapshots,
        WorkingPaperExportService         wpExporter,
        MultiAiOrchestrator               multiAi,
        AiContentReaderService            aiContentReader,
        AiBackendStatusService            aiBackendStatus)
    {
        _extractor       = extractor;
        _comparer        = comparer;
        _exporter        = exporter;
        _session         = session;
        _comparisonStore = comparisonStore;
        _env             = env;
        _snapshots       = snapshots;
        _wpExporter      = wpExporter;
        _multiAi         = multiAi;
        _aiContentReader = aiContentReader;
        _aiBackendStatus = aiBackendStatus;
    }

    private string SessionId => HttpContext.Session.Id;

    private ComparisonResult? GetActiveComparison()
    {
        var live = _comparisonStore.Get(SessionId);
        if (live != null)
            return live;

        return _session.GetComparison();
    }

    private static ComparisonResult CreateSessionSafeComparison(ComparisonResult result)
    {
        return new ComparisonResult
        {
            Report1 = result.Report1,
            Report2 = result.Report2,
            FullDiff = new(),
            Counts = new Dictionary<string, int>(result.Counts),
            NumCmp = new NumberComparisonResult
            {
                InBoth = result.NumCmp.InBoth.ToList(),
                OnlyInAfs1 = result.NumCmp.OnlyInAfs1.ToList(),
                OnlyInAfs2 = result.NumCmp.OnlyInAfs2.ToList(),
                CountAfs1 = result.NumCmp.CountAfs1,
                CountAfs2 = result.NumCmp.CountAfs2,
                SimilarityPct = result.NumCmp.SimilarityPct,
            },
            PageDiffs = result.PageDiffs.Select(pd => new PageDiffResult
            {
                PairIndex = pd.PairIndex,
                PageAfs1 = pd.PageAfs1,
                PageAfs2 = pd.PageAfs2,
                AlignSim = pd.AlignSim,
                PctSame = pd.PctSame,
                Same = pd.Same,
                Changed = pd.Changed,
                Added = pd.Added,
                Removed = pd.Removed,
                TotalLines = pd.TotalLines,
                SecondaryMatch = pd.SecondaryMatch,
                Unmatched = pd.Unmatched,
                Diffs = new(),
            }).ToList(),
            Alignment = result.Alignment.Select(a => new AlignmentEntry
            {
                I1 = a.I1,
                I2 = a.I2,
                Sim = a.Sim,
            }).ToList(),
            ComparedAt = result.ComparedAt,
        };
    }

    private void StoreComparison(ComparisonResult result)
    {
        _comparisonStore.Set(SessionId, result);
        _session.SetComparison(CreateSessionSafeComparison(result));
    }

    private void DeleteQueuedPdfFiles(IEnumerable<PdfReport> reports)
    {
        foreach (var report in reports)
        {
            try
            {
                if (System.IO.File.Exists(report.PdfPath))
                    System.IO.File.Delete(report.PdfPath);
            }
            catch
            {
                // Best effort only. A locked file should not block session reset.
            }
        }
    }

    private void ResetComparisonSession(bool deleteQueuedFiles)
    {
        if (deleteQueuedFiles)
            DeleteQueuedPdfFiles(_session.GetReports());

        _comparisonStore.Remove(SessionId);
        _session.SetReports(new());
        _session.SetComparison(null!);
        _session.SetComments(new());
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 1 · ENGAGEMENT DETAILS
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// GET /  — Display the engagement details form.
    /// Pre-populates the form from any previously saved session data.
    /// </summary>
    [HttpGet("")]
    public IActionResult Index()
    {
        var vm = new EngagementViewModel
        {
            Details = _session.GetEngagement(),
        };
        return View(vm);
    }

    /// <summary>
    /// POST /engagement — Save engagement details to session.
    /// Redirects back to the form on success with a success flag.
    /// </summary>
    [HttpPost("engagement")]
    [ValidateAntiForgeryToken]
    public IActionResult SaveEngagement(EngagementDetails details)
    {
        _session.SetEngagement(details);
        TempData["EngagementSaved"] = true;
        return RedirectToAction(nameof(Index));
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 2 · UPLOAD PDFs
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// GET /upload — Show the upload queue with any previously queued files.
    /// </summary>
    [HttpGet("upload")]
    public IActionResult Upload()
    {
        var reports = _session.GetReports();
        var vm = new UploadViewModel
        {
            QueuedFiles = reports.Select(r => r.Filename).ToList(),
            Reports     = reports,
        };
        return View(vm);
    }

    /// <summary>
    /// POST /upload — Accept one or more PDF files and queue them in session.
    /// Files are written to wwwroot/uploads/{guid}.pdf so they survive the
    /// session (the temp path is stored in PdfReport.PdfPath).
    /// </summary>
    [HttpPost("upload")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Upload(List<IFormFile> files)
    {
        if (files == null || !files.Any())
        {
            TempData["UploadError"] = "No files selected.";
            return RedirectToAction(nameof(Upload));
        }

        var uploadsDir = Path.Combine(_env.WebRootPath, "uploads");
        Directory.CreateDirectory(uploadsDir);

        var acceptedFiles = files
            .Where(file =>
                file.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) &&
                file.ContentType.Contains("pdf", StringComparison.OrdinalIgnoreCase))
            .Take(2)
            .ToList();

        if (acceptedFiles.Count == 0)
        {
            TempData["UploadError"] = "Only PDF files are accepted.";
            return RedirectToAction(nameof(Upload));
        }

        ResetComparisonSession(deleteQueuedFiles: true);

        var reports = new List<PdfReport>();

        foreach (var file in acceptedFiles)
        {
            // Write to a GUID-named server path (no user-controlled path component)
            string safeName = Guid.NewGuid().ToString("N") + ".pdf";
            string serverPath = Path.Combine(uploadsDir, safeName);
            await using var fs = new FileStream(serverPath, FileMode.Create);
            await file.CopyToAsync(fs);

            // Create a placeholder report (not yet extracted) so the queue shows it
            reports.Add(new PdfReport
            {
                Filename = file.FileName,
                PdfPath  = serverPath,
                Label    = "AFS " + (reports.Count + 1),
            });
        }

        _session.SetReports(reports);

        int ignoredCount = files.Count - acceptedFiles.Count;
        string ignoredMsg = ignoredCount > 0
            ? $" Ignored {ignoredCount} extra or non-PDF file(s)."
            : "";
        TempData["UploadSuccess"] =
            $"{reports.Count} PDF(s) queued as the active comparison set.{ignoredMsg}";
        return RedirectToAction(nameof(Upload));
    }

    /// <summary>
    /// POST /clear — Remove all queued files and reset the session.
    /// Deletes server-side temp files to free disk space.
    /// </summary>
    [HttpPost("clear")]
    [ValidateAntiForgeryToken]
    public IActionResult Clear()
    {
        ResetComparisonSession(deleteQueuedFiles: true);
        TempData["UploadSuccess"] = "Queue cleared.";
        return RedirectToAction(nameof(Upload));
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 3 · EXTRACT
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// POST /extract — Runs page extraction on all queued PDFs.
    /// Updates PdfReport objects with extracted text, page count, years etc.
    /// Redirects to upload page with extraction summary in TempData.
    /// </summary>
    [HttpPost("extract")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Extract()
    {
        var reports = _session.GetReports();
        if (!reports.Any())
        {
            TempData["ExtractError"] = "No PDFs queued.";
            return RedirectToAction(nameof(Upload));
        }

        var messages = new List<string>();
        for (int i = 0; i < reports.Count; i++)
        {
            var r = reports[i];
            try
            {
                var extracted = await _extractor.ExtractAsync(r.PdfPath, r.Filename, HttpContext.RequestAborted);
                extracted.Label   = "AFS " + (i + 1);
                extracted.PdfPath = r.PdfPath;
                reports[i]        = extracted;
                messages.Add($"AFS {i+1} | {r.Filename} | {extracted.DocType} | " +
                             $"pages: {extracted.PageCount} | year: {extracted.PrimaryYear?.ToString() ?? "unknown"}");
            }
            catch (Exception ex)
            {
                messages.Add($"AFS {i+1} | {r.Filename} | ERROR: {ex.Message}");
            }
        }

        _session.SetReports(reports);
        TempData["ExtractMessages"] = string.Join("|", messages);
        return RedirectToAction(nameof(Upload));
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 4 · COMPARE
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// GET /compare — Show comparison results if available, otherwise prompt.
    /// </summary>
    [HttpGet("compare")]
    public IActionResult Compare(int pair = 0)
    {
        var cmp = GetActiveComparison();
        if (cmp == null)
        {
            TempData["CompareError"] = "No comparison found. Please run Steps 2 & 3 first.";
            return RedirectToAction(nameof(Upload));
        }
        var vm = new ComparisonViewModel
        {
            Engagement        = _session.GetEngagement(),
            Comparison        = cmp,
            Comments          = _session.GetComments(),
            SelectedPairIndex = pair,
        };
        return View(vm);
    }

    /// <summary>
    /// POST /compare — Run the full comparison pipeline on the two extracted reports.
    /// Stores ComparisonResult in session and redirects to the compare view.
    /// </summary>
    [HttpPost("compare")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> RunComparison()
    {
        var reports = _session.GetReports();
        if (reports.Count < 2)
        {
            TempData["CompareError"] = "Need at least 2 extracted PDFs.";
            return RedirectToAction(nameof(Upload));
        }

        // Ensure both reports have extracted text
        if (reports[0].PageCount == 0 || reports[1].PageCount == 0)
        {
            TempData["CompareError"] = "PDFs not yet extracted. Please run Step 3 first.";
            return RedirectToAction(nameof(Upload));
        }

        var unreadableReport = reports.FirstOrDefault(r =>
            r.WordCount == 0 ||
            (r.DocType?.Contains("corrupted", StringComparison.OrdinalIgnoreCase) ?? false));
        if (unreadableReport != null)
        {
            TempData["CompareError"] =
                $"{unreadableReport.Label} ({unreadableReport.Filename}) is not comparable. " +
                $"Extraction result: {unreadableReport.DocType}. Repair/re-save the PDF and run Step 3 again.";
            return RedirectToAction(nameof(Upload));
        }

        try
        {
            var result = await _comparer.BuildComparisonAsync(reports[0], reports[1]);
            StoreComparison(result);
            return RedirectToAction(nameof(Compare));
        }
        catch (Exception ex)
        {
            TempData["CompareError"] = "Comparison error: " + ex.Message;
            return RedirectToAction(nameof(Upload));
        }
    }

    /// <summary>
    /// POST /comment — AJAX endpoint: save a single auditor comment.
    /// Returns JSON { success: true } on success.
    /// </summary>
    [HttpPost("comment")]
    [ValidateAntiForgeryToken]
    public IActionResult SaveComment([FromForm] string key, [FromForm] string comment)
    {
        // Sanitise the comment key to prevent unexpected session keys
        if (string.IsNullOrWhiteSpace(key) ||
            !Regex.IsMatch(key, @"^(changed|page|overall):[0-9a-z_\-]*$",
                           RegexOptions.IgnoreCase))
        {
            return BadRequest(new { success = false, error = "Invalid key." });
        }
        _session.UpsertComment(key, comment ?? "");
        return Json(new { success = true });
    }

    // ─────────────────────────────────────────────────────────────────────────
    // STEP 5 · EXPORT
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// GET /export — Show the export options form.
    /// </summary>
    [HttpGet("export")]
    public IActionResult Export()
    {
        var vm = new ExportViewModel
        {
            Engagement     = _session.GetEngagement(),
            HasComparison  = GetActiveComparison() != null,
        };
        return View(vm);
    }

    /// <summary>
    /// POST /export — Generate requested file formats and save to the output folder.
    /// Shows success/error messages in the view.
    /// </summary>
    [HttpPost("export")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> RunExport(ExportViewModel vm)
    {
        var cmp = GetActiveComparison();
        if (cmp == null)
        {
            vm.Errors.Add("No comparison found, or detailed comparison expired after restart. Run Steps 2–4 again.");
            vm.HasComparison = false;
            vm.Engagement = _session.GetEngagement();
            return View("Export", vm);
        }

        var eng      = _session.GetEngagement();
        var comments = _session.GetComments();

        // Sanitise output folder: prevent path traversal.
        // Only allow absolute paths that exist or can be created.
        string folder = vm.OutputFolder?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(folder))
        {
            vm.Errors.Add("Output folder path is empty.");
            vm.HasComparison = true;
            vm.Engagement = eng;
            return View("Export", vm);
        }
        // Block relative paths and UNC paths that could reach unexpected locations
        if (!Path.IsPathRooted(folder) || folder.Contains(".."))
        {
            vm.Errors.Add("Output folder must be an absolute path.");
            vm.HasComparison = true;
            vm.Engagement = eng;
            return View("Export", vm);
        }

        try { Directory.CreateDirectory(folder); }
        catch (Exception ex) { vm.Errors.Add("Cannot create folder: " + ex.Message); }

        string ts     = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        // Sanitise client name for use in filename (remove Windows-illegal chars)
        string client = Regex.Replace(eng.Client?.Trim() ?? "AFS", @"[\\/*?:""<>|]", "_");
        if (string.IsNullOrWhiteSpace(client)) client = "AFS";

        // ── PDF ────────────────────────────────────────────────────────────
        if (vm.ExportPdf)
        {
            string p = Path.Combine(folder, $"{client}_Working_Paper_{ts}.pdf");
            try
            {
                var pdfSnapshots = await PreRenderSnapshotsAsync(cmp);
                _exporter.BuildPdf(cmp, eng, comments, p, pdfSnapshots);
                vm.SavedFiles.Add("PDF Working Paper: " + p);
            }
            catch (Exception ex) { vm.Errors.Add("PDF error: " + ex.Message); }
        }

        // ── Word (Professional Working Paper via WorkingPaperExportService) ─────
        if (vm.ExportWord)
        {
            string p = Path.Combine(folder, $"{client}_Working_Paper_{ts}.docx");
            try
            {
                byte[]? pdf1Bytes = null, pdf2Bytes = null;
                try { if (System.IO.File.Exists(cmp.Report1.PdfPath)) pdf1Bytes = await System.IO.File.ReadAllBytesAsync(cmp.Report1.PdfPath); } catch { }
                try { if (System.IO.File.Exists(cmp.Report2.PdfPath)) pdf2Bytes = await System.IO.File.ReadAllBytesAsync(cmp.Report2.PdfPath); } catch { }
                await _wpExporter.Build(cmp, eng, comments, pdf1Bytes, pdf2Bytes, p);
                vm.SavedFiles.Add("Word Working Paper: " + p);
            }
            catch (Exception ex) { vm.Errors.Add("Word error: " + ex.Message); }
        }

        // ── Excel ──────────────────────────────────────────────────────────
        if (vm.ExportExcel)
        {
            string p = Path.Combine(folder, $"{client}_AFS_Comparison_{ts}.xlsx");
            try
            {
                _exporter.BuildExcel(cmp, eng, comments, p);
                vm.SavedFiles.Add("Excel Workbook: " + p);
            }
            catch (Exception ex) { vm.Errors.Add("Excel error: " + ex.Message); }
        }

        // ── Text ───────────────────────────────────────────────────────────
        if (vm.ExportText)
        {
            string p = Path.Combine(folder, $"{client}_AFS_Comparison_{ts}.txt");
            try
            {
                string content = _exporter.BuildText(cmp, eng, comments);
                System.IO.File.WriteAllText(p, content, System.Text.Encoding.UTF8);
                vm.SavedFiles.Add("Text Report: " + p);
            }
            catch (Exception ex) { vm.Errors.Add("Text error: " + ex.Message); }
        }

        vm.HasComparison = true;
        vm.Engagement    = eng;
        return View("Export", vm);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // UTILITY · DOWNLOAD EXPORT FILES AS ZIP (optional convenience endpoint)
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// GET /download?file={path} — Allows downloading a generated working-paper
    /// file by its absolute server path.
    /// Only files inside the system's desktop AFS_Comparison folder are allowed
    /// (prevents arbitrary file read).
    /// </summary>
    [HttpGet("download")]
    public IActionResult Download(string file)
    {
        if (string.IsNullOrWhiteSpace(file))
            return BadRequest();

        // Security: restrict to Desktop\AFS_Comparison (non-traversal)
        string allowed = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "AFS_Comparison");

        string fullPath = Path.GetFullPath(file);
        if (!fullPath.StartsWith(Path.GetFullPath(allowed), StringComparison.OrdinalIgnoreCase))
            return Forbid();

        if (!System.IO.File.Exists(fullPath))
            return NotFound();

        string ext  = Path.GetExtension(fullPath).ToLowerInvariant();
        string mime = ext switch
        {
            ".pdf"  => "application/pdf",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".txt"  => "text/plain",
            _       => "application/octet-stream",
        };

        return PhysicalFile(fullPath, mime, Path.GetFileName(fullPath));
    }

    // =========================================================================
    // JSON API ENDPOINTS — used by the single-page accordion front-end
    // These mirror the Python widget event handlers and return compact JSON.
    // =========================================================================

    // ── API: Save Engagement ─────────────────────────────────────────────────
    [HttpPost("api/engagement")]
    [ValidateAntiForgeryToken]
    public IActionResult ApiSaveEngagement([FromForm] EngagementDetails details)
    {
        _session.SetEngagement(details);
        return Json(new { success = true, message = "Engagement details saved." });
    }

    // ── API: Upload PDFs ─────────────────────────────────────────────────────
    [HttpPost("api/upload")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ApiUpload(List<IFormFile> files)
    {
        if (files == null || !files.Any())
            return Json(new { success = false, message = "No files selected." });

        var uploadsDir = Path.Combine(_env.WebRootPath, "uploads");
        Directory.CreateDirectory(uploadsDir);

        var acceptedFiles = files
            .Where(file => file.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            .Take(2)
            .ToList();

        if (acceptedFiles.Count == 0)
            return Json(new { success = false, message = "Only PDF files are accepted." });

        ResetComparisonSession(deleteQueuedFiles: true);

        var reports = new List<PdfReport>();

        foreach (var file in acceptedFiles)
        {
            string safeName   = Guid.NewGuid().ToString("N") + ".pdf";
            string serverPath = Path.Combine(uploadsDir, safeName);
            await using var fs = new FileStream(serverPath, FileMode.Create);
            await file.CopyToAsync(fs);

            reports.Add(new PdfReport
            {
                Filename = file.FileName,
                PdfPath  = serverPath,
                Label    = "AFS " + (reports.Count + 1),
            });
        }

        _session.SetReports(reports);

        var queue = reports.Select((r, i) => new
        {
            label    = "AFS " + (i + 1),
            filename = r.Filename,
            kb       = System.IO.File.Exists(r.PdfPath)
                       ? (int)(new FileInfo(r.PdfPath).Length / 1024)
                       : 0,
            docType  = string.IsNullOrEmpty(r.DocType) ? "queued" : r.DocType,
        }).ToList();

        int ignoredCount = files.Count - acceptedFiles.Count;
        string ignoredMsg = ignoredCount > 0
            ? $" Ignored {ignoredCount} extra or non-PDF file(s)."
            : "";

        return Json(new
        {
            success = true,
            added = reports.Count,
            queue,
            message = $"{reports.Count} PDF(s) queued as the active comparison set.{ignoredMsg}"
        });
    }

    // ── API: Clear Queue ─────────────────────────────────────────────────────
    [HttpPost("api/clear")]
    [ValidateAntiForgeryToken]
    public IActionResult ApiClear()
    {
        ResetComparisonSession(deleteQueuedFiles: true);
        return Json(new { success = true, message = "Queue cleared." });
    }

    // ── API: Extract ─────────────────────────────────────────────────────────
    [HttpPost("api/extract")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ApiExtract()
    {
        var reports = _session.GetReports();
        if (!reports.Any())
            return Json(new { success = false, message = "No PDFs in queue." });

        var results = new List<object>();
        var errors  = new List<string>();
        var ct      = HttpContext.RequestAborted;

        for (int i = 0; i < reports.Count; i++)
        {
            var r = reports[i];
            try
            {
                var extracted = await _extractor.ExtractAsync(r.PdfPath, r.Filename, ct);
                extracted.Label   = "AFS " + (i + 1);
                extracted.PdfPath = r.PdfPath;
                reports[i]        = extracted;
                results.Add(new
                {
                    label    = extracted.Label,
                    filename = extracted.Filename,
                    docType  = extracted.DocType,
                    year     = extracted.PrimaryYear?.ToString() ?? "unknown",
                    pages    = extracted.PageCount,
                    words    = extracted.WordCount,
                });
                if (extracted.WordCount == 0)
                {
                    string reason =
                        extracted.DocType?.Contains("invalid streams", StringComparison.OrdinalIgnoreCase) == true
                            ? "PDF structure is damaged. Re-export from the source system or Print to PDF, then extract again."
                        : extracted.DocType?.Contains("AI OCR key missing", StringComparison.OrdinalIgnoreCase) == true
                            ? "The PDF needs OCR but GEMINI_API_KEY is not set. Add the key to .env or install Tesseract."
                        : extracted.DocType?.Contains("AI OCR disabled", StringComparison.OrdinalIgnoreCase) == true
                            ? "The PDF needs OCR but AI OCR is disabled. Enable AiOcr or install Tesseract."
                        : extracted.DocType?.Contains("AI OCR failed", StringComparison.OrdinalIgnoreCase) == true
                            ? "Image-based PDF — local OCR and AI OCR both failed. Retry or review the AI OCR configuration."
                        : extracted.DocType?.Contains("OCR unavailable", StringComparison.OrdinalIgnoreCase) == true
                            ? "No OCR engine found. Install Tesseract or re-save the PDF with selectable text."
                        : extracted.DocType?.Contains("corrupted", StringComparison.OrdinalIgnoreCase) == true
                            ? "PDF appears corrupted. Re-save or repair the file and try again."
                            : $"No readable text was recovered ({extracted.DocType}).";
                    errors.Add($"AFS {i + 1} — {r.Filename}: {reason}");
                }
            }
            catch (Exception ex)
            {
                results.Add(new { label = "AFS " + (i + 1), filename = r.Filename, docType = "error", year = "", pages = 0, words = 0 });
                errors.Add($"AFS {i + 1} — {r.Filename}: {ex.Message}");
            }
        }

        _session.SetReports(reports);
        return Json(new { success = errors.Count == 0, reports = results, errors,
                          message = errors.Count == 0
                              ? $"Extraction done. {reports.Count} PDF(s) extracted."
                              : $"Extracted with {errors.Count} error(s)." });
    }

    // ── API: Compare ─────────────────────────────────────────────────────────
    [HttpPost("api/compare")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ApiCompare()
    {
        var reports = _session.GetReports();
        if (reports.Count == 0)
            return Json(new
            {
                success = false,
                message = "Session expired or server restarted. Upload and extract the PDFs again."
            });

        if (reports.Count < 2)
            return Json(new
            {
                success = false,
                message = "Need 2 PDFs in the active session. Upload and extract the PDFs again."
            });

        if (reports[0].PageCount == 0 || reports[1].PageCount == 0)
            return Json(new { success = false, message = "PDFs not yet extracted. Run Step 3 first." });

        var unreadableReport = reports.FirstOrDefault(r =>
            r.WordCount == 0 ||
            (r.DocType?.Contains("corrupted", StringComparison.OrdinalIgnoreCase) ?? false));
        if (unreadableReport != null)
        {
            return Json(new
            {
                success = false,
                message =
                    $"{unreadableReport.Label} ({unreadableReport.Filename}) is not comparable. " +
                    $"Extraction result: {unreadableReport.DocType}. Repair/re-save the PDF and extract again."
            });
        }

        try
        {
            var result = await _comparer.BuildComparisonAsync(reports[0], reports[1]);
            StoreComparison(result);

            // ── Build compact JSON for frontend ──────────────────────────
            int totalLines = result.FullDiff.Count;

            // Changed lines (first 100)
            var changedLines = result.ChangedLines.Take(100).Select((d, i) => new
            {
                i         = i + 1,
                line1     = d.Line1,
                line2     = d.Line2,
                sim       = d.Similarity,
                numDiff   = d.NumDiff,
                wordDiff  = d.WordDiff.Select(w => new { w = w.Word, t = w.Tag }).ToList(),
            }).ToList();

            // Added lines (first 100)
            var addedLines = result.AddedLines.Take(100)
                .Select((d, i) => new { i = i + 1, line = d.Line2 }).ToList();

            // Removed lines (first 100)
            var removedLines = result.RemovedLines.Take(100)
                .Select((d, i) => new { i = i + 1, line = d.Line1 }).ToList();

            // Number comparison
            var nc = result.NumCmp;
            var numCmp = new
            {
                simPct      = nc.SimilarityPct,
                inBoth      = nc.InBoth.Take(50).ToList(),
                onlyAfs1    = nc.OnlyInAfs1.Take(50).ToList(),
                onlyAfs2    = nc.OnlyInAfs2.Take(50).ToList(),
                countAfs1   = nc.CountAfs1,
                countAfs2   = nc.CountAfs2,
            };

            // Page alignment
            int matched   = result.Alignment.Count(a => a.I2 >= 0);
            int unmatched = result.Alignment.Count(a => a.I2 < 0);

            var pageDiffs = result.PageDiffs.Select((p, k) => new
            {
                k         = k,
                pairIndex = p.PairIndex,
                pageAfs1  = p.PageAfs1,
                pageAfs2  = p.PageAfs2,
                alignSim  = p.AlignSim,
                pctSame   = p.PctSame,
                same      = p.Same,
                changed   = p.Changed,
                added     = p.Added,
                removed   = p.Removed,
                secondaryMatch = p.SecondaryMatch,
                unmatched = p.Unmatched,
                issues    = p.Changed + p.Added + p.Removed,
                hasDiffDetail = p.Diffs.Any(d => d.Status != "same" &&
                                                 d.Status != "same-cross-page" &&
                                                 d.Status != "same-different-page"),
            }).ToList();

            var counts = result.Counts;
            var smart = result.SmartComparison;
            var di = result.DocumentIntelligence;
            var ai = result.AiAssist;
            var smartChanges = smart?.AllDiffs.Take(200).Select(d => new
            {
                type        = d.ChangeType.ToString(),
                section     = d.Section,
                noteKey     = d.NoteKey,
                label       = d.Label,
                oldText     = d.OldText.Length > 200 ? d.OldText[..200] + "…" : d.OldText,
                newText     = d.NewText.Length > 200 ? d.NewText[..200] + "…" : d.NewText,
                oldNumbers  = d.OldNumbers,
                newNumbers  = d.NewNumbers,
                pageAfs1    = d.PageAfs1,
                pageAfs2    = d.PageAfs2,
                isMaterial  = d.IsMaterial,
                summary     = d.ChangedWordsDescription,
            }).ToList();

            return Json(new
            {
                success      = true,
                message      = "Comparison complete.",
                counts       = new
                {
                    same     = counts.GetValueOrDefault("same", 0),
                    changed  = counts.GetValueOrDefault("changed", 0),
                    added    = counts.GetValueOrDefault("added", 0),
                    removed  = counts.GetValueOrDefault("removed", 0),
                    total    = totalLines,
                },
                alignment    = new { matched, unmatched,
                    afs1Pages = reports[0].PageCount, afs2Pages = reports[1].PageCount },
                afs1Label    = reports[0].Label,
                afs2Label    = reports[1].Label,
                changedLines,
                addedLines,
                removedLines,
                numCmp,
                pageDiffs,
                smartChanges,
                smartMaterial  = smart?.MaterialCount ?? 0,
                smartNarrative = smart?.NarrativeRewrites.Count ?? 0,
                smartNumeric   = smart?.MaterialNumeric.Count ?? 0,
                aiAssist = ai == null ? null : new
                {
                    enabled = ai.Enabled,
                    provider = ai.Provider,
                    model = ai.Model,
                    status = ai.Status,
                    message = ai.Message,
                    executiveSummary = ai.ExecutiveSummary,
                    overallRisk = ai.OverallRisk,
                    generatedAt = ai.GeneratedAt,
                    highlightGuidance = ai.HighlightGuidance,
                    limitations = ai.Limitations,
                    materialFindings = ai.MaterialFindings.Select(f => new
                    {
                        title = f.Title,
                        severity = f.Severity,
                        pairIndex = f.PairIndex,
                        pageAfs1 = f.PageAfs1,
                        pageAfs2 = f.PageAfs2,
                        focusDiffOrdinals = f.FocusDiffOrdinals,
                        auditRationale = f.AuditRationale,
                        recommendedProcedure = f.RecommendedProcedure,
                        changedNumbers = f.ChangedNumbers,
                    }).ToList(),
                    pageFocus = ai.PageFocus.Select(f => new
                    {
                        pairIndex = f.PairIndex,
                        pageAfs1 = f.PageAfs1,
                        pageAfs2 = f.PageAfs2,
                        priority = f.Priority,
                        why = f.Why,
                        focusDiffOrdinals = f.FocusDiffOrdinals,
                    }).ToList(),
                },
                documentIntelligence = di == null ? null : new
                {
                    summary       = di.Summary,
                    totalPages1   = di.TotalPages1,
                    totalPages2   = di.TotalPages2,
                    sections      = di.Sections.Select(s => new
                    {
                        changeKind    = s.ChangeKind,
                        isMaterial    = s.IsMaterial,
                        description   = s.Description,
                        contentSim    = Math.Round(s.ContentSim * 100, 0),
                        afs1Title     = s.Section1?.Title,
                        afs1Start     = s.Section1?.StartPage,
                        afs1End       = s.Section1?.EndPage,
                        afs1Pages     = s.Section1?.PageCount,
                        afs2Title     = s.Section2?.Title,
                        afs2Start     = s.Section2?.StartPage,
                        afs2End       = s.Section2?.EndPage,
                        afs2Pages     = s.Section2?.PageCount,
                        lineDiffCount = s.LineDiffs.Count(d => d.Status != "same"),
                    }).ToList(),
                    materialSections = di.Sections.Count(s => s.IsMaterial),
                },
                totalChanged = result.ChangedLines.Count(),
                totalAdded   = result.AddedLines.Count(),
                totalRemoved = result.RemovedLines.Count(),
                autoReview   = result.AutoReview.Select(f => new
                {
                    id           = f.Id,
                    category     = f.Category,
                    title        = f.Title,
                    description  = f.Description,
                    severity     = f.Severity,
                    pageAfs1     = f.PageAfs1,
                    pageAfs2     = f.PageAfs2,
                    pairIndex    = f.PairIndex,
                    before       = f.Before,
                    after        = f.After,
                    pctChange    = f.PctChange,
                    absChange    = f.AbsChange,
                    confidence   = f.Confidence,
                    sectionGroup = f.SectionGroup,
                    signOff      = f.SignOff,
                }).ToList(),
                tableCellDiffs = result.TableCellDiffs.Select(t => new
                {
                    t.Row, t.Col, t.Label, t.ColHeader,
                    t.ValueAfs1, t.ValueAfs2, t.Status, t.PctChange,
                    t.PageAfs1, t.PageAfs2, t.PairIndex,
                }).ToList(),
                labeledNumberChanges = result.LabeledNumberChanges.Select(n => new
                {
                    label       = n.Label,
                    rawBefore   = n.RawBefore,
                    rawAfter    = n.RawAfter,
                    before      = n.Before,
                    after       = n.After,
                    pctChange   = n.PctChange,
                    absChange   = n.AbsChange,
                    severity    = n.Severity,
                    pageAfs1    = n.PageAfs1,
                    pageAfs2    = n.PageAfs2,
                    pairIndex   = n.PairIndex,
                    signChanged = n.SignChanged,
                }).ToList(),
                materialityThreshold = result.MaterialityThreshold,
            });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = "Comparison error: " + ex.Message });
        }
    }

    // ── API: Sign-off a System Review finding (#5) ───────────────────────────
    [HttpPost("api/finding-signoff")]
    [ValidateAntiForgeryToken]
    public IActionResult ApiFindingSignoff([FromForm] string id, [FromForm] string action)
    {
        var cmp = GetActiveComparison();
        if (cmp == null)
            return Json(new { success = false, message = "No active comparison." });

        var allowed = new HashSet<string> { "cleared", "query", "escalate", "" };
        if (!allowed.Contains(action ?? ""))
            return Json(new { success = false, message = "Invalid action." });

        var finding = cmp.AutoReview.FirstOrDefault(f => f.Id == id);
        if (finding == null)
            return Json(new { success = false, message = "Finding not found." });

        finding.SignOff = action ?? "";
        StoreComparison(cmp);
        return Json(new { success = true, id, action = finding.SignOff });
    }

    // ── API: Set materiality threshold (#6) ──────────────────────────────────
    [HttpPost("api/materiality")]
    [ValidateAntiForgeryToken]
    public IActionResult ApiSetMateriality([FromForm] decimal threshold)
    {
        var cmp = GetActiveComparison();
        if (cmp == null)
            return Json(new { success = false, message = "No active comparison." });

        cmp.MaterialityThreshold = Math.Max(0, threshold);
        StoreComparison(cmp);
        return Json(new { success = true, threshold = cmp.MaterialityThreshold });
    }

    // ── API: AI Analyze (runs after compare, non-blocking) ────────────────────
    [HttpPost("api/ai-analyze")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ApiAiAnalyze(CancellationToken cancellationToken)
    {
        var cmp = GetActiveComparison();
        if (cmp == null)
            return Json(new { success = false, message = "No comparison found. Run Step 4 first." });

        // Run multi-model AI analysis + content reader concurrently
        var assistTask = _multiAi.AnalyzeAsync(cmp, cancellationToken: cancellationToken);
        var readTask   = _aiContentReader.ReadAndCompareAsync(cmp.Report1, cmp.Report2, cancellationToken);
        await Task.WhenAll(assistTask, readTask);

        cmp.AiAssist      = assistTask.Result;
        cmp.AiContentRead = readTask.Result;

        // Resolve PairIndex on section diffs and line items using page numbers
        if (cmp.AiContentRead?.SectionDiffs != null)
        {
            foreach (var s in cmp.AiContentRead.SectionDiffs)
            {
                if (s.PairIndex.HasValue) continue;
                var pd = cmp.PageDiffs.FirstOrDefault(p =>
                    (s.PageAfs1.HasValue && p.PageAfs1 == s.PageAfs1) ||
                    (s.PageAfs2.HasValue && p.PageAfs2 == s.PageAfs2));
                if (pd != null) s.PairIndex = pd.PairIndex;
            }
        }
        if (cmp.AiContentRead?.LineItemChanges != null)
        {
            foreach (var l in cmp.AiContentRead.LineItemChanges)
            {
                if (l.PairIndex.HasValue) continue;
                var pd = cmp.PageDiffs.FirstOrDefault(p =>
                    (l.PageAfs1.HasValue && p.PageAfs1 == l.PageAfs1) ||
                    (l.PageAfs2.HasValue && p.PageAfs2 == l.PageAfs2));
                if (pd != null) l.PairIndex = pd.PairIndex;
            }
        }

        // Re-store so snapshot highlights can pick up AI severity
        StoreComparison(cmp);

        var ai   = cmp.AiAssist;
        var read = cmp.AiContentRead;

        return Json(new
        {
            success = true,
            aiAssist = ai == null ? null : new
            {
                enabled          = ai.Enabled,
                status           = ai.Status,
                message          = ai.Message,
                executiveSummary = ai.ExecutiveSummary,
                overallRisk      = ai.OverallRisk,
                generatedAt      = ai.GeneratedAt,
                highlightGuidance = ai.HighlightGuidance,
                limitations      = ai.Limitations,
                materialFindings = ai.MaterialFindings.Select(f => new
                {
                    title                = f.Title,
                    severity             = f.Severity,
                    pairIndex            = f.PairIndex,
                    pageAfs1             = f.PageAfs1,
                    pageAfs2             = f.PageAfs2,
                    focusDiffOrdinals    = f.FocusDiffOrdinals,
                    auditRationale       = f.AuditRationale,
                    recommendedProcedure = f.RecommendedProcedure,
                    changedNumbers       = f.ChangedNumbers,
                }).ToList(),
                pageFocus = ai.PageFocus.Select(f => new
                {
                    pairIndex         = f.PairIndex,
                    pageAfs1          = f.PageAfs1,
                    pageAfs2          = f.PageAfs2,
                    priority          = f.Priority,
                    why               = f.Why,
                    focusDiffOrdinals = f.FocusDiffOrdinals,
                }).ToList(),
            },
            aiContentRead = read == null ? null : new
            {
                enabled            = read.Enabled,
                status             = read.Status,
                message            = read.Message,
                overallAssessment  = read.OverallAssessment,
                riskLevel          = read.RiskLevel,
                auditConclusion    = read.AuditConclusion,
                sectionDiffs       = read.SectionDiffs.Select(s => new
                {
                    sectionName       = s.SectionName,
                    changeType        = s.ChangeType,
                    severity          = s.Severity,
                    afs1Summary       = s.Afs1Summary,
                    afs2Summary       = s.Afs2Summary,
                    changeDescription = s.ChangeDescription,
                    changedValues     = s.ChangedValues,
                    pageAfs1          = s.PageAfs1,
                    pageAfs2          = s.PageAfs2,
                    pairIndex         = s.PairIndex,
                }).ToList(),
                lineItemChanges    = read.LineItemChanges.Select(l => new
                {
                    lineItem        = l.LineItem,
                    valueInAfs1     = l.ValueInAfs1,
                    valueInAfs2     = l.ValueInAfs2,
                    pctChange       = l.PctChange,
                    materialityNote = l.MaterialityNote,
                    pageAfs1        = l.PageAfs1,
                    pageAfs2        = l.PageAfs2,
                    pairIndex       = l.PairIndex,
                }).ToList(),
                newDisclosures     = read.NewDisclosures,
                removedDisclosures = read.RemovedDisclosures,
                policyChanges      = read.PolicyChanges,
                limitations        = read.Limitations,
            },
        });
    }

    [HttpGet("api/page-diff-detail")]
    public IActionResult ApiPageDiffDetail(int pairIndex)
    {
        var cmp = GetActiveComparison();
        if (cmp == null)
            return Json(new { success = false, message = "No comparison in session." });

        PageDiffResult? pageDiff = null;
        int resolvedPairIndex = -1;

        if (pairIndex >= 0 && pairIndex < cmp.PageDiffs.Count)
        {
            pageDiff = cmp.PageDiffs[pairIndex];
            resolvedPairIndex = pairIndex;
        }
        else
        {
            var matchedRawIndices = cmp.PageDiffs
                .Select((pd, rawIdx) => (pd, rawIdx))
                .Where(x => x.pd.PageAfs1.HasValue &&
                            x.pd.PageAfs2.HasValue &&
                            !x.pd.SecondaryMatch)
                .Select(x => x.rawIdx)
                .ToList();

            if (pairIndex >= 0 && pairIndex < matchedRawIndices.Count)
            {
                resolvedPairIndex = matchedRawIndices[pairIndex];
                pageDiff = cmp.PageDiffs[resolvedPairIndex];
            }
        }

        if (pageDiff == null)
            return Json(new { success = false, message = "Pair index out of range." });

        var diffs = pageDiff.Diffs
            .Where(d => d.Status != "same" &&
                        d.Status != "same-cross-page" &&
                        d.Status != "same-different-page")
            .Take(200)
            .Select(d => new
            {
                status = d.Status,
                line1 = d.Line1,
                line2 = d.Line2,
                numDiff = d.NumDiff,
                crossPageNote = d.CrossPageNote,
                wordDiff = d.WordDiff.Select(w => new { w = w.Word, t = w.Tag }).ToList(),
            })
            .ToList();

        return Json(new
        {
            success = true,
            pairIndex = resolvedPairIndex,
            diffs,
        });
    }

    // ── API: Export ──────────────────────────────────────────────────────────
    [HttpPost("api/export")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> ApiExport(
        [FromForm] string outputFolder,
        [FromForm] bool   exportPdf,
        [FromForm] bool   exportWord,
        [FromForm] bool   exportExcel,
        [FromForm] bool   exportText)
    {
        var cmp = GetActiveComparison();
        if (cmp == null)
            return Json(new { success = false, message = "No comparison found, or detailed comparison expired after restart. Run Step 4 again." });

        var eng      = _session.GetEngagement();
        var comments = _session.GetComments();

        string folder = (outputFolder ?? "").Trim();
        if (string.IsNullOrEmpty(folder))
            folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "AFS_Comparison");

        if (!Path.IsPathRooted(folder) || folder.Contains(".."))
            return Json(new { success = false, message = "Invalid folder path." });

        try { Directory.CreateDirectory(folder); }
        catch (Exception ex)
        {
            return Json(new { success = false, message = "Cannot create folder: " + ex.Message });
        }

        string ts     = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string client = Regex.Replace(eng.Client?.Trim() ?? "AFS", @"[\\/*?:""<>|]", "_");
        if (string.IsNullOrWhiteSpace(client)) client = "AFS";

        var saved  = new List<string>();
        var errors = new List<string>();

        if (exportPdf)
        {
            string p = Path.Combine(folder, $"{client}_Working_Paper_{ts}.pdf");
            try
            {
                var pdfSnapshots = await PreRenderSnapshotsAsync(cmp);
                _exporter.BuildPdf(cmp, eng, comments, p, pdfSnapshots);
                saved.Add("PDF: " + p);
            }
            catch (Exception ex) { errors.Add("PDF: " + ex.Message); }
        }
        if (exportWord)
        {
            string p = Path.Combine(folder, $"{client}_Working_Paper_{ts}.docx");
            try
            {
                byte[]? pdf1Bytes = null, pdf2Bytes = null;
                try { if (System.IO.File.Exists(cmp.Report1.PdfPath)) pdf1Bytes = await System.IO.File.ReadAllBytesAsync(cmp.Report1.PdfPath); } catch { }
                try { if (System.IO.File.Exists(cmp.Report2.PdfPath)) pdf2Bytes = await System.IO.File.ReadAllBytesAsync(cmp.Report2.PdfPath); } catch { }
                await _wpExporter.Build(cmp, eng, comments, pdf1Bytes, pdf2Bytes, p);
                saved.Add("Word: " + p);
            }
            catch (Exception ex) { errors.Add("Word: " + ex.Message); }
        }
        if (exportExcel)
        {
            string p = Path.Combine(folder, $"{client}_AFS_Comparison_{ts}.xlsx");
            try   { _exporter.BuildExcel(cmp, eng, comments, p); saved.Add("Excel: " + p); }
            catch (Exception ex) { errors.Add("Excel: " + ex.Message); }
        }
        if (exportText)
        {
            string p = Path.Combine(folder, $"{client}_AFS_Comparison_{ts}.txt");
            try
            {
                System.IO.File.WriteAllText(p, _exporter.BuildText(cmp, eng, comments),
                    System.Text.Encoding.UTF8);
                saved.Add("Text: " + p);
            }
            catch (Exception ex) { errors.Add("Text: " + ex.Message); }
        }

        return Json(new
        {
            success = saved.Count > 0,
            saved,
            errors,
            folder,
            message = saved.Count > 0
                ? $"{saved.Count} file(s) saved to {folder}"
                : "Nothing saved — see errors.",
        });
    }

    // ── API: Get current engagement (for pre-filling form) ───────────────────
    [HttpGet("api/engagement")]
    public IActionResult ApiGetEngagement()
    {
        var e = _session.GetEngagement();
        return Json(e);
    }

    // ── API: Get queue status ─────────────────────────────────────────────────
    [HttpGet("api/queue")]
    public IActionResult ApiGetQueue()
    {
        var reports = _session.GetReports();
        var queue   = reports.Select((r, i) => new
        {
            label    = "AFS " + (i + 1),
            filename = r.Filename,
            kb       = System.IO.File.Exists(r.PdfPath)
                       ? (int)(new FileInfo(r.PdfPath).Length / 1024)
                       : 0,
            docType  = r.DocType,
            pages    = r.PageCount,
            year     = r.PrimaryYear?.ToString() ?? "",
        }).ToList();
        return Json(new { queue, hasComparison = GetActiveComparison() != null });
    }

    [HttpGet("api/model-status")]
    public IActionResult ApiModelStatus()
    {
        return Json(_aiBackendStatus.GetSnapshot());
    }

    private List<HighlightSpec> BuildSnapshotHighlightSpecs(
        ComparisonResult cmp,
        PageDiffResult pageDiff,
        int afsNum)
    {
        bool afs1 = afsNum == 1;
        var ownReport = afs1 ? cmp.Report1 : cmp.Report2;
        var oppositeReport = afs1 ? cmp.Report2 : cmp.Report1;
        int? oppositePageNumber = afs1 ? pageDiff.PageAfs2 : pageDiff.PageAfs1;
        var siblingPageTexts = GetSharedSnapshotSiblingPageTexts(cmp, pageDiff, afsNum);

        var oppositePageTexts = GetNearbySnapshotPageTexts(oppositeReport, oppositePageNumber);
        bool hasStrongChangedContext = pageDiff.Diffs.Any(d =>
            d.Status == "changed" && d.Similarity >= 0.85);
        var issueDiffs = pageDiff.Diffs
            .Where(d => d.Status != "same")
            .ToList();

        var relevantDiffs = issueDiffs
            .Where(d => afs1
                ? d.Status == "changed" || d.Status == "removed"
                : d.Status == "changed" || d.Status == "added")
            .ToList();

        var issueOrdinals = issueDiffs
            .Select((diff, index) => new { diff, ordinal = index + 1 })
            .ToDictionary(x => x.diff, x => x.ordinal);
        var aiFocusScores    = GetAiSnapshotFocusScores(cmp, pageDiff);
        var aiSeverityByOrd  = GetAiSnapshotSeverityByOrdinal(cmp, pageDiff);
        var aiProtectedOrdinals = GetAiProtectedSnapshotOrdinals(aiFocusScores, aiSeverityByOrd);

        return relevantDiffs
            .Where((d, idx) => !ShouldSuppressSnapshotDiff(
                d,
                issueDiffs,
                afsNum,
                oppositePageTexts,
                siblingPageTexts,
                oppositeReport.FullText,
                hasStrongChangedContext,
                HasSnapshotIssueCluster(relevantDiffs, idx, afsNum),
                IsAiProtectedSnapshotDiff(d, issueOrdinals, aiProtectedOrdinals)))
            .OrderByDescending(d => GetAiSnapshotFocusScore(d, issueOrdinals, aiFocusScores))
            .ThenByDescending(d => d.Status == "changed")
            .ThenByDescending(d => d.NumDiff.Count)
            .ThenByDescending(d => d.Similarity)
            .Select(d =>
            {
                var spec = BuildHighlightSpec(d, afsNum);
                if (issueOrdinals.TryGetValue(d, out int ord) &&
                    aiSeverityByOrd.TryGetValue(ord, out var sev))
                    spec.Severity = sev;
                return spec;
            })
            .Where(s => s.LineText.Length >= 4)
            .Take(50)
            .ToList();
    }

    private static Dictionary<int, int> GetAiSnapshotFocusScores(
        ComparisonResult cmp,
        PageDiffResult pageDiff)
    {
        var scores = new Dictionary<int, int>();
        var ai = cmp.AiAssist;
        if (ai == null || !string.Equals(ai.Status, "ready", StringComparison.OrdinalIgnoreCase))
            return scores;

        foreach (var pageFocus in ai.PageFocus.Where(f => f.PairIndex == pageDiff.PairIndex))
            AddAiSnapshotFocusOrdinals(scores, pageFocus.FocusDiffOrdinals, 3);

        foreach (var finding in ai.MaterialFindings.Where(f => f.PairIndex == pageDiff.PairIndex))
            AddAiSnapshotFocusOrdinals(scores, finding.FocusDiffOrdinals, 5);

        return scores;
    }

    /// <summary>Maps diff ordinal → AI severity for the current page pair, used to colour highlights.</summary>
    private static Dictionary<int, AiHighlightSeverity> GetAiSnapshotSeverityByOrdinal(
        ComparisonResult cmp,
        PageDiffResult pageDiff)
    {
        var result = new Dictionary<int, AiHighlightSeverity>();
        var ai = cmp.AiAssist;
        if (ai == null || !string.Equals(ai.Status, "ready", StringComparison.OrdinalIgnoreCase))
            return result;

        // PageFocus ordinals → Medium (unless already promoted to High by a finding)
        foreach (var pf in ai.PageFocus.Where(f => f.PairIndex == pageDiff.PairIndex))
        {
            foreach (int ord in pf.FocusDiffOrdinals.Where(o => o > 0))
                result.TryAdd(ord, AiHighlightSeverity.Medium);
        }

        // MaterialFindings ordinals → severity from the finding (High/Medium/Low → enum)
        foreach (var finding in ai.MaterialFindings.Where(f => f.PairIndex == pageDiff.PairIndex))
        {
            var sev = string.Equals(finding.Severity, "High", StringComparison.OrdinalIgnoreCase)
                ? AiHighlightSeverity.High
                : string.Equals(finding.Severity, "Medium", StringComparison.OrdinalIgnoreCase)
                    ? AiHighlightSeverity.Medium
                    : AiHighlightSeverity.Normal;

            foreach (int ord in finding.FocusDiffOrdinals.Where(o => o > 0))
            {
                // Promote: only upgrade severity, never downgrade
                if (!result.TryGetValue(ord, out var existing) || existing < sev)
                    result[ord] = sev;
            }
        }

        return result;
    }

    private static HashSet<int> GetAiProtectedSnapshotOrdinals(
        IReadOnlyDictionary<int, int> aiFocusScores,
        IReadOnlyDictionary<int, AiHighlightSeverity> aiSeverityByOrd)
    {
        return aiFocusScores.Keys
            .Concat(aiSeverityByOrd.Keys)
            .Where(o => o > 0)
            .ToHashSet();
    }

    private static bool IsAiProtectedSnapshotDiff(
        LineDiffResult diff,
        IReadOnlyDictionary<LineDiffResult, int> issueOrdinals,
        IReadOnlySet<int> aiProtectedOrdinals)
    {
        return issueOrdinals.TryGetValue(diff, out int ordinal) &&
               aiProtectedOrdinals.Contains(ordinal);
    }

    private static void AddAiSnapshotFocusOrdinals(
        IDictionary<int, int> scores,
        IEnumerable<int> ordinals,
        int weight)
    {
        foreach (int ordinal in ordinals.Where(o => o > 0).Distinct())
        {
            scores[ordinal] = scores.TryGetValue(ordinal, out int existing)
                ? existing + weight
                : weight;
        }
    }

    private static int GetAiSnapshotFocusScore(
        LineDiffResult diff,
        IReadOnlyDictionary<LineDiffResult, int> issueOrdinals,
        IReadOnlyDictionary<int, int> aiFocusScores)
    {
        if (!issueOrdinals.TryGetValue(diff, out int ordinal))
            return 0;

        return aiFocusScores.TryGetValue(ordinal, out int score) ? score : 0;
    }

    private static List<string> GetNearbySnapshotPageTexts(PdfReport report, int? pageNumber)
    {
        var texts = new List<string>();
        if (pageNumber is null || report.Pages.Count == 0)
            return texts;

        for (int i = pageNumber.Value - 2; i <= pageNumber.Value; i++)
        {
            if (i < 0 || i >= report.Pages.Count)
                continue;

            string pageText = report.Pages[i] ?? "";
            if (!string.IsNullOrWhiteSpace(pageText))
                texts.Add(pageText);
        }

        return texts;
    }

    private static List<string> GetSharedSnapshotSiblingPageTexts(
        ComparisonResult cmp,
        PageDiffResult pageDiff,
        int afsNum)
    {
        if (pageDiff.SecondaryMatch)
            return new List<string>();

        bool afs1 = afsNum == 1;
        int? ownPageNumber = afs1 ? pageDiff.PageAfs1 : pageDiff.PageAfs2;
        if (ownPageNumber is null)
            return new List<string>();

        var oppositeReport = afs1 ? cmp.Report2 : cmp.Report1;
        var texts = new List<string>();

        foreach (var sibling in cmp.PageDiffs.Where(pd =>
                     pd.PairIndex != pageDiff.PairIndex &&
                     pd.SecondaryMatch &&
                     (afs1 ? pd.PageAfs1 : pd.PageAfs2) == ownPageNumber))
        {
            int? siblingOppositePage = afs1 ? sibling.PageAfs2 : sibling.PageAfs1;
            if (siblingOppositePage is null)
                continue;

            int index = siblingOppositePage.Value - 1;
            if (index < 0 || index >= oppositeReport.Pages.Count)
                continue;

            string pageText = oppositeReport.Pages[index] ?? "";
            if (!string.IsNullOrWhiteSpace(pageText))
                texts.Add(pageText);
        }

        return texts;
    }

    private static bool ShouldSuppressSnapshotDiff(
        LineDiffResult diff,
        IReadOnlyList<LineDiffResult> peerDiffs,
        int afsNum,
        IReadOnlyList<string> oppositePageTexts,
        IReadOnlyList<string> siblingPageTexts,
        string oppositeFullText,
        bool hasStrongChangedContext,
        bool hasNearbyIssueCluster,
        bool aiProtected)
    {
        if (HasMirrorSnapshotMoveDiff(diff, peerDiffs, afsNum))
            return true;

        if (IsPageNumberOnlySnapshotDiff(diff, afsNum))
            return true;

        if (aiProtected)
            return false;

        bool afs1 = afsNum == 1;
        string ownText = afs1 ? (diff.Line1 ?? "") : (diff.Line2 ?? "");
        string otherText = afs1 ? (diff.Line2 ?? "") : (diff.Line1 ?? "");

        if (string.IsNullOrWhiteSpace(ownText))
            return true;

        if (diff.Status == "changed" && !HasTargetSideChangedTokens(diff, afsNum))
            return true;

        if (diff.Status == "changed" &&
            diff.Similarity < 0.78 &&
            HasComplementaryOppositeFragment(diff, peerDiffs, afsNum))
        {
            return true;
        }

        if (diff.Status == "changed" &&
            LooksLikeSplitSnapshotDiffArtifact(diff, peerDiffs, afsNum))
        {
            return true;
        }

        if (diff.Status is "added" or "removed")
        {
            // Only suppress pure-numeric lines that are short standalone page stamps (≤8 chars).
            // Financial table rows like "(155 190 247) (155 186 940)" are pure-numeric but LONG
            // and must be highlighted — they represent real financial differences.
            if (IsPureNumericSnapshotLine(ownText) && ownText.Trim().Length <= 8)
                return true;

            int ownAnchorCount = CountSnapshotLexicalAnchors(ownText);
            bool protectClusterLine = hasNearbyIssueCluster &&
                (LooksLikeStandaloneSnapshotHeadingOrDate(ownText) || ownAnchorCount <= 2);

            if (!protectClusterLine &&
                AppearsOnOppositeSnapshotPage(ownText, oppositePageTexts) &&
                (LooksLikeStandaloneSnapshotHeadingOrDate(ownText) ||
                 LooksLikeSnapshotFragment(ownText) ||
                 ownAnchorCount <= 2))
            return true;

            if (!protectClusterLine &&
                ownAnchorCount <= 2 &&
                AppearsInSnapshotText(ownText, oppositeFullText))
            {
                return true;
            }

            if (siblingPageTexts.Count > 0 &&
                CountUsefulSnapshotAnchorTokens(ownText) >= 3 &&
                AppearsOnOppositeSnapshotPage(ownText, siblingPageTexts))
            {
                return true;
            }

            if (!protectClusterLine &&
                hasStrongChangedContext &&
                LooksLikeStandaloneSnapshotHeadingOrDate(ownText))
                return true;
        }

        if (diff.Status == "changed")
        {
            if (LooksLikeSnapshotFragment(ownText) || LooksLikeSnapshotFragment(otherText))
            {
                int ownAnchorCount = CountSnapshotLexicalAnchors(ownText);
                int otherAnchorCount = CountSnapshotLexicalAnchors(otherText);

                if (Math.Min(ownAnchorCount, otherAnchorCount) <= 2 || diff.Similarity < 0.80)
                    return true;
            }
        }

        return false;
    }

    private static bool HasMirrorSnapshotMoveDiff(
        LineDiffResult diff,
        IReadOnlyList<LineDiffResult> issueDiffs,
        int afsNum)
    {
        if (diff.Status is not ("added" or "removed"))
            return false;

        bool afs1 = afsNum == 1;
        string ownText = afs1 ? (diff.Line1 ?? "") : (diff.Line2 ?? "");
        if (string.IsNullOrWhiteSpace(ownText))
            return false;

        string ownNorm = SnapNorm(ownText);
        if (ownNorm.Length < 4)
            return false;

        return issueDiffs.Any(other =>
        {
            if (ReferenceEquals(other, diff))
                return false;

            if (diff.Status == "removed" && other.Status != "added")
                return false;

            if (diff.Status == "added" && other.Status != "removed")
                return false;

            string mirrorText = diff.Status == "removed"
                ? (other.Line2 ?? "")
                : (other.Line1 ?? "");

            return SnapNorm(mirrorText) == ownNorm;
        });
    }

    private static bool IsPageNumberOnlySnapshotDiff(LineDiffResult diff, int afsNum)
    {
        if (!string.Equals(diff.Status, "changed", StringComparison.OrdinalIgnoreCase))
            return false;

        string targetTag = afsNum == 1 ? "removed" : "added";
        var changedWords = (diff.WordDiff ?? new List<WordDiffToken>())
            .Where(t => string.Equals(t.Tag, targetTag, StringComparison.OrdinalIgnoreCase))
            .Select(t => SnapTokenNormRaw(t.Word ?? ""))
            .Where(w => w.Length > 0)
            .ToList();

        if (changedWords.Count == 0) return false;

        // Changed tokens must ALL be plain digit strings (no letters, no brackets).
        // "(1 227)" → SnapTokenNormRaw → "1" and "227" which ARE plain digits.
        // But we only suppress if the ENTIRE LINE is a short standalone page stamp.
        if (!changedWords.All(w => Regex.IsMatch(w, @"^\d{1,5}$"))) return false;

        string lineText = (afsNum == 1 ? diff.Line1 : diff.Line2) ?? "";
        string stripped = lineText.Trim();

        // Any line with letters is a content line (label, heading, TOC entry, narrative).
        // These must NOT be suppressed — they go through BuildHighlightSpec with WholeLineMode.
        if (Regex.IsMatch(stripped, @"[a-zA-Z]")) return false;

        // Financial table rows (e.g. "2 074 640 (382 000)") are pure-number but LONG.
        // Only suppress genuinely short standalone page stamps: "37", "- 40 -", "(40)".
        // A financial amount row is always ≥ 10 characters; a page stamp is ≤ 8.
        return stripped.Length <= 8;
    }

    private static bool HasSnapshotIssueCluster(
        IReadOnlyList<LineDiffResult> diffs,
        int index,
        int afsNum)
    {
        bool afs1 = afsNum == 1;
        int nearby = 0;

        for (int i = Math.Max(0, index - 2); i <= Math.Min(diffs.Count - 1, index + 2); i++)
        {
            string text = afs1 ? (diffs[i].Line1 ?? "") : (diffs[i].Line2 ?? "");
            if (!string.IsNullOrWhiteSpace(text))
                nearby++;
        }

        return nearby >= 3;
    }

    private static bool HasTargetSideChangedTokens(LineDiffResult diff, int afsNum)
    {
        string targetTag = afsNum == 1 ? "removed" : "added";
        return (diff.WordDiff ?? new List<WordDiffToken>()).Any(t =>
            string.Equals(t.Tag, targetTag, StringComparison.OrdinalIgnoreCase) &&
            SnapTokenNormRaw(t.Word ?? "").Length > 0);
    }

    private static bool HasComplementaryOppositeFragment(
        LineDiffResult diff,
        IReadOnlyList<LineDiffResult> peerDiffs,
        int afsNum)
    {
        string targetTag = afsNum == 1 ? "removed" : "added";
        string complementaryStatus = afsNum == 1 ? "added" : "removed";

        var changedTokens = (diff.WordDiff ?? new List<WordDiffToken>())
            .Where(t => string.Equals(t.Tag, targetTag, StringComparison.OrdinalIgnoreCase))
            .Select(t => SnapTokenNormRaw(t.Word ?? ""))
            .Where(t => t.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (changedTokens.Count < 3)
            return false;

        foreach (var peer in peerDiffs)
        {
            if (ReferenceEquals(peer, diff) || !string.Equals(peer.Status, complementaryStatus, StringComparison.OrdinalIgnoreCase))
                continue;

            string peerText = afsNum == 1 ? (peer.Line2 ?? "") : (peer.Line1 ?? "");
            if (string.IsNullOrWhiteSpace(peerText))
                continue;

            var peerTokens = SplitSnapshotWords(peerText)
                .Select(SnapTokenNormRaw)
                .Where(t => t.Length > 0)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            if (peerTokens.Count == 0)
                continue;

            double coverage = (double)changedTokens.Count(t => peerTokens.Contains(t)) / changedTokens.Count;
            if (coverage >= 0.60)
                return true;
        }

        return false;
    }

    private static bool LooksLikeSplitSnapshotDiffArtifact(
        LineDiffResult diff,
        IReadOnlyList<LineDiffResult> peerDiffs,
        int afsNum)
    {
        bool afs1 = afsNum == 1;
        string targetText = afs1 ? (diff.Line1 ?? "") : (diff.Line2 ?? "");
        string oppositeText = afs1 ? (diff.Line2 ?? "") : (diff.Line1 ?? "");

        if (string.IsNullOrWhiteSpace(targetText) || string.IsNullOrWhiteSpace(oppositeText))
            return false;

        var targetTokens = SnapshotLexicalTokenSet(targetText);
        var oppositeTokens = SnapshotLexicalTokenSet(oppositeText);
        if (targetTokens.Count < 5 || oppositeTokens.Count < 3)
            return false;

        double baseCoverage = (double)oppositeTokens.Intersect(targetTokens).Count() / targetTokens.Count;
        if (baseCoverage >= 0.85)
            return false;

        var combined = new HashSet<string>(oppositeTokens, StringComparer.OrdinalIgnoreCase);
        foreach (var peer in peerDiffs)
        {
            if (ReferenceEquals(peer, diff) || peer.Status == "same")
                continue;

            string peerOppositeText = afs1 ? (peer.Line2 ?? "") : (peer.Line1 ?? "");
            if (string.IsNullOrWhiteSpace(peerOppositeText))
                continue;

            foreach (var token in SnapshotLexicalTokenSet(peerOppositeText))
                combined.Add(token);
        }

        double combinedCoverage = (double)combined.Intersect(targetTokens).Count() / targetTokens.Count;
        bool numbersMatch = LineComparatorService.ExtractNumbers(targetText)
            .SetEquals(LineComparatorService.ExtractNumbers(oppositeText));
        return numbersMatch && baseCoverage <= 0.75 && combinedCoverage >= 0.90;
    }

    private static HashSet<string> SnapshotLexicalTokenSet(string text) =>
        SplitSnapshotWords(text)
            .Select(SnapTokenNormRaw)
            .Where(t => t.Length > 2)
            .Where(t => !IsNumericLikeToken(t) && !IsConnectorToken(t))
            .Where(t => t is not "the" and not "and" and not "for" and
                        not "with" and not "of" and not "to" and
                        not "its" and not "into")
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

    private static bool AppearsInSnapshotText(string lineText, string documentText)
    {
        if (string.IsNullOrWhiteSpace(documentText))
            return false;

        return AppearsOnOppositeSnapshotPage(
            lineText,
            documentText.Split('\f').ToList());
    }

    private static bool AppearsOnOppositeSnapshotPage(string lineText, IReadOnlyList<string> pageTexts)
    {
        string targetNorm = SnapNorm(lineText);
        if (targetNorm.Length < 4)
            return false;

        foreach (var pageText in pageTexts)
        {
            foreach (var rawLine in pageText.Split('\n'))
            {
                string candidateNorm = SnapNorm(rawLine);
                if (candidateNorm.Length == 0)
                    continue;

                if (candidateNorm == targetNorm)
                    return true;

                if ((ContainsWholeLine(candidateNorm, targetNorm) ||
                     ContainsWholeLine(targetNorm, candidateNorm)) &&
                    SnapshotLineMatchScore(candidateNorm, targetNorm) >= 0.72)
                {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool IsPureNumericSnapshotLine(string text)
    {
        var tokens = SplitSnapshotWords(text)
            .Select(SnapTokenNormRaw)
            .Where(t => t.Length > 0)
            .ToList();

        return tokens.Count > 0 && tokens.All(t => IsNumericLikeToken(t) || IsConnectorToken(t));
    }

    private static bool LooksLikeSnapshotFragment(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        var tokens = SplitSnapshotWords(text);
        int anchorCount = CountSnapshotLexicalAnchors(text);
        bool startsLower = text.TrimStart().FirstOrDefault() is char first && char.IsLower(first);
        bool containsMarker = text.Contains('*') || text.Contains("(continued)", StringComparison.OrdinalIgnoreCase);

        return anchorCount <= 2 && (tokens.Count <= 6 || startsLower || containsMarker);
    }

    private static int CountUsefulSnapshotAnchorTokens(string text) =>
        SplitSnapshotWords(text).Count(IsUsefulSnapshotAnchorToken);

    private static int CountSnapshotLexicalAnchors(string text) =>
        SplitSnapshotWords(text).Count(word =>
        {
            string token = SnapTokenNormRaw(word);
            if (token.Length <= 2)
                return false;

            if (IsNumericLikeToken(token) || IsConnectorToken(token))
                return false;

            return token is not "the" and not "and" and not "for" and
                   not "with" and not "of" and not "to";
        });

    private static bool LooksLikeStandaloneSnapshotHeadingOrDate(string text)
    {
        var words = SplitSnapshotWords(text);
        if (words.Count == 0 || words.Count > 6)
            return false;

        if (Regex.IsMatch(text, @"\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b"))
            return true;

        var alphaWords = words
            .Select(w => SnapStrip(w).Trim().Trim('.', ',', ';', ':', '(', ')'))
            .Where(w => w.Any(char.IsLetter))
            .ToList();

        if (alphaWords.Count is < 2 or > 4)
            return false;

        return alphaWords.All(w => char.IsUpper(w[0]));
    }

    // ── API: Page Visual Snapshot ─────────────────────────────────────────────
    // Renders one PDF page side (AFS1 or AFS2) with yellow highlight bands over changed lines.
    // Called twice in parallel by the JS — once with afsNum=1 and once with afsNum=2.
    // Uses SkiaSharp renderer exclusively (no Poppler dependency).
    // Reference: Python _pdf_page_to_b64() + _show_page_snapshot()
    [HttpGet("api/snapshot")]
    public async Task<IActionResult> ApiSnapshot([FromQuery] int pairIndex = 0, [FromQuery] int afsNum = 1)
    {
        Response.Headers.CacheControl = "no-store, no-cache, must-revalidate";
        Response.Headers.Pragma = "no-cache";
        Response.Headers.Expires = "0";

        var cmp = GetActiveComparison();
        if (cmp == null)
            return Json(new { success = false, message = "No comparison in session." });

        PageDiffResult? pageDiff = null;
        int resolvedPairIndex = -1;

        if (pairIndex >= 0 && pairIndex < cmp.PageDiffs.Count)
        {
            pageDiff = cmp.PageDiffs[pairIndex];
            resolvedPairIndex = pairIndex;
        }
        else
        {
            var matchedRawIndices = cmp.PageDiffs
                .Select((pd, rawIdx) => (pd, rawIdx))
                .Where(x => x.pd.PageAfs1.HasValue &&
                            x.pd.PageAfs2.HasValue &&
                            !x.pd.SecondaryMatch)
                .Select(x => x.rawIdx)
                .ToList();

            if (pairIndex >= 0 && pairIndex < matchedRawIndices.Count)
            {
                resolvedPairIndex = matchedRawIndices[pairIndex];
                pageDiff = cmp.PageDiffs[resolvedPairIndex];
            }
        }

        if (pageDiff == null)
            return Json(new { success = false, message = "Pair index out of range." });

        if ((afsNum == 1 && !pageDiff.PageAfs1.HasValue) ||
            (afsNum == 2 && !pageDiff.PageAfs2.HasValue))
        {
            return Json(new
            {
                success  = true,
                b64      = (string?)null,
                pdfUrl   = (string?)null,
                pageAfs1 = pageDiff.PageAfs1.HasValue ? pageDiff.PageAfs1.Value - 1 : (int?)null,
                pageAfs2 = pageDiff.PageAfs2.HasValue ? pageDiff.PageAfs2.Value - 1 : (int?)null,
                pctSame  = pageDiff.PctSame,
                changed  = pageDiff.Changed,
                added    = pageDiff.Added,
                removed  = pageDiff.Removed,
                same     = pageDiff.Same,
                unmatched = pageDiff.Unmatched,
            });
        }

        string? pdfPath;
        int     pageIdx;
        List<HighlightSpec> highlightSpecs;

        if (afsNum == 1)
        {
            pdfPath = cmp.Report1.PdfPath;
            // PageAfs1 is 1-based; RenderPageToBase64 takes 0-based index
            pageIdx = (pageDiff.PageAfs1 ?? 1) - 1;
            highlightSpecs = BuildSnapshotHighlightSpecs(cmp, pageDiff, afsNum: 1);
        }
        else
        {
            pdfPath = cmp.Report2.PdfPath;
            // PageAfs2 is 1-based; RenderPageToBase64 takes 0-based index
            pageIdx = (pageDiff.PageAfs2 ?? 1) - 1;
            highlightSpecs = BuildSnapshotHighlightSpecs(cmp, pageDiff, afsNum: 2);
        }

        if (string.IsNullOrEmpty(pdfPath) || !System.IO.File.Exists(pdfPath))
            return Json(new { success = false, message = "PDF file not found on server." });

        try
        {
            string? b64 = RenderPageToPngWithHighlights(pdfPath, pageIdx, highlightSpecs);
            if (string.IsNullOrWhiteSpace(b64))
            {
                byte[] pdfBytes = System.IO.File.ReadAllBytes(pdfPath);
                b64 = await _snapshots.RenderPageToBase64(pdfBytes, pageIdx, highlightSpecs);
            }

            string pdfUrl   = "/uploads/" + Path.GetFileName(pdfPath) + "#page=" + (pageIdx + 1);
            return Json(new
            {
                success  = true,
                b64      = b64,
                pdfUrl   = pdfUrl,
                pairIndex = resolvedPairIndex,
                // Return 0-based page indices so JS can display +1 to the user.
                pageAfs1 = pageDiff.PageAfs1.HasValue ? pageDiff.PageAfs1.Value - 1 : (int?)null,
                pageAfs2 = pageDiff.PageAfs2.HasValue ? pageDiff.PageAfs2.Value - 1 : (int?)null,
                pctSame  = pageDiff.PctSame,
                changed  = pageDiff.Changed,
                added    = pageDiff.Added,
                removed  = pageDiff.Removed,
                same     = pageDiff.Same,
                unmatched = pageDiff.Unmatched,
            });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = "Snapshot error: " + ex.Message });
        }
    }

    // ── Poppler path discovery (mirrors Python _find_poppler()) ──────────────
    private static readonly string[] _popplerCandidates =
    {
        @"C:\Users\Mamishi.Madire\AppData\Local\Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin",
        @"C:\Program Files\poppler\Library\bin",
        @"C:\Program Files\poppler-24.02.0\Library\bin",
        @"C:\Program Files (x86)\poppler\bin",
    };

    private static string? _popplerBinCache;

    private static string? FindPopplerBin()
    {
        if (_popplerBinCache != null) return _popplerBinCache;
        foreach (var c in _popplerCandidates)
            if (System.IO.Directory.Exists(c)) { _popplerBinCache = c; return c; }
        return null;
    }

    // ── Core renderer: pdftoppm → raw PNG bytes ──────────────────────────────
    // Returns the raw PNG bytes, or null when Poppler is not available.
    private static byte[]? RenderPageToPngBytes(string? pdfPath, int pageIndex)
    {
        if (string.IsNullOrEmpty(pdfPath) || !System.IO.File.Exists(pdfPath)) return null;

        var bin = FindPopplerBin();
        if (bin == null) return null;

        var pdftoppm = System.IO.Path.Combine(bin, "pdftoppm.exe");
        if (!System.IO.File.Exists(pdftoppm)) return null;

        var prefix  = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                          "afs_snap_" + Guid.NewGuid().ToString("N"));
        int pageNum = pageIndex + 1;

        try
        {
            var psi = new System.Diagnostics.ProcessStartInfo(pdftoppm)
            {
                Arguments             = $"-png -r 220 -f {pageNum} -l {pageNum} \"{pdfPath}\" \"{prefix}\"",
                RedirectStandardError = true,
                UseShellExecute       = false,
                CreateNoWindow        = true,
            };
            using var proc = System.Diagnostics.Process.Start(psi);
            proc?.WaitForExit(15_000);

            var dir   = System.IO.Path.GetDirectoryName(prefix)!;
            var name  = System.IO.Path.GetFileName(prefix);
            var files = System.IO.Directory.GetFiles(dir, name + "*.png");
            if (files.Length == 0) return null;

            return System.IO.File.ReadAllBytes(files[0]);
        }
        finally
        {
            var dir  = System.IO.Path.GetDirectoryName(prefix)!;
            var name = System.IO.Path.GetFileName(prefix);
            foreach (var f in System.IO.Directory.GetFiles(dir, name + "*.png"))
                try { System.IO.File.Delete(f); } catch { }
        }
    }

    private static RectangleF PdfBoxToImageRect(
        double pdfLeft,
        double pdfBottom,
        double pdfRight,
        double pdfTop,
        double pageHeight,
        double scaleX,
        double scaleY,
        float padX = 2f,
        float padY = 2f)
    {
        float imgLeft = (float)(pdfLeft * scaleX) - padX;
        float imgTop = (float)((pageHeight - pdfTop) * scaleY) - padY;
        float imgRight = (float)(pdfRight * scaleX) + padX;
        float imgBottom = (float)((pageHeight - pdfBottom) * scaleY) + padY;

        float width = Math.Max(1f, imgRight - imgLeft);
        float height = Math.Max(1f, imgBottom - imgTop);
        return new RectangleF(imgLeft, imgTop, width, height);
    }

    private static HighlightSpec BuildHighlightSpec(LineDiffResult diff, int afsNum)
    {
        bool afs1 = afsNum == 1;
        string lineText = afs1 ? (diff.Line1 ?? "") : (diff.Line2 ?? "");
        if (diff.Status != "changed")
            return new HighlightSpec { LineText = lineText };

        string targetTag = afs1 ? "removed" : "added";
        string otherTag = afs1 ? "added" : "removed";
        var tokens = diff.WordDiff ?? new List<WordDiffToken>();
        var changedWords = tokens
            .Where(t => string.Equals(t.Tag, targetTag, StringComparison.OrdinalIgnoreCase))
            .Select(t => t.Word ?? "")
            .Where(w => SnapTokenNormRaw(w).Length > 0)
            .ToList();

        var sideTokens = tokens
            .Where(t => string.Equals(t.Tag, "same", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(t.Tag, targetTag, StringComparison.OrdinalIgnoreCase))
            .ToList();

        var anchorWords = new List<string>();
        var changedPhrases = new List<string>();

        bool numericTailChange = IsSnapshotNumericTailChange(tokens, lineText);
        if (numericTailChange)
        {
            var tailWords = ExtractTrailingNumericFieldWords(lineText);
            if (tailWords.Count > 0)
            {
                changedWords = tailWords.ToList();
                anchorWords.AddRange(ExtractTrailingFieldContextWords(lineText, tailWords.Count));
                changedPhrases.Add(string.Join(" ", tailWords));
            }
        }

        int i = 0;
        while (!numericTailChange && i < sideTokens.Count)
        {
            if (!string.Equals(sideTokens[i].Tag, targetTag, StringComparison.OrdinalIgnoreCase))
            {
                i++;
                continue;
            }

            int start = i;
            int end = i;
            while (end + 1 < sideTokens.Count &&
                   string.Equals(sideTokens[end + 1].Tag, targetTag, StringComparison.OrdinalIgnoreCase))
            {
                end++;
            }

            int phraseStart = Math.Max(0, start - 3);
            int phraseEnd = Math.Min(sideTokens.Count - 1, end + 3);
            var phraseWords = sideTokens
                .Skip(phraseStart)
                .Take(phraseEnd - phraseStart + 1)
                .Select(t => t.Word ?? "")
                .Where(w => SnapStrip(w).Length > 0)
                .ToList();

            string phrase = string.Join(" ", phraseWords).Trim();
            if (SnapNorm(phrase).Length >= 3)
                changedPhrases.Add(phrase);

            anchorWords.AddRange(phraseWords.Where(IsUsefulSnapshotAnchorToken));
            i = end + 1;
        }

        if (changedWords.Count == 0)
        {
            var adjacentWords = ExtractAdjacentSnapshotAnchorWords(tokens, targetTag, otherTag);
            if (adjacentWords.Count > 0)
            {
                changedWords = adjacentWords.ToList();
                anchorWords.AddRange(adjacentWords);
                changedPhrases.Add(string.Join(" ", adjacentWords));
            }
        }

        if (changedPhrases.Count == 0 && lineText.Length > 0)
            changedPhrases.Add(lineText);

        var spec = new HighlightSpec
        {
            LineText = lineText,
            ChangedWords = changedWords,
            AnchorWords = anchorWords
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList(),
            ChangedPhrases = changedPhrases
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList()
        };

        // If all changed tokens are short numeric strings (e.g. SA financial amounts
        // "2 074 640" tokenised as ["2","074","640"], or TOC page refs like "6"→"16"),
        // clear word-level paint data and use WholeLineMode so the renderer highlights
        // the entire matching line at the relaxed 0.45 threshold.
        // This covers: TOC entries, financial table rows, and any pure-number change.
        bool onlyNumbersChanged =
            spec.ChangedWords.Count > 0 &&
            spec.ChangedWords.All(w =>
                System.Text.RegularExpressions.Regex.IsMatch(
                    (w ?? "").Trim(), @"^\d{1,5}[-–]?\d{0,5}$"));

        if (onlyNumbersChanged)
        {
            spec.ChangedWords.Clear();
            if (lineText.Length > 0)
            {
                spec.ChangedPhrases = new List<string> { lineText };
            }
            spec.WholeLineMode = true;
        }

        return spec;
    }

    private sealed class SnapshotRenderLine
    {
        public List<UglyToad.PdfPig.Content.Word> Words { get; } = new();
        public string Text { get; set; } = string.Empty;
        public string NormalizedText => SnapNorm(Text);
    }

    private static readonly Regex SnapshotInvisible =
        new(@"[\p{C}\u00AD\u200B-\u200F\uFEFF\uFFF0-\uFFFD]", RegexOptions.Compiled);

    private static readonly Regex SnapshotPageStamp =
        new(@"^[-–—\s]*\(?\s*\d{1,4}\s*\)?[-–—\s]*$", RegexOptions.Compiled);

    private static string SnapStrip(string s) =>
        string.IsNullOrEmpty(s) ? s :
            SnapshotInvisible.Replace(s.Normalize(NormalizationForm.FormKD), "");

    private static string SnapNorm(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        string t = SnapStrip(s.ToLowerInvariant());
        t = Regex.Replace(t, @"\s+", " ").Trim();
        if (SnapshotPageStamp.IsMatch(t)) return string.Empty;
        t = Regex.Replace(t, @"\br\s+(?=\d)", "r");
        t = Regex.Replace(t, @"(?<=\d)\s*\.\s*(?=\d)", ".");
        t = Regex.Replace(t, @"(?<=\d)\s*,\s*(?=\d)", ",");
        return t;
    }

    private static string SnapTokenNorm(string s)
    {
        string t = SnapNorm(s);
        if (t.Length == 0) return string.Empty;
        t = Regex.Replace(t, @"^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$", "");
        return t.Trim();
    }

    private static string SnapTokenNormRaw(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;

        string t = SnapStrip(s.ToLowerInvariant());
        t = Regex.Replace(t, @"\s+", " ").Trim();
        t = Regex.Replace(t, @"\br\s+(?=\d)", "r");
        t = Regex.Replace(t, @"(?<=\d)\s*\.\s*(?=\d)", ".");
        t = Regex.Replace(t, @"(?<=\d)\s*,\s*(?=\d)", ",");
        t = Regex.Replace(t, @"^[^\p{L}\p{N}]+|[^\p{L}\p{N}]+$", "");
        return t.Trim();
    }

    private static bool IsUsefulSnapshotAnchorToken(string token)
    {
        string t = SnapTokenNormRaw(token);
        if (t.Length == 0)
            return false;

        if (IsNumericLikeToken(token))
            return true;

        if (t.Length <= 2)
            return false;

        return t is not "the" and not "and" and not "for" and not "with" and not "of" and not "to";
    }

    private static List<string> SplitSnapshotWords(string text) =>
        (text ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).ToList();

    private static List<string> ExtractTrailingNumericFieldWords(string lineText)
    {
        var words = SplitSnapshotWords(lineText);
        var tail = new List<string>();

        for (int i = words.Count - 1; i >= 0; i--)
        {
            if (IsNumericLikeToken(words[i]) || IsConnectorToken(words[i]))
            {
                tail.Insert(0, words[i]);
                continue;
            }

            break;
        }

        return tail;
    }

    private static List<string> ExtractTrailingFieldContextWords(string lineText, int tailCount)
    {
        var words = SplitSnapshotWords(lineText);
        if (words.Count == 0 || tailCount <= 0 || tailCount >= words.Count)
            return new List<string>();

        int contextEnd = words.Count - tailCount - 1;
        int contextStart = Math.Max(0, contextEnd - 2);
        return words
            .Skip(contextStart)
            .Take(contextEnd - contextStart + 1)
            .Where(IsUsefulSnapshotAnchorToken)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsSnapshotNumericTailChange(List<WordDiffToken> tokens, string lineText)
    {
        if (tokens.Count == 0)
            return false;

        var changedTokens = tokens
            .Where(t => string.Equals(t.Tag, "added", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(t.Tag, "removed", StringComparison.OrdinalIgnoreCase))
            .Select(t => t.Word ?? "")
            .Where(w => SnapTokenNormRaw(w).Length > 0)
            .ToList();

        if (changedTokens.Count == 0 || changedTokens.Any(t => !IsNumericLikeToken(t) && !IsConnectorToken(t)))
            return false;

        var words = SplitSnapshotWords(lineText);
        if (words.Count == 0 || words.Count > 10)
            return false;

        var tail = ExtractTrailingNumericFieldWords(lineText);
        return tail.Count > 0;
    }

    private static List<string> ExtractAdjacentSnapshotAnchorWords(
        List<WordDiffToken> tokens,
        string targetTag,
        string otherTag)
    {
        var anchors = new List<string>();

        for (int i = 0; i < tokens.Count; i++)
        {
            if (!string.Equals(tokens[i].Tag, otherTag, StringComparison.OrdinalIgnoreCase))
                continue;

            int left = i - 1;
            while (left >= 0 &&
                   string.Equals(tokens[left].Tag, "same", StringComparison.OrdinalIgnoreCase) &&
                   (IsNumericLikeToken(tokens[left].Word) || IsConnectorToken(tokens[left].Word)))
            {
                anchors.Insert(0, tokens[left].Word);
                left--;
            }

            int right = i + 1;
            while (right < tokens.Count &&
                   string.Equals(tokens[right].Tag, "same", StringComparison.OrdinalIgnoreCase) &&
                   (IsNumericLikeToken(tokens[right].Word) || IsConnectorToken(tokens[right].Word)))
            {
                anchors.Add(tokens[right].Word);
                right++;
            }
        }

        return anchors
            .Where(w => SnapTokenNormRaw(w).Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool HasStrongSnapshotLineMatch(
        string lineNorm,
        string specNorm,
        HighlightSpec spec,
        HashSet<string> lineTokenSet,
        double phraseMatch)
    {
        if (lineNorm.Length == 0 || specNorm.Length == 0)
            return false;

        if (lineNorm == specNorm)
            return true;

        var changedTokens = spec.ChangedWords
            .Select(SnapTokenNormRaw)
            .Where(token => token.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var anchorTokens = spec.AnchorWords
            .Select(SnapTokenNormRaw)
            .Where(token => token.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        int changedHits = changedTokens.Count(lineTokenSet.Contains);
        int anchorHits = anchorTokens.Count(lineTokenSet.Contains);
        bool orderedFragment = IsOrderedSnapshotFragmentMatch(lineNorm, specNorm);

        if (changedTokens.Count > 0)
        {
            if (changedHits == 0)
                return false;

            int requiredAnchorHits = Math.Min(2, anchorTokens.Count);
            bool hasAnchorContext = requiredAnchorHits > 0 && anchorHits >= requiredAnchorHits;
            return orderedFragment || hasAnchorContext || phraseMatch >= 0.8;
        }

        if (spec.WholeLineMode)
        {
            int requiredAnchorHits = Math.Min(2, anchorTokens.Count);
            bool hasAnchorContext = requiredAnchorHits > 0 && anchorHits >= requiredAnchorHits;
            return orderedFragment && (hasAnchorContext || phraseMatch >= 0.8);
        }

        return orderedFragment;
    }

    private static bool IsOrderedSnapshotFragmentMatch(string firstNorm, string secondNorm)
    {
        var firstTokens = firstNorm.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var secondTokens = secondNorm.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (firstTokens.Length < 3 || secondTokens.Length < 3)
            return false;

        var shorter = firstTokens.Length <= secondTokens.Length ? firstTokens : secondTokens;
        var longer = firstTokens.Length <= secondTokens.Length ? secondTokens : firstTokens;
        if ((double)shorter.Length / longer.Length < 0.6)
            return false;

        for (int start = 0; start <= longer.Length - shorter.Length; start++)
        {
            bool matches = true;
            for (int offset = 0; offset < shorter.Length; offset++)
            {
                if (!string.Equals(longer[start + offset], shorter[offset], StringComparison.Ordinal))
                {
                    matches = false;
                    break;
                }
            }

            if (matches)
                return true;
        }

        return false;
    }

    private static double SnapshotPhraseMatchScore(SnapshotRenderLine line, List<string> phrases)
    {
        if (phrases.Count == 0)
            return 0.0;

        string lineNorm = line.NormalizedText;
        if (lineNorm.Length == 0)
            return 0.0;

        double best = 0.0;
        var lineTokens = line.Words
            .Select(w => SnapTokenNormRaw(w.Text ?? ""))
            .Where(t => t.Length > 0)
            .ToList();

        foreach (var phrase in phrases)
        {
            string phraseNorm = SnapNorm(phrase ?? "");
            if (phraseNorm.Length == 0)
                continue;

            if (lineNorm.Contains(phraseNorm, StringComparison.Ordinal))
                return 1.0;

            var phraseTokens = phraseNorm
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .ToList();
            if (phraseTokens.Count == 0)
                continue;

            int hits = phraseTokens.Count(t => lineTokens.Contains(t, StringComparer.OrdinalIgnoreCase));
            best = Math.Max(best, (double)hits / phraseTokens.Count);
        }

        return best;
    }

    private static bool HasUsefulSnapshotPhrase(IEnumerable<string> phrases)
    {
        foreach (var phrase in phrases)
        {
            var tokens = SnapNorm(phrase ?? "")
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);

            foreach (var token in tokens)
            {
                if (!IsNumericLikeToken(token) && !IsConnectorToken(token))
                    return true;
            }
        }

        return false;
    }

    private static List<(int Start, int End)> FindSnapshotPhraseSpans(
        List<UglyToad.PdfPig.Content.Word> lineWords,
        IEnumerable<string> phrases)
    {
        var spans = new List<(int Start, int End)>();
        var lineTokens = lineWords
            .Select(w => SnapTokenNormRaw(w.Text ?? ""))
            .ToList();

        foreach (var phrase in phrases
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .OrderByDescending(p => p.Length))
        {
            var phraseTokens = SnapNorm(phrase)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(t => t.Length > 0)
                .ToList();

            if (phraseTokens.Count == 0 || phraseTokens.Count > lineTokens.Count)
                continue;

            for (int start = 0; start <= lineTokens.Count - phraseTokens.Count; start++)
            {
                bool matches = true;
                for (int offset = 0; offset < phraseTokens.Count; offset++)
                {
                    if (!string.Equals(lineTokens[start + offset], phraseTokens[offset], StringComparison.OrdinalIgnoreCase))
                    {
                        matches = false;
                        break;
                    }
                }

                if (matches)
                    spans.Add((start, start + phraseTokens.Count - 1));
            }
        }

        return spans
            .Distinct()
            .OrderBy(s => s.Start)
            .ToList();
    }

    private static List<RectangleF> MergeSnapshotRects(List<RectangleF> rects)
    {
        if (rects.Count <= 1)
            return rects;

        var ordered = rects
            .OrderBy(r => r.Top)
            .ThenBy(r => r.Left)
            .ToList();

        var merged = new List<RectangleF> { ordered[0] };
        for (int i = 1; i < ordered.Count; i++)
        {
            var prev = merged[^1];
            var curr = ordered[i];

            bool sameBand = Math.Abs(curr.Top - prev.Top) <= Math.Max(prev.Height, curr.Height) * 0.6f;
            float gap = curr.Left - prev.Right;
            bool closeEnough = gap <= Math.Max(8f, Math.Max(prev.Height, curr.Height) * 1.4f);

            if (sameBand && closeEnough)
            {
                float left = Math.Min(prev.Left, curr.Left);
                float top = Math.Min(prev.Top, curr.Top);
                float right = Math.Max(prev.Right, curr.Right);
                float bottom = Math.Max(prev.Bottom, curr.Bottom);
                merged[^1] = RectangleF.FromLTRB(left, top, right, bottom);
            }
            else
            {
                merged.Add(curr);
            }
        }

        return merged;
    }

    private static void AddSnapshotWordSpanRect(
        List<UglyToad.PdfPig.Content.Word> lineWords,
        int startIndex,
        int endIndex,
        double sx,
        double sy,
        double pageHeight,
        List<RectangleF> rects)
    {
        if (startIndex < 0 || endIndex < startIndex || endIndex >= lineWords.Count)
            return;

        var spanWords = lineWords.Skip(startIndex).Take(endIndex - startIndex + 1).ToList();
        if (spanWords.Count == 0)
            return;

        double left = spanWords.Min(w => w.BoundingBox.Left);
        double right = spanWords.Max(w => w.BoundingBox.Right);
        double top = spanWords.Max(w => w.BoundingBox.Top);
        double bottom = spanWords.Min(w => w.BoundingBox.Bottom);
        rects.Add(PdfBoxToImageRect(left, bottom, right, top, pageHeight, sx, sy, 3f, 2f));
    }

    private static double SnapshotWordGap(
        List<UglyToad.PdfPig.Content.Word> lineWords,
        int leftIndex,
        int rightIndex)
    {
        if (leftIndex < 0 || rightIndex >= lineWords.Count || leftIndex >= rightIndex)
            return double.MaxValue;

        return lineWords[rightIndex].BoundingBox.Left - lineWords[leftIndex].BoundingBox.Right;
    }

    private static (int Start, int End) ExpandSnapshotHighlightCluster(
        List<UglyToad.PdfPig.Content.Word> lineWords,
        int pivotIndex)
    {
        if (pivotIndex < 0 || pivotIndex >= lineWords.Count)
            return (pivotIndex, pivotIndex);

        string pivotToken = SnapTokenNormRaw(lineWords[pivotIndex].Text ?? "");
        bool pivotNumeric = IsNumericLikeToken(pivotToken) || IsConnectorToken(pivotToken);
        if (!pivotNumeric)
            return (pivotIndex, pivotIndex);

        const double clusterGapPts = 10.5;

        int start = pivotIndex;
        while (start > 0)
        {
            string prevToken = SnapTokenNormRaw(lineWords[start - 1].Text ?? "");
            if (!(IsNumericLikeToken(prevToken) || IsConnectorToken(prevToken)))
                break;

            if (SnapshotWordGap(lineWords, start - 1, start) > clusterGapPts)
                break;

            start--;
        }

        int end = pivotIndex;
        while (end + 1 < lineWords.Count)
        {
            string nextToken = SnapTokenNormRaw(lineWords[end + 1].Text ?? "");
            if (!(IsNumericLikeToken(nextToken) || IsConnectorToken(nextToken)))
                break;

            if (SnapshotWordGap(lineWords, end, end + 1) > clusterGapPts)
                break;

            end++;
        }

        return (start, end);
    }

    private static string BuildLineText(IEnumerable<UglyToad.PdfPig.Content.Word> words) =>
        string.Join(" ",
            words.Select(w => SnapStrip(w.Text ?? "").Trim())
                 .Where(w => w.Length > 0));

    private static int WordCount(string text) =>
        (text ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;

    private static bool ContainsWholeLine(string haystack, string needle) =>
        !string.IsNullOrWhiteSpace(needle) &&
        (" " + haystack + " ").Contains(" " + needle + " ", StringComparison.Ordinal);

    private static bool IsSnapshotLineMatch(string lineNorm, string specNorm)
    {
        return SnapshotLineMatchScore(lineNorm, specNorm) >= 0.72;
    }

    private static double SnapshotLineMatchScore(string lineNorm, string specNorm)
    {
        if (string.IsNullOrWhiteSpace(lineNorm) || string.IsNullOrWhiteSpace(specNorm))
            return 0.0;

        if (lineNorm == specNorm)
            return 1.0;

        double best = 0.0;

        if (specNorm.Length >= 30 && lineNorm.Length >= 30)
        {
            string shorter = specNorm.Length <= lineNorm.Length ? specNorm : lineNorm;
            string longer = specNorm.Length <= lineNorm.Length ? lineNorm : specNorm;
            double ratio = (double)shorter.Length / longer.Length;
            if (ratio >= 0.60 && longer.StartsWith(shorter, StringComparison.Ordinal))
                best = Math.Max(best, 0.90 * ratio);
        }

        var lineTokens = lineNorm.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var specTokens = specNorm.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (lineTokens.Length >= 6 && specTokens.Length >= 6)
        {
            var lineSet = lineTokens.ToHashSet(StringComparer.Ordinal);
            var specSet = specTokens.ToHashSet(StringComparer.Ordinal);
            int common = lineSet.Intersect(specSet).Count();
            int maxWc = Math.Max(lineSet.Count, specSet.Count);
            int minWc = Math.Min(lineSet.Count, specSet.Count);
            double wordOverlap = (double)common / maxWc;
            double wordRatio = (double)minWc / maxWc;
            best = Math.Max(best, wordOverlap * 0.75 + wordRatio * 0.25);
        }

        return best;
    }

    private static bool IsNumericLikeToken(string token)
    {
        if (string.IsNullOrWhiteSpace(token))
            return false;

        string t = SnapTokenNormRaw(token);
        return Regex.IsMatch(t, @"^(?:r?\d+(?:[.,]\d+)?|\d+\-\d+)$", RegexOptions.IgnoreCase);
    }

    private static bool IsConnectorToken(string token)
    {
        if (string.IsNullOrWhiteSpace(token))
            return false;

        string t = SnapStrip(token).Trim();
        return t is "-" or "–" or "—" or "/" or ":" or "to";
    }

    private static List<SnapshotRenderLine> BuildSnapshotLines(UglyToad.PdfPig.Content.Page pdfPage, double pageHeight)
    {
        var raw = pdfPage.GetWords()
            .GroupBy(w =>
            {
                double top = pageHeight - w.BoundingBox.Top;
                return (int)(Math.Round(top / 8.0) * 8);
            })
            .OrderBy(g => g.Key)
            .Select(g =>
            {
                var words = g.OrderBy(w => w.BoundingBox.Left).ToList();
                var line = new SnapshotRenderLine { Text = BuildLineText(words) };
                line.Words.AddRange(words);
                return line;
            })
            .Where(l => !string.IsNullOrWhiteSpace(l.Text))
            .ToList();

        if (raw.Count == 0) return raw;

        var merged = new List<SnapshotRenderLine> { raw[0] };
        for (int i = 1; i < raw.Count; i++)
        {
            var prev = merged[^1];
            var curr = raw[i];

            if (string.Equals(prev.NormalizedText, curr.NormalizedText, StringComparison.OrdinalIgnoreCase))
                continue;

            if (WordCount(curr.Text) <= 3 && ContainsWholeLine(prev.NormalizedText, curr.NormalizedText))
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

        var deduped = new List<SnapshotRenderLine>(merged.Count);
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

    // ── Highlight renderer: pdftoppm PNG + PdfPig word positions + System.Drawing overlay ──
    // Uses Poppler for the base page image and stricter word/line matching so the
    // overlay marks the actual changed text instead of broad empty bands.
    private static string? RenderPageToPngWithHighlights(
        string? pdfPath, int pageIndex, List<HighlightSpec>? specs)
    {
        byte[]? rawPng = RenderPageToPngBytes(pdfPath, pageIndex);
        if (rawPng == null) return null;

        // No highlights? Return the plain PNG.
        if (specs == null || specs.Count == 0)
            return Convert.ToBase64String(rawPng);

        try
        {
            // ── Get page dimensions and word positions via PdfPig ─────────────
            using var pdfDoc  = PdfDocument.Open(pdfPath!);
            var pdfPage       = pdfDoc.GetPage(pageIndex + 1); // PdfPig is 1-based
            double pageWidth  = pdfPage.Width;
            double pageHeight = pdfPage.Height;

            if (pageWidth <= 0 || pageHeight <= 0)
                return Convert.ToBase64String(rawPng);

            // ── Load the PNG into a 32bppARGB bitmap for alpha blending ───────
            using var srcStream = new System.IO.MemoryStream(rawPng);
            using var srcBmp    = new Bitmap(srcStream);

            int imgW = srcBmp.Width;
            int imgH = srcBmp.Height;
            double sx = imgW / pageWidth;
            double sy = imgH / pageHeight;

            // Create a result canvas (32bpp so semi-transparent fills work)
            using var result  = new Bitmap(imgW, imgH, PixelFormat.Format32bppArgb);
            using var g       = Graphics.FromImage(result);
            g.CompositingMode    = CompositingMode.SourceOver;
            g.CompositingQuality = CompositingQuality.HighQuality;

            // Draw at explicit pixel size. The simpler DrawImage(image, x, y) overload
            // uses the source image's physical DPI, which shrinks the Poppler raster
            // inside the canvas and makes highlight coordinates drift.
            g.DrawImage(srcBmp, new Rectangle(0, 0, imgW, imgH));

            var renderLines = BuildSnapshotLines(pdfPage, pageHeight);
            if (renderLines.Count == 0)
                return Convert.ToBase64String(rawPng);

            var lineTokenSets = renderLines
                .Select(line => line.Words
                    .Select(w => SnapTokenNormRaw(w.Text ?? ""))
                    .Where(t => t.Length > 0)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase))
                .ToList();

            using var fillBrush  = new SolidBrush(Color.FromArgb(100, 255, 220, 0));
            using var outlinePen = new Pen(Color.FromArgb(220, 200, 140, 0), 1.2f);

            var lineAssignments = new Dictionary<int, List<HighlightSpec>>();
            foreach (var spec in specs)
            {
                string specNorm = SnapNorm((spec.LineText ?? "").Trim());
                if (specNorm.Length < 3)
                    continue;

                var specTokens = specNorm
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                    .Select(SnapTokenNormRaw)
                    .Where(t => t.Length > 0)
                    .ToList();
                if (specTokens.Count == 0)
                    continue;

                var specChangedTokens = spec.ChangedWords
                    .Select(w => SnapTokenNormRaw(w ?? ""))
                    .Where(w => w.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var specAnchorTokens = spec.AnchorWords
                    .Select(w => SnapTokenNormRaw(w ?? ""))
                    .Where(w => w.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                int bestLineIndex = -1;
                double bestScore = 0.0;
                bool hasUsefulPhrase = HasUsefulSnapshotPhrase(spec.ChangedPhrases);
                for (int lineIndex = 0; lineIndex < renderLines.Count; lineIndex++)
                {
                    var lineTokenSet = lineTokenSets[lineIndex];
                    if (lineTokenSet.Count == 0)
                        continue;

                    int overlapHits = specTokens.Count(t => lineTokenSet.Contains(t));
                    double overlapScore = (double)overlapHits / Math.Max(specTokens.Count, lineTokenSet.Count);
                    double lineMatch = SnapshotLineMatchScore(renderLines[lineIndex].NormalizedText, specNorm);
                    double phraseMatch = hasUsefulPhrase
                        ? SnapshotPhraseMatchScore(renderLines[lineIndex], spec.ChangedPhrases)
                        : 0.0;
                    if (!HasStrongSnapshotLineMatch(
                            renderLines[lineIndex].NormalizedText,
                            specNorm,
                            spec,
                            lineTokenSet,
                            phraseMatch))
                    {
                        continue;
                    }

                    double score;

                    if (specChangedTokens.Count > 0 || specAnchorTokens.Count > 0)
                    {
                        int changedHits = specChangedTokens.Count(t => lineTokenSet.Contains(t));
                        double changedCoverage = specChangedTokens.Count > 0
                            ? (double)changedHits / specChangedTokens.Count
                            : 0.0;

                        int anchorHits = specAnchorTokens.Count(t => lineTokenSet.Contains(t));
                        double anchorCoverage = specAnchorTokens.Count > 0
                            ? (double)anchorHits / specAnchorTokens.Count
                            : 0.0;

                        double phraseScore = phraseMatch * 0.95 + changedCoverage * 0.05;
                        double structuralScore =
                            (changedCoverage * 0.45) +
                            (anchorCoverage * 0.35) +
                            (lineMatch * 0.20);
                        double guardedOverlap = specAnchorTokens.Count > 0
                            ? overlapScore * 0.35
                            : overlapScore;

                        if (specAnchorTokens.Count > 0 && anchorCoverage == 0)
                        {
                            // Numeric-only phrases like "6" or "31-32" must not steal the
                            // match from the real TOC row unless the full line text is also close.
                            score = lineMatch >= 0.55
                                ? Math.Max(lineMatch * 0.75, phraseScore * 0.25)
                                : Math.Max(changedCoverage * 0.20, guardedOverlap * 0.20);
                        }
                        else
                        {
                            score = Math.Max(phraseScore, Math.Max(structuralScore, guardedOverlap));
                        }
                    }
                    else
                    {
                        score = Math.Max(Math.Max(phraseMatch * 0.95, lineMatch), overlapScore);
                    }

                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestLineIndex = lineIndex;
                    }
                }

                // WholeLineMode specs (TOC page-number changes, added/removed whole lines)
                // need a relaxed threshold — the content match is still meaningful.
                double threshold = spec.WholeLineMode ? 0.45
                    : (specChangedTokens.Count > 0 || specAnchorTokens.Count > 0) ? 0.40
                    : 0.72;
                if (bestLineIndex >= 0 && bestScore >= threshold)
                {
                    if (!lineAssignments.TryGetValue(bestLineIndex, out var assignedSpecs))
                    {
                        assignedSpecs = new List<HighlightSpec>();
                        lineAssignments[bestLineIndex] = assignedSpecs;
                    }

                    assignedSpecs.Add(spec);
                }
            }

            for (int lineIndex = 0; lineIndex < renderLines.Count; lineIndex++)
            {
                var lineWords = renderLines[lineIndex].Words;
                if (lineWords.Count == 0) continue;
                if (!lineAssignments.TryGetValue(lineIndex, out var matchedSpecs) || matchedSpecs.Count == 0)
                    continue;

                var changedSet = matchedSpecs
                    .SelectMany(spec => spec.ChangedWords)
                    .Select(w => SnapTokenNormRaw(w ?? ""))
                    .Where(w => w.Length >= 1)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                var wordRects = new List<RectangleF>();
                bool highlightWholeLine = matchedSpecs.Any(spec => spec.ChangedWords.Count == 0) || changedSet.Count == 0;
                bool numericOnlyChangedSet = changedSet.Count > 0 &&
                    changedSet.All(token => IsNumericLikeToken(token) || IsConnectorToken(token));

                if (highlightWholeLine)
                {
                    double left = lineWords.Min(w => w.BoundingBox.Left);
                    double right = lineWords.Max(w => w.BoundingBox.Right);
                    double top = lineWords.Max(w => w.BoundingBox.Top);
                    double bottom = lineWords.Min(w => w.BoundingBox.Bottom);
                    wordRects.Add(PdfBoxToImageRect(left, bottom, right, top, pageHeight, sx, sy));
                }
                else
                {
                    var matchedWordIndexes = new List<int>();
                    for (int wordIndex = 0; wordIndex < lineWords.Count; wordIndex++)
                    {
                        string token = SnapTokenNormRaw(lineWords[wordIndex].Text ?? "");
                        if (token.Length > 0 && changedSet.Contains(token))
                            matchedWordIndexes.Add(wordIndex);
                    }

                    if (matchedWordIndexes.Count > 0)
                    {
                        var groupedIndexes = matchedWordIndexes
                            .GroupBy(idx =>
                            {
                                var word = lineWords[idx];
                                double top = pageHeight - word.BoundingBox.Top;
                                return (int)(Math.Round(top / 5.0) * 5);
                            })
                            .ToList();

                        var expandedRanges = new List<(int Start, int End)>();
                        foreach (var bandGroup in groupedIndexes)
                        {
                            var sorted = bandGroup.OrderBy(x => x).ToList();
                            if (sorted.Count == 0)
                                continue;

                            foreach (int wordIndex in sorted)
                            {
                                expandedRanges.Add(ExpandSnapshotHighlightCluster(lineWords, wordIndex));
                            }
                        }

                        var mergedRanges = expandedRanges
                            .Distinct()
                            .OrderBy(r => r.Start)
                            .ToList();

                        if (mergedRanges.Count > 0)
                        {
                            int start = mergedRanges[0].Start;
                            int end = mergedRanges[0].End;

                            for (int i = 1; i < mergedRanges.Count; i++)
                            {
                                if (mergedRanges[i].Start <= end + 1)
                                {
                                    end = Math.Max(end, mergedRanges[i].End);
                                }
                                else
                                {
                                    AddSnapshotWordSpanRect(lineWords, start, end, sx, sy, pageHeight, wordRects);
                                    start = mergedRanges[i].Start;
                                    end = mergedRanges[i].End;
                                }
                            }

                            AddSnapshotWordSpanRect(lineWords, start, end, sx, sy, pageHeight, wordRects);
                        }
                    }

                    if (wordRects.Count == 0 && !numericOnlyChangedSet)
                    {
                        var phraseSpans = FindSnapshotPhraseSpans(
                            lineWords,
                            matchedSpecs.SelectMany(spec => spec.ChangedPhrases));

                        foreach (var (start, end) in phraseSpans)
                            AddSnapshotWordSpanRect(lineWords, start, end, sx, sy, pageHeight, wordRects);
                    }

                    if (wordRects.Count == 0)
                    {
                        double left = lineWords.Min(w => w.BoundingBox.Left);
                        double right = lineWords.Max(w => w.BoundingBox.Right);
                        double top = lineWords.Max(w => w.BoundingBox.Top);
                        double bottom = lineWords.Min(w => w.BoundingBox.Bottom);
                        wordRects.Add(PdfBoxToImageRect(left, bottom, right, top, pageHeight, sx, sy));
                    }
                }

                foreach (var rect in MergeSnapshotRects(wordRects))
                {
                    g.FillRectangle(fillBrush, rect);
                    g.DrawRectangle(outlinePen, rect.X, rect.Y, rect.Width, rect.Height);
                }
            }

            using var outMs = new System.IO.MemoryStream();
            result.Save(outMs, ImageFormat.Png);
            return Convert.ToBase64String(outMs.ToArray());
        }
        catch
        {
            // Fallback: return plain PNG if anything goes wrong
            return Convert.ToBase64String(rawPng);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // UTILITY · Pre-render all page snapshot pairs for PDF export
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Renders every paired page from both PDFs into Base64 PNG strings
    /// (with yellow highlights) so that the synchronous QuestPDF builder can
    /// embed them without needing to await inside the content pipeline.
    /// Returns one tuple per paired PageDiff, in order.
    /// </summary>
    private async Task<List<(string? Left, string? Right)>> PreRenderSnapshotsAsync(
        ComparisonResult cmp)
    {
        byte[]? pdf1 = null, pdf2 = null;
        try { if (System.IO.File.Exists(cmp.Report1.PdfPath)) pdf1 = await System.IO.File.ReadAllBytesAsync(cmp.Report1.PdfPath); } catch { }
        try { if (System.IO.File.Exists(cmp.Report2.PdfPath)) pdf2 = await System.IO.File.ReadAllBytesAsync(cmp.Report2.PdfPath); } catch { }

        var result = new List<(string? Left, string? Right)>();
        foreach (var pd in cmp.PageDiffs.Where(pd =>
            pd.PageAfs1.HasValue &&
            pd.PageAfs2.HasValue &&
            !pd.Unmatched &&
            !pd.SecondaryMatch))
        {
            int p1i = pd.PageAfs1!.Value - 1;
            int p2i = pd.PageAfs2!.Value - 1;
            var hl1 = BuildSnapshotHighlightSpecs(cmp, pd, 1);
            var hl2 = BuildSnapshotHighlightSpecs(cmp, pd, 2);
            string? left = null, right = null;
            try { left = RenderPageToPngWithHighlights(cmp.Report1.PdfPath, p1i, hl1); } catch { }
            try { right = RenderPageToPngWithHighlights(cmp.Report2.PdfPath, p2i, hl2); } catch { }
            try { if (string.IsNullOrWhiteSpace(left) && pdf1 != null) left = await _snapshots.RenderPageToBase64(pdf1, p1i, hl1); } catch { }
            try { if (string.IsNullOrWhiteSpace(right) && pdf2 != null) right = await _snapshots.RenderPageToBase64(pdf2, p2i, hl2); } catch { }
            result.Add((left, right));
        }
        return result;
    }
}
