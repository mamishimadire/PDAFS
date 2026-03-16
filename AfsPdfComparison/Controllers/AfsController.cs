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
    private readonly PageExtractorService _extractor;
    private readonly ComparisonService    _comparer;
    private readonly ExportService        _exporter;
    private readonly SessionStateService  _session;
    private readonly IWebHostEnvironment  _env;
    private readonly PageSnapshotService  _snapshots;

    public AfsController(
        PageExtractorService extractor,
        ComparisonService    comparer,
        ExportService        exporter,
        SessionStateService  session,
        IWebHostEnvironment  env,
        PageSnapshotService  snapshots)
    {
        _extractor = extractor;
        _comparer  = comparer;
        _exporter  = exporter;
        _session   = session;
        _env       = env;
        _snapshots = snapshots;
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

        var reports = _session.GetReports();
        var existing = reports.Select(r => r.Filename).ToHashSet();

        foreach (var file in files)
        {
            // Security: only accept PDFs identified by extension and content type.
            if (!file.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                !file.ContentType.Contains("pdf", StringComparison.OrdinalIgnoreCase))
            {
                TempData["UploadError"] = "Only PDF files are accepted.";
                continue;
            }

            if (existing.Contains(file.FileName)) continue;

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
            existing.Add(file.FileName);
        }

        _session.SetReports(reports);
        TempData["UploadSuccess"] = files.Count + " file(s) queued.";
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
        foreach (var r in _session.GetReports())
        {
            try { if (System.IO.File.Exists(r.PdfPath)) System.IO.File.Delete(r.PdfPath); }
            catch { /* best effort */ }
        }
        _session.SetReports(new());
        _session.SetComparison(null!);
        _session.SetComments(new());
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
    public IActionResult Extract()
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
                var extracted = _extractor.Extract(r.PdfPath, r.Filename);
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
        var cmp = _session.GetComparison();
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
    public IActionResult RunComparison()
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

        try
        {
            var result = _comparer.BuildComparison(reports[0], reports[1]);
            _session.SetComparison(result);
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
            HasComparison  = _session.GetComparison() != null,
        };
        return View(vm);
    }

    /// <summary>
    /// POST /export — Generate requested file formats and save to the output folder.
    /// Shows success/error messages in the view.
    /// </summary>
    [HttpPost("export")]
    [ValidateAntiForgeryToken]
    public IActionResult RunExport(ExportViewModel vm)
    {
        var cmp = _session.GetComparison();
        if (cmp == null)
        {
            vm.Errors.Add("No comparison found. Run Steps 2–4 first.");
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
                _exporter.BuildPdf(cmp, eng, comments, p);
                vm.SavedFiles.Add("PDF Working Paper: " + p);
            }
            catch (Exception ex) { vm.Errors.Add("PDF error: " + ex.Message); }
        }

        // ── Word ───────────────────────────────────────────────────────────
        if (vm.ExportWord)
        {
            string p = Path.Combine(folder, $"{client}_Working_Paper_{ts}.docx");
            try
            {
                _exporter.BuildWord(cmp, eng, comments, p);
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

        var reports  = _session.GetReports();
        var existing = reports.Select(r => r.Filename).ToHashSet();
        int added    = 0;

        foreach (var file in files)
        {
            if (!file.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                continue;
            if (existing.Contains(file.FileName)) continue;

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
            existing.Add(file.FileName);
            added++;
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

        return Json(new { success = true, added, queue,
                          message = $"{added} PDF(s) added. Total: {reports.Count}" });
    }

    // ── API: Clear Queue ─────────────────────────────────────────────────────
    [HttpPost("api/clear")]
    [ValidateAntiForgeryToken]
    public IActionResult ApiClear()
    {
        foreach (var r in _session.GetReports())
        {
            try { if (System.IO.File.Exists(r.PdfPath)) System.IO.File.Delete(r.PdfPath); }
            catch { /* best effort */ }
        }
        _session.SetReports(new());
        _session.SetComparison(null!);
        _session.SetComments(new());
        return Json(new { success = true, message = "Queue cleared." });
    }

    // ── API: Extract ─────────────────────────────────────────────────────────
    [HttpPost("api/extract")]
    [ValidateAntiForgeryToken]
    public IActionResult ApiExtract()
    {
        var reports = _session.GetReports();
        if (!reports.Any())
            return Json(new { success = false, message = "No PDFs in queue." });

        var results  = new List<object>();
        var errors   = new List<string>();

        for (int i = 0; i < reports.Count; i++)
        {
            var r = reports[i];
            try
            {
                var extracted  = _extractor.Extract(r.PdfPath, r.Filename);
                extracted.Label   = "AFS " + (i + 1);
                extracted.PdfPath = r.PdfPath;
                reports[i]        = extracted;
                results.Add(new
                {
                    label     = extracted.Label,
                    filename  = extracted.Filename,
                    docType   = extracted.DocType,
                    year      = extracted.PrimaryYear?.ToString() ?? "unknown",
                    pages     = extracted.PageCount,
                    words     = extracted.WordCount,
                });
            }
            catch (Exception ex)
            {
                errors.Add($"AFS {i+1} — {r.Filename}: {ex.Message}");
                results.Add(new
                {
                    label    = "AFS " + (i + 1),
                    filename = r.Filename,
                    docType  = "error",
                    year     = "",
                    pages    = 0,
                    words    = 0,
                });
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
    public IActionResult ApiCompare()
    {
        var reports = _session.GetReports();
        if (reports.Count < 2)
            return Json(new { success = false, message = "Need at least 2 extracted PDFs." });

        if (reports[0].PageCount == 0 || reports[1].PageCount == 0)
            return Json(new { success = false, message = "PDFs not yet extracted. Run Step 3 first." });

        try
        {
            var result = _comparer.BuildComparison(reports[0], reports[1]);
            _session.SetComparison(result);

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
                pageAfs1  = p.PageAfs1,
                pageAfs2  = p.PageAfs2,
                alignSim  = p.AlignSim,
                pctSame   = p.PctSame,
                same      = p.Same,
                changed   = p.Changed,
                added     = p.Added,
                removed   = p.Removed,
                issues    = p.Changed + p.Added + p.Removed,
                diffs     = p.Diffs.Where(d => d.Status != "same").Take(25).Select(d => new
                {
                    status  = d.Status,
                    line1   = d.Line1,
                    line2   = d.Line2,
                    numDiff = d.NumDiff,
                    wordDiff = d.WordDiff.Select(w => new { w = w.Word, t = w.Tag }).ToList(),
                }).ToList(),
            }).ToList();

            var counts = result.Counts;
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
                totalChanged = result.ChangedLines.Count(),
                totalAdded   = result.AddedLines.Count(),
                totalRemoved = result.RemovedLines.Count(),
            });
        }
        catch (Exception ex)
        {
            return Json(new { success = false, message = "Comparison error: " + ex.Message });
        }
    }

    // ── API: Export ──────────────────────────────────────────────────────────
    [HttpPost("api/export")]
    [ValidateAntiForgeryToken]
    public IActionResult ApiExport(
        [FromForm] string outputFolder,
        [FromForm] bool   exportPdf,
        [FromForm] bool   exportWord,
        [FromForm] bool   exportExcel,
        [FromForm] bool   exportText)
    {
        var cmp = _session.GetComparison();
        if (cmp == null)
            return Json(new { success = false, message = "No comparison found. Run Step 4 first." });

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
            try   { _exporter.BuildPdf(cmp, eng, comments, p); saved.Add("PDF: " + p); }
            catch (Exception ex) { errors.Add("PDF: " + ex.Message); }
        }
        if (exportWord)
        {
            string p = Path.Combine(folder, $"{client}_Working_Paper_{ts}.docx");
            try   { _exporter.BuildWord(cmp, eng, comments, p); saved.Add("Word: " + p); }
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
        return Json(new { queue, hasComparison = _session.GetComparison() != null });
    }

    // ── API: Page Visual Snapshot ─────────────────────────────────────────────
    // Renders a PDF page to PNG using Poppler pdftoppm (already installed by notebook).
    // Returns base64-encoded PNG strings for both pages in the selected pair.
    // Reference: Python _pdf_page_to_b64() + _show_page_snapshot()
    [HttpGet("api/snapshot")]
    public IActionResult ApiSnapshot([FromQuery] int pairIndex)
    {
        var cmp = _session.GetComparison();
        if (cmp == null)
            return Json(new { success = false, message = "No comparison found. Run Step 4 first." });

        if (pairIndex < 0 || pairIndex >= cmp.PageDiffs.Count)
            return Json(new { success = false, message = "Pair index out of range." });

        var pd  = cmp.PageDiffs[pairIndex];
        string? b1 = null, b2 = null;

        // Collect highlight texts — lines that changed/were removed on AFS1 side,
        // and lines that changed/were added on AFS2 side (mirrors Python hl1/hl2).
        var hl1 = pd.Diffs
            .Where(d => d.Status == "changed" || d.Status == "removed")
            .Select(d => d.Line1)
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToArray();
        var hl2 = pd.Diffs
            .Where(d => d.Status == "changed" || d.Status == "added")
            .Select(d => d.Line2)
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToArray();

        try { b1 = RenderPageToPngWithHighlights(cmp.Report1.PdfPath, (pd.PageAfs1 ?? 1) - 1, hl1); } catch { }
        if (pd.PageAfs2.HasValue)
            try { b2 = RenderPageToPngWithHighlights(cmp.Report2.PdfPath, pd.PageAfs2.Value - 1, hl2); } catch { }

        // BUG FIX #2: Fallback to SkiaSharp-based renderer when Poppler is not available.
        // PageSnapshotService uses PdfPig + SkiaSharp (no external process required),
        // ensuring yellow highlight bands always appear even without Poppler installed.
        if (b1 == null && !string.IsNullOrEmpty(cmp.Report1.PdfPath) &&
            System.IO.File.Exists(cmp.Report1.PdfPath))
        {
            try
            {
                var bytes = System.IO.File.ReadAllBytes(cmp.Report1.PdfPath);
                b1 = _snapshots.RenderPageToBase64(bytes, (pd.PageAfs1 ?? 1) - 1, hl1.ToList());
                if (b1 == "") b1 = null;
            }
            catch { }
        }
        if (b2 == null && pd.PageAfs2.HasValue && !string.IsNullOrEmpty(cmp.Report2.PdfPath) &&
            System.IO.File.Exists(cmp.Report2.PdfPath))
        {
            try
            {
                var bytes = System.IO.File.ReadAllBytes(cmp.Report2.PdfPath);
                b2 = _snapshots.RenderPageToBase64(bytes, pd.PageAfs2.Value - 1, hl2.ToList());
                if (b2 == "") b2 = null;
            }
            catch { }
        }

        bool ok = b1 != null || b2 != null;
        return Json(new
        {
            success   = ok,
            b64Afs1   = b1,
            b64Afs2   = b2,
            afs1Label = $"AFS 1 — page {pd.PageAfs1} — {cmp.Report1.Filename}",
            afs2Label = pd.PageAfs2.HasValue
                ? $"AFS 2 — page {pd.PageAfs2} — {cmp.Report2.Filename}"
                : "UNMATCHED",
            message   = ok ? "ok" : "Images unavailable — Poppler not found.",
            pctSame   = pd.PctSame,
            same      = pd.Same,
            changed   = pd.Changed,
            added     = pd.Added,
            removed   = pd.Removed,
        });
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
                Arguments             = $"-png -r 130 -f {pageNum} -l {pageNum} \"{pdfPath}\" \"{prefix}\"",
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

    // ── Highlight renderer: pdftoppm PNG + PdfPig word positions + System.Drawing overlay ──
    // Mirrors Python _pdf_page_to_b64(... highlight_texts=...) which uses pdfplumber +
    // PIL ImageDraw.rectangle with fill=(255,220,0,90) and outline=(255,160,0,220).
    //
    // Coordinate mapping (PDF → image):
    //   PDF origin = bottom-left  (PdfPig BoundingBox.Bottom/Top in points)
    //   Image origin = top-left   (pixels, Y increases downward)
    //   sx = imageWidth / page.Width    sy = imageHeight / page.Height
    //   pixelY_top    = (pageHeight - word.BoundingBox.Top)    * sy
    //   pixelY_bottom = (pageHeight - word.BoundingBox.Bottom) * sy
    private static string? RenderPageToPngWithHighlights(
        string? pdfPath, int pageIndex, string[]? highlightTexts)
    {
        byte[]? rawPng = RenderPageToPngBytes(pdfPath, pageIndex);
        if (rawPng == null) return null;

        // No highlights? Return the plain PNG.
        if (highlightTexts == null || highlightTexts.Length == 0)
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

            // Draw the source page
            g.DrawImage(srcBmp, 0, 0);

            // ── Group words into logical lines (5-point Y-bucket, top-down) ───
            // Mirrors Python: y_key = round(float(w['top']) / 5) * 5
            // pdfplumber w['top']  = pageHeight - PdfPig BoundingBox.Top
            // OrderBy ensures groups are processed top-to-bottom (ascending distance from page top).
            var lineGroups = pdfPage.GetWords()
                .GroupBy(w => (int)(Math.Round((pageHeight - w.BoundingBox.Top) / 5.0) * 5))
                .OrderBy(g => g.Key)
                .ToList();

            using var fillBrush  = new SolidBrush(Color.FromArgb(90, 255, 220, 0));
            using var outlinePen = new Pen(Color.FromArgb(220, 255, 160, 0), 1.5f);

            foreach (var grp in lineGroups)
            {
                // Sort words left-to-right within the line — matches ExtractPageText ordering.
                var lineWords = grp.OrderBy(w => w.BoundingBox.Left).ToList();
                // Guard against PdfPig returning null Text, then collapse whitespace
                // so matching behaves like Python's str.split() join.
                string lineTxt = System.Text.RegularExpressions.Regex
                    .Replace(string.Join(" ", lineWords.Select(w => w.Text ?? "")), @"\s+", " ")
                    .Trim()
                    .ToLowerInvariant();
                if (lineTxt.Length == 0) continue;

                foreach (string ht in highlightTexts)
                {
                    if (string.IsNullOrWhiteSpace(ht)) continue;

                    // ── Exact Python port of _pdf_page_to_b64 matching ─────────────────
                    // Python: needle = ht.strip().lower()[:30]
                    //         if needle and (needle in line_txt or line_txt[:30] in needle)
                    // Using [:30] prevents short strings (e.g. page number "2") from
                    // matching via substring because the long full-text needle contains "2".
                    string rawNeedle = System.Text.RegularExpressions.Regex
                        .Replace(ht.Trim(), @"\s+", " ")
                        .ToLowerInvariant();
                    // Skip noise-level needles (mirrors Python: 'if needle').
                    // 4 chars is the same floor used by TextNormaliser.IsNoise.
                    if (rawNeedle.Length < 4) continue;
                    string needle      = rawNeedle.Length > 30 ? rawNeedle[..30] : rawNeedle;
                    string lineHead30  = lineTxt.Length   > 30 ? lineTxt[..30]   : lineTxt;

                    // Forward check:  the needle (first 30 chars of the highlight text)
                    //                 appears verbatim anywhere in the rendered line.
                    // Reverse check:  the rendered line is a PREFIX of the needle — i.e.
                    //                 PdfPig split the table row into a shorter bucket
                    //                 (description only, numbers in a separate bucket).
                    //                 StartsWith (not Contains) is the key false-positive fix:
                    //                 a phrase inside the needle (e.g. "before tax") cannot
                    //                 match a different line whose text does not START with it.
                    //                 Guard >= 4 still blocks 1-3 char page-stamp buckets
                    //                 like "-2-" which PdfPig splits into a bucket of just "2".
                    bool matched = lineTxt.Contains(needle) ||
                                   (lineHead30.Length >= 4 && needle.StartsWith(lineHead30));

                    if (matched)
                    {
                        float x0 = (float)(lineWords.Min(w => w.BoundingBox.Left)   * sx) - 2;
                        float x1 = (float)(lineWords.Max(w => w.BoundingBox.Right)  * sx) + 2;
                        float y0 = (float)((pageHeight - lineWords.Max(w => w.BoundingBox.Top))    * sy) - 2;
                        float y1 = (float)((pageHeight - lineWords.Min(w => w.BoundingBox.Bottom)) * sy) + 2;

                        float w  = Math.Max(1, x1 - x0);
                        float h  = Math.Max(1, y1 - y0);
                        g.FillRectangle(fillBrush,  x0, y0, w, h);
                        g.DrawRectangle(outlinePen, x0, y0, w, h);
                        break; // only one highlight per line group
                    }
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
}
