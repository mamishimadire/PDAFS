// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  CONTROLLER — ComparisonController                                           ║
// ║  Provides a traditional MVC (form-POST → result view) alternative to        ║
// ║  the existing AfsController JSON-API / single-page-app flow.               ║
// ║                                                                              ║
// ║  Route: POST /comparison/compare                                             ║
// ║  View:  Views/Comparison/Result.cshtml                                       ║
// ║                                                                              ║
// ║  NOTE: This controller is ADDITIVE — it does not replace AfsController.    ║
// ║   All existing routes, session state, and frontend JS remain untouched.    ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using AfsPdfComparison.Models;
using AfsPdfComparison.Services;
using Microsoft.AspNetCore.Mvc;

namespace AfsPdfComparison.Controllers
{
    [Route("comparison")]
    public class ComparisonController : Controller
    {
        private readonly PdfExtractionService  _extractor;
        private readonly LineComparisonService _comparator;
        private readonly PageSnapshotService   _snapshots;

        public ComparisonController(
            PdfExtractionService  extractor,
            LineComparisonService comparator,
            PageSnapshotService   snapshots)
        {
            _extractor  = extractor;
            _comparator = comparator;
            _snapshots  = snapshots;
        }

        /// <summary>
        /// GET /comparison — Show the file-upload form.
        /// </summary>
        [HttpGet("")]
        public IActionResult Index() => View();

        /// <summary>
        /// POST /comparison/compare — Accept two PDF files, compare them,
        /// render page-0 snapshots, and display the Result view.
        /// </summary>
        [HttpPost("compare")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Compare(IFormFile afs1, IFormFile afs2)
        {
            if (afs1 == null || afs2 == null)
            {
                ModelState.AddModelError("", "Both PDF files are required.");
                return View("Index");
            }

            if (!afs1.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                !afs2.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                ModelState.AddModelError("", "Only PDF files are accepted.");
                return View("Index");
            }

            // Read uploaded PDFs into byte arrays
            byte[] bytes1, bytes2;
            using (var ms = new MemoryStream())
            { await afs1.CopyToAsync(ms); bytes1 = ms.ToArray(); }
            using (var ms = new MemoryStream())
            { await afs2.CopyToAsync(ms); bytes2 = ms.ToArray(); }

            // Extract text and numbers
            var report1 = _extractor.Extract(bytes1, afs1.FileName);
            var report2 = _extractor.Extract(bytes2, afs2.FileName);

            // Run line comparison on full documents
            var diffs  = _comparator.CompareLines(report1.FullText, report2.FullText);
            var counts = _comparator.CountResults(diffs);

            // Number similarity
            var nums1 = report1.UniqueNumbers;
            var nums2 = report2.UniqueNumbers;
            var numBoth  = nums1.Intersect(nums2).ToList();
            var numOnly1 = nums1.Except(nums2).OrderBy(x => x).ToList();
            var numOnly2 = nums2.Except(nums1).OrderBy(x => x).ToList();
            double numSimilarity = (nums1.Count + nums2.Count) == 0 ? 100.0
                : Math.Round(numBoth.Count * 2.0 / (nums1.Count + nums2.Count) * 100, 1);

            // Build page snapshots for page 1 (index 0) with highlight texts
            var changedLines = diffs
                .Where(d => d.Status == DiffStatus.Changed || d.Status == DiffStatus.Removed)
                .Select(d => d.Line1).ToList();
            var addedLines = diffs
                .Where(d => d.Status == DiffStatus.Changed || d.Status == DiffStatus.Added)
                .Select(d => d.Line2).ToList();

            string snap1 = "", snap2 = "";
            try { snap1 = _snapshots.RenderPageToBase64(bytes1, 0, changedLines); } catch { }
            try { snap2 = _snapshots.RenderPageToBase64(bytes2, 0, addedLines); }  catch { }

            var vm = new ComparisonResultViewModel
            {
                Report1Filename  = report1.Filename,
                Report2Filename  = report2.Filename,
                Report1PageCount = report1.PageCount,
                Report2PageCount = report2.PageCount,
                Report1WordCount = report1.WordCount,
                Report2WordCount = report2.WordCount,
                Counts           = counts,
                Diffs            = diffs,
                NumSimilarityPct = numSimilarity,
                NumOnlyInAfs1    = numOnly1,
                NumOnlyInAfs2    = numOnly2,
                Snapshot1Base64  = snap1,
                Snapshot2Base64  = snap2,
            };
            return View("Result", vm);
        }
    }
}
