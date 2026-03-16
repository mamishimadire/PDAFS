// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3                                                               ║
// ║                                                                              ║
// ║  MODELS — Domain data structures for the entire application.                 ║
// ║  Every class here mirrors the Python STATE / report / comparison dicts.      ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • ISA 500 — Audit Evidence (IAASB)                                        ║
// ║   • SequenceMatcher algorithm (difflib) — Ratcliff/Obershelp string metric  ║
// ║   • SNG Grant Thornton CAATs methodology                                    ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text.Json.Serialization;

namespace AfsPdfComparison.Models;

// ─────────────────────────────────────────────────────────────────────────────
// VALUE-OBJECT REPLACEMENTS FOR VALUE TUPLES
// System.Text.Json cannot deserialize value tuples, so we use these simple
// classes instead. They are stored in session and serialised/deserialised.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>One token from a word-level diff (replaces the (Word, Tag) value tuple).</summary>
public class WordDiffToken
{
    public string Word { get; set; } = string.Empty;
    public string Tag  { get; set; } = string.Empty;
}

/// <summary>One entry in the page alignment map (replaces the (I1, I2, Sim) value tuple).</summary>
public class AlignmentEntry
{
    public int    I1  { get; set; }
    public int    I2  { get; set; }
    public double Sim { get; set; }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 1 · ENGAGEMENT DETAILS
// Stores all metadata entered in Step 1 of the wizard.  Persisted in the
// server-side session so it can be stamped on every exported working paper.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Engagement metadata entered by the auditor before running a comparison.
/// Maps directly to the Python <c>STATE['engagement']</c> dictionary.
/// </summary>
public class EngagementDetails
{
    /// <summary>Ultimate parent entity name (e.g. holding company).</summary>
    public string Parent      { get; set; } = string.Empty;

    /// <summary>Audit client / entity under audit.</summary>
    public string Client      { get; set; } = string.Empty;

    /// <summary>Descriptive name of the engagement (e.g. "Annual Audit FY2024").</summary>
    public string EngagementName { get; set; } = string.Empty;

    /// <summary>Unique engagement reference number assigned in the audit system.</summary>
    public string EngagementNumber { get; set; } = string.Empty;

    /// <summary>Start date of the financial year under review (e.g. "2024-01-01").</summary>
    public string FinancialYearStart { get; set; } = string.Empty;

    /// <summary>End date of the financial year under review (e.g. "2024-12-31").</summary>
    public string FinancialYearEnd   { get; set; } = DateTime.Today.Year.ToString();

    /// <summary>Name of the auditor who prepared this working paper (Mamishi Tonny Madire).</summary>
    public string PreparedBy  { get; set; } = string.Empty;

    /// <summary>Name of the engagement director responsible for sign-off.</summary>
    public string Director    { get; set; } = string.Empty;

    /// <summary>Name of the engagement manager.</summary>
    public string Manager     { get; set; } = string.Empty;

    /// <summary>
    /// ISA 500-aligned audit objective describing why the AFS comparison is performed.
    /// Pre-populated with a standard CAAT objective; the auditor may customise it.
    /// </summary>
    public string Objective   { get; set; } =
        "To independently compare two versions of the Annual Financial Statements (AFS) " +
        "through automated computer-assisted audit techniques (CAATs), identifying all " +
        "textual, numerical, and structural differences on a line-by-line and number-by-number " +
        "basis, in order to support audit evidence gathering and financial statement accuracy " +
        "assessments in accordance with ISA 500 (Audit Evidence).";
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 2 · PDF REPORT (EXTRACTED DOCUMENT)
// Produced by the PageExtractor service after reading a PDF file.
// One PdfReport object per uploaded file.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Represents an extracted PDF document.  Produced by <see cref="Services.PageExtractorService"/>.
/// Contains the full text split by page as well as metadata detected from the content.
/// Maps to the Python <c>page_extractor.extract()</c> return dict.
/// </summary>
public class PdfReport
{
    /// <summary>Human-readable label, e.g. "AFS 1" or "AFS 2".</summary>
    public string Label       { get; set; } = string.Empty;

    /// <summary>Original filename as uploaded by the user.</summary>
    public string Filename    { get; set; } = string.Empty;

    /// <summary>
    /// Extraction method: "digital" if pdfpig found selectable text,
    /// "scanned" if text came from a rasterised image (OCR not implemented in
    /// this C# version — scanned PDFs return empty pages with a warning).
    /// </summary>
    public string DocType     { get; set; } = "digital";

    /// <summary>
    /// Text extracted per page.  Index 0 = page 1.
    /// Empty string means the page contained no selectable text.
    /// </summary>
    public List<string> Pages { get; set; } = new();

    /// <summary>Concatenation of all Pages joined by newline — used for full-document diff.</summary>
    public string FullText    { get; set; } = string.Empty;

    /// <summary>Total number of pages in the PDF.</summary>
    public int PageCount      { get; set; }

    /// <summary>Four-digit financial years detected in the text (regex: 19xx / 20xx).</summary>
    public List<int> Years    { get; set; } = new();

    /// <summary>The most recent year found — used as the document's primary year.</summary>
    public int? PrimaryYear   { get; set; }

    /// <summary>
    /// Absolute server-side path to the temporary PDF stored in wwwroot/uploads.
    /// Needed to render page snapshots from the PDF raster engine.
    /// </summary>
    public string PdfPath     { get; set; } = string.Empty;

    /// <summary>
    /// Word count of FullText — matches Python notebook v4.3: len(full_text.split()).
    /// Splits on all whitespace (spaces, newlines, tabs) and counts non-empty tokens.
    /// </summary>
    [JsonIgnore] public int WordCount =>
        FullText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 3 · LINE DIFF RESULT
// Produced by the LineComparator for every line pair.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Result of comparing one line from AFS 1 against its best-matching line in AFS 2.
/// Maps to the Python <c>results</c> list items inside <c>LineComparator.compare_lines()</c>.
///
/// Status values:
///   "same"    — lines are identical (after canonical normalisation)
///   "changed" — matched but content differs
///   "added"   — line exists only in AFS 2
///   "removed" — line exists only in AFS 1
/// </summary>
public class LineDiffResult
{
    /// <summary>Comparison outcome: same | changed | added | removed.</summary>
    public string Status      { get; set; } = "same";

    /// <summary>Original line text from AFS 1 (empty for "added" lines).</summary>
    public string Line1       { get; set; } = string.Empty;

    /// <summary>Original line text from AFS 2 (empty for "removed" lines).</summary>
    public string Line2       { get; set; } = string.Empty;

    /// <summary>
    /// Ratcliff/Obershelp similarity score 0–1 after canonical normalisation.
    /// ≥ 0.94 → "same"; ≥ 0.55 → "changed"; &lt; 0.55 → no match.
    /// Reference: Python difflib.SequenceMatcher.ratio()
    /// </summary>
    public double Similarity  { get; set; }

    /// <summary>
    /// Numbers that changed between AFS 1 and AFS 2 on this line
    /// (symmetric difference of extracted numeric tokens).
    /// Empty when Status ≠ "changed".
    /// </summary>
    public List<string> NumDiff  { get; set; } = new();

    /// <summary>
    /// Word-level diff tokens for rendering coloured diffs in the UI.
    /// Each tuple is (word, "same"|"added"|"removed").
    /// </summary>
    public List<WordDiffToken> WordDiff { get; set; } = new();
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 4 · PAGE DIFF RESULT
// One per aligned page pair, produced by compare_pages_aligned.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Comparison result for one aligned page pair (AFS 1 page ↔ AFS 2 page).
/// Maps to the Python <c>page_diffs</c> list items.
///
/// A page pair is created by the Smart Page Alignment Engine which uses
/// content-fingerprint Jaccard similarity to match pages across PDFs with
/// different page counts.  See <see cref="Services.PageAlignmentService"/>.
/// </summary>
public class PageDiffResult
{
    /// <summary>Sequential pair number (1-based).</summary>
    public int  PairIndex     { get; set; }

    /// <summary>1-based page number from AFS 1 (baseline document).</summary>
    public int? PageAfs1      { get; set; }

    /// <summary>1-based page number from AFS 2 (null if no match was found).</summary>
    public int? PageAfs2      { get; set; }

    /// <summary>
    /// Content-similarity score 0–1 used by the alignment algorithm to match
    /// the two pages.  0.0 means the page is unmatched.
    /// </summary>
    public double AlignSim    { get; set; }

    /// <summary>Number of unchanged lines on this page pair.</summary>
    public int Same           { get; set; }

    /// <summary>Number of lines whose content changed (same semantic intent, different text).</summary>
    public int Changed        { get; set; }

    /// <summary>Number of lines present in AFS 2 only.</summary>
    public int Added          { get; set; }

    /// <summary>Number of lines present in AFS 1 only.</summary>
    public int Removed        { get; set; }

    /// <summary>Total lines analysed on this page pair.</summary>
    public int TotalLines     { get; set; }

    /// <summary>Percentage of unchanged lines (0–100).</summary>
    public double PctSame     { get; set; }

    /// <summary>Detailed line diffs for this page pair.</summary>
    public List<LineDiffResult> Diffs { get; set; } = new();

    /// <summary>True if AFS 1 page has no matching AFS 2 page.</summary>
    public bool Unmatched     { get; set; }

    // Convenience helpers for the Razor view
    [JsonIgnore] public int IssueCount     => Changed + Added + Removed;
    [JsonIgnore] public string StatusFlag  => IssueCount == 0 ? "OK" : IssueCount > 5 ? "!!" : "!";
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 5 · NUMBER COMPARISON RESULT
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Numeric token comparison between two full-document texts.
/// Extracted by <see cref="Services.LineComparatorService.CompareNumbers"/>.
///
/// Numbers are extracted as strings to preserve leading zeros (e.g. account codes)
/// and to allow exact equality checking even after stripping whitespace/commas.
/// Reference: Python <c>_extract_numbers()</c> / <c>compare_numbers()</c>.
/// </summary>
public class NumberComparisonResult
{
    /// <summary>Numbers present in both AFS 1 and AFS 2.</summary>
    public List<string> InBoth       { get; set; } = new();

    /// <summary>Numbers that appear in AFS 1 but not in AFS 2 (potentially deleted).</summary>
    public List<string> OnlyInAfs1   { get; set; } = new();

    /// <summary>Numbers that appear in AFS 2 but not in AFS 1 (potentially new/changed).</summary>
    public List<string> OnlyInAfs2   { get; set; } = new();

    /// <summary>Unique numeric token count in AFS 1.</summary>
    public int CountAfs1             { get; set; }

    /// <summary>Unique numeric token count in AFS 2.</summary>
    public int CountAfs2             { get; set; }

    /// <summary>
    /// Jaccard similarity of number sets as a percentage.
    /// Formula: |InBoth| / |InBoth ∪ OnlyInAfs1 ∪ OnlyInAfs2| × 100
    /// 100% means every numeric value in both documents is identical.
    /// </summary>
    public double SimilarityPct      { get; set; }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 6 · FULL COMPARISON RESULT
// The top-level object stored in session after Step 4 completes.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Master result object produced by running a full AFS comparison.
/// Stored in the server-side session under the key "Comparison".
/// Maps to the Python <c>STATE['comparison']</c> dict.
///
/// Populated by <see cref="Services.ComparisonService.BuildComparison"/>.
/// </summary>
public class ComparisonResult
{
    /// <summary>The baseline (first uploaded) PDF report.</summary>
    public PdfReport Report1          { get; set; } = null!;

    /// <summary>The comparand (second uploaded) PDF report.</summary>
    public PdfReport Report2          { get; set; } = null!;

    /// <summary>
    /// Full-document line diff — every line in AFS 2 matched to its best counterpart
    /// in AFS 1 using the inverted-index candidate search and SequenceMatcher scoring.
    /// </summary>
    public List<LineDiffResult> FullDiff { get; set; } = new();

    /// <summary>
    /// Count of line diff results grouped by status:
    /// key = "same" | "changed" | "added" | "removed".
    /// </summary>
    public Dictionary<string, int> Counts { get; set; } = new();

    /// <summary>Number comparison result across the full document text.</summary>
    public NumberComparisonResult NumCmp  { get; set; } = new();

    /// <summary>
    /// Page-level diff results — one entry per aligned page pair plus one entry per
    /// unmatched AFS 1 page.
    /// </summary>
    public List<PageDiffResult> PageDiffs { get; set; } = new();

    /// <summary>
    /// Smart alignment map: list of (afs1PageIndex, afs2PageIndex, similarityScore).
    /// afs2PageIndex = -1 means the AFS 1 page found no matching AFS 2 page.
    /// Reference: <see cref="Services.PageAlignmentService.BuildAlignment"/>.
    /// </summary>
    public List<AlignmentEntry> Alignment { get; set; } = new();

    // Convenience filtered views used by the Razor pages and export services
    // [JsonIgnore] prevents these computed views from being serialised to session
    // (the data is already in FullDiff, duplicating it would bloat session storage)
    [JsonIgnore] public IEnumerable<LineDiffResult> ChangedLines  => FullDiff.Where(d => d.Status == "changed");
    [JsonIgnore] public IEnumerable<LineDiffResult> AddedLines    => FullDiff.Where(d => d.Status == "added");
    [JsonIgnore] public IEnumerable<LineDiffResult> RemovedLines  => FullDiff.Where(d => d.Status == "removed");

    /// <summary>UTC timestamp at which the comparison was completed.</summary>
    public DateTime ComparedAt { get; set; } = DateTime.UtcNow;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 7 · AUDITOR COMMENT
// Key-value pairs saved per changed-line or per page pair.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// An auditor annotation saved against a specific diff item.
/// Keys follow the naming convention used in the Python notebook:
///   "changed:N"  — comment on changed line N
///   "page:N"     — comment on page pair N
///   "overall"    — overall conclusion comment
/// Stored in session under "Comments" as a Dictionary&lt;string, string&gt;.
/// </summary>
public class AuditorComment
{
    /// <summary>Unique key identifying which diff item this comment belongs to.</summary>
    public string Key     { get; set; } = string.Empty;

    /// <summary>Free-text comment entered by the auditor.</summary>
    public string Comment { get; set; } = string.Empty;
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 8 · VIEW MODELS (passed from controllers to Razor views)
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// View model for the main comparison dashboard (Step 4 results page).
/// Combines the comparison result with engagement context and saved comments.
/// </summary>
public class ComparisonViewModel
{
    public EngagementDetails Engagement   { get; set; } = new();
    public ComparisonResult  Comparison   { get; set; } = null!;
    public Dictionary<string, string> Comments { get; set; } = new();
    public int SelectedPairIndex          { get; set; }
}

/// <summary>
/// View model for the engagement details form (Step 1).
/// </summary>
public class EngagementViewModel
{
    public EngagementDetails Details      { get; set; } = new();
    public bool Saved                     { get; set; }
}

/// <summary>
/// View model for the upload &amp; extraction page (Steps 2 &amp; 3).
/// </summary>
public class UploadViewModel
{
    public List<string> QueuedFiles       { get; set; } = new();
    public List<PdfReport> Reports        { get; set; } = new();
    public string? ExtractionMessage      { get; set; }
}

/// <summary>
/// View model for the export page (Step 5).
/// </summary>
public class ExportViewModel
{
    public EngagementDetails Engagement   { get; set; } = new();
    public bool HasComparison             { get; set; }
    public string OutputFolder            { get; set; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "AFS_Comparison");
    public bool ExportPdf                 { get; set; } = true;
    public bool ExportWord                { get; set; } = true;
    public bool ExportExcel               { get; set; } = true;
    public bool ExportText                { get; set; } = true;
    public List<string> SavedFiles        { get; set; } = new();
    public List<string> Errors            { get; set; } = new();
}
