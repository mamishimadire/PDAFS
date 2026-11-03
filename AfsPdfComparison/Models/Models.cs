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
using AfsPdfComparison.Services;

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

/// <summary>AI-assigned highlight severity — drives the highlight colour in snapshots.</summary>
public enum AiHighlightSeverity { Normal, Medium, High }

/// <summary>Specifies a line to locate in a PDF snapshot and which words to highlight.</summary>
public class HighlightSpec
{
    /// <summary>Full line text used to locate the physical PDF Y-band.</summary>
    public string LineText { get; set; } = string.Empty;
    /// <summary>Specific changed words to highlight. Empty list = highlight entire line band.</summary>
    public List<string> ChangedWords { get; set; } = new();
    /// <summary>Stable context words used to anchor the highlight to the right visual line.</summary>
    public List<string> AnchorWords { get; set; } = new();
    /// <summary>Short phrase windows around changed words, used to paint the actual changed span.</summary>
    public List<string> ChangedPhrases { get; set; } = new();
    /// <summary>AI-driven severity: High = red, Medium = orange, Normal = yellow.</summary>
    public AiHighlightSeverity Severity { get; set; } = AiHighlightSeverity.Normal;
    /// <summary>
    /// When true the renderer uses a lower match threshold (0.45) and highlights the
    /// entire line band. Set for added/removed lines and for changed lines where only
    /// a trailing number differs (e.g. TOC page references).
    /// </summary>
    public bool WholeLineMode { get; set; } = false;
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
    /// Extraction method: "digital" if PdfPig found selectable text,
    /// or an OCR-derived variant when text was recovered from a rasterised scan.
    /// </summary>
    public string DocType     { get; set; } = string.Empty;

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

    /// <summary>
    /// Optional note explaining that this diff belongs to a continuation /
    /// cross-page match rather than a normal primary page pair.
    /// </summary>
    public string? CrossPageNote { get; set; }
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

    /// <summary>
    /// True when this is a secondary continuation/reference match rather than
    /// the primary one-to-one page alignment.
    /// </summary>
    public bool SecondaryMatch { get; set; }

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

    /// <summary>
    /// Semantic smart-comparison result built on semantic blocks and smart diffing.
    /// Stored alongside the legacy line/page diff model so existing UI paths keep working.
    /// </summary>
    public SmartComparisonResult? SmartComparison { get; set; }

    /// <summary>
    /// Document-derived section comparison built from rich page extraction and
    /// heading/section discovery from the PDF itself.
    /// </summary>
    public DocumentComparisonReport? DocumentIntelligence { get; set; }

    /// <summary>
    /// Optional Gemini-assisted review layered on top of the deterministic
    /// comparison output. This does not replace the core diff engine.
    /// </summary>
    public AiAssistResult? AiAssist { get; set; }

    /// <summary>Result from Gemini directly reading and comparing both AFS documents (content-level AI).</summary>
    public AiContentReadResult? AiContentRead { get; set; }

    /// <summary>Deterministic system-generated audit findings — no AI required.</summary>
    public List<AutoReviewFinding>   AutoReview           { get; set; } = new();

    /// <summary>Labeled financial figure changes extracted from table rows.</summary>
    public List<LabeledNumberChange> LabeledNumberChanges { get; set; } = new();

    /// <summary>#3 — Table cell-level diffs (row × col comparison).</summary>
    public List<TableCellDiff> TableCellDiffs { get; set; } = new();

    /// <summary>#4 — Prior-year comparative column mismatches (GRAP 3 / IAS 8).</summary>
    public List<ComparativeMismatch> ComparativeMismatches { get; set; } = new();

    /// <summary>#6 — Auditor-set materiality threshold in rands (default 0 = show all).</summary>
    public decimal MaterialityThreshold { get; set; } = 0;

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

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 9 · AUTO REVIEW FINDING
// Deterministic system-generated finding — no AI required.
// ─────────────────────────────────────────────────────────────────────────────
public class AutoReviewFinding
{
    public string  Id           { get; set; } = Guid.NewGuid().ToString("N")[..8];
    public string  Category     { get; set; } = "";   // numeric-change | section-added | section-removed | policy-change | director-change | sign-change | text-change | content-added | content-removed
    public string  Title        { get; set; } = "";
    public string  Description  { get; set; } = "";
    public string  Severity     { get; set; } = "medium";   // high | medium | low
    public int?    PageAfs1     { get; set; }
    public int?    PageAfs2     { get; set; }
    public int?    PairIndex    { get; set; }
    public string? Before       { get; set; }
    public string? After        { get; set; }
    public double? PctChange    { get; set; }
    public decimal? AbsChange   { get; set; }
    // #7 — Confidence 0–100: how likely this is a genuine change (not a layout artefact)
    public int     Confidence   { get; set; } = 80;
    // #1 — Section group (e.g. "Note 12 — Provisions", "Statement of Financial Position")
    public string  SectionGroup { get; set; } = "Other";
    // #5 — Auditor sign-off: "" | "cleared" | "query" | "escalate"
    public string  SignOff      { get; set; } = "";
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 10 · LABELED NUMBER CHANGE
// A financial figure with its label and the computed change.
// ─────────────────────────────────────────────────────────────────────────────
// SECTION 10B · TABLE CELL DIFF  (#3 — table-aware comparison)
// ─────────────────────────────────────────────────────────────────────────────
public class TableCellDiff
{
    public int     Row        { get; set; }
    public int     Col        { get; set; }
    public string  Label      { get; set; } = "";  // row header label
    public string  ColHeader  { get; set; } = "";  // column header
    public string  ValueAfs1  { get; set; } = "";
    public string  ValueAfs2  { get; set; } = "";
    public string  Status     { get; set; } = "same";  // same | changed | added | removed
    public double? PctChange  { get; set; }
    public int?    PageAfs1   { get; set; }
    public int?    PageAfs2   { get; set; }
    public int?    PairIndex  { get; set; }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 10C · COMPARATIVE COLUMN MISMATCH  (#4 — year-over-year validation)
// ─────────────────────────────────────────────────────────────────────────────
public class ComparativeMismatch
{
    public string  Label       { get; set; } = "";
    public string  ValueAfs1   { get; set; } = "";   // current-year value in AFS 1
    public string  ValueAfs2   { get; set; } = "";   // prior-year comparative in AFS 2 (should match ValueAfs1)
    public double? PctDiff     { get; set; }
    public string  Severity    { get; set; } = "high";
    public int?    PageAfs1    { get; set; }
    public int?    PageAfs2    { get; set; }
    public int?    PairIndex   { get; set; }
}

// ─────────────────────────────────────────────────────────────────────────────
public class LabeledNumberChange
{
    public string  Label       { get; set; } = "";
    public string  RawBefore   { get; set; } = "";
    public string  RawAfter    { get; set; } = "";
    public decimal Before      { get; set; }
    public decimal After       { get; set; }
    public double  PctChange   { get; set; }
    public decimal AbsChange   { get; set; }
    public string  Severity    { get; set; } = "medium";
    public int?    PageAfs1    { get; set; }
    public int?    PageAfs2    { get; set; }
    public int?    PairIndex   { get; set; }
    public bool    SignChanged  { get; set; }   // positive→negative or vice versa
}
