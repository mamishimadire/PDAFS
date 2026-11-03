// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  SERVICE — WorkingPaperExportService  v4.3.2                                 ║
// ║  Generates a professional Word (.docx) audit working paper that mirrors      ║
// ║  the Python notebook output exactly.                                         ║
// ║                                                                              ║
// ║  Structure (8 sections):                                                     ║
// ║    Cover page — SNG template, engagement details, audit objective            ║
// ║    TOC                                                                       ║
// ║    §1  Documents compared                                                    ║
// ║    §2  Page-by-page snapshot review (side-by-side with yellow highlights)    ║
// ║    §3  Comparison summary                                                    ║
// ║    §4  Changed lines with word diff                                          ║
// ║    §5  Added lines                                                           ║
// ║    §6  Removed lines                                                         ║
// ║    §7  Number comparison                                                     ║
// ║    §8  Consolidated auditor comments + overall sign-off box                  ║
// ║                                                                              ║
// ║  v4.3.2 FIX: 'users' hyperlink false-highlight eliminated via               ║
// ║    TextNormaliser NFKD + \p{C} stripping (invisible Unicode chars).          ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text.RegularExpressions;
using AfsPdfComparison.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SkiaSharp;

namespace AfsPdfComparison.Services;

// ─────────────────────────────────────────────────────────────────────────────
// SECTION WP · WORKING PAPER EXPORT SERVICE
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Builds a complete Word working paper from a ComparisonResult.
/// Registered as scoped in DI — one instance per request.
/// </summary>
public class WorkingPaperExportService
{
    private readonly PageSnapshotService _snapshots;
    private uint _nextDrawingId = 1;   // unique per-document drawing ID (Word requires unique IDs)

    // ── SNG brand colours ────────────────────────────────────────────────────
    private const string C_PURPLE  = "5C2D91";
    private const string C_DPURPLE = "3D1A6E";
    private const string C_NAVY    = "0F1C3F";
    private const string C_DARK    = "1E2F5A";
    private const string C_AMBER   = "F59E0B";
    private const string C_GREEN   = "10B981";
    private const string C_RED     = "EF4444";
    private const string C_WHITE   = "FFFFFF";
    private const string C_LTGRAY  = "F1F5F9";
    private const string C_CMT_BG  = "FFFBEB";

    // ── A4 content width in DXA (17 cm = 9638 DXA) ──────────────────────────
    private const int CONTENT_DXA = 9638;

    public WorkingPaperExportService(PageSnapshotService snapshots)
    {
        _snapshots = snapshots;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // PUBLIC ENTRY POINT
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Builds the working paper and writes it to <paramref name="outputPath"/>.
    /// </summary>
    public async Task Build(
        ComparisonResult            cmp,
        EngagementDetails           eng,
        Dictionary<string, string>  comments,
        byte[]?                     pdf1Bytes,
        byte[]?                     pdf2Bytes,
        string                      outputPath)
    {
        string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        using var wordDoc = WordprocessingDocument.Create(
            outputPath, WordprocessingDocumentType.Document);
        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        // ── Document settings (compatibility mode 15 = Word 2013+) ───────────
        // REQUIRED: without this Word 2016+ may show repair prompts.
        var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
        settingsPart.Settings = new Settings(
            new Compatibility(
                new CompatibilitySetting
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri  = new StringValue("http://schemas.microsoft.com/office/word"),
                    Val  = new StringValue("15"),
                }));
        settingsPart.Settings.Save();

        // NOTE: SectionProperties MUST be the last element in Body per OOXML spec.
        // It is added at the very end of this method, just before Document.Save().

        // ═════════════════════════════════════════════════════════════════════
        // COVER PAGE
        // ═════════════════════════════════════════════════════════════════════
        AddCoverPage(body, eng, ts);
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // TABLE OF CONTENTS
        // ═════════════════════════════════════════════════════════════════════
        AddH1(body, "Table of Contents");
        HRule(body);
        string[] tocNums   = { "1","2","3","4","5","6","7","8","9","  9.1","  9.2","  9.3","  9.4","  9.5" };
        string[] tocTitles = {
            "Documents Compared",
            "Page-by-Page Snapshot Review",
            "Comparison Summary",
            "Changed Lines",
            "Added Lines",
            "Removed Lines",
            "Number Comparison",
            "Consolidated Auditor Comments",
            "Materiality Review Findings",
            "Overall Assessment",
            "Section-by-Section Changes",
            "Line-Item Financial Changes",
            "Audit Conclusion",
            "System Review Findings (Deterministic)",
        };
        for (int i = 0; i < tocNums.Length; i++)
        {
            var p = body.AppendChild(new Paragraph());
            p.AppendChild(MakeRun(tocNums[i] + "  ", bold: true, colorHex: C_PURPLE, halfPt: 20));
            p.AppendChild(MakeRun(tocTitles[i], halfPt: 20));
        }
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §1  DOCUMENTS COMPARED
        // ═════════════════════════════════════════════════════════════════════
        AddH1(body, "1  Documents Compared");
        HRule(body);
        var r1 = cmp.Report1; var r2 = cmp.Report2;
        AddTable(body,
            headers: new[] { "", "Label", "Filename", "Type", "Year", "Pages", "Words" },
            rows: new[] {
                new[] { "AFS 1", r1.Label, r1.Filename, r1.DocType,
                        r1.PrimaryYear?.ToString() ?? "?",
                        r1.PageCount.ToString(), r1.WordCount.ToString() },
                new[] { "AFS 2", r2.Label, r2.Filename, r2.DocType,
                        r2.PrimaryYear?.ToString() ?? "?",
                        r2.PageCount.ToString(), r2.WordCount.ToString() },
            },
            headerFill: C_DPURPLE,
            colWidths: new[] { 680, 1020, 3686, 907, 794, 794, 757 }
        );
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §2  PAGE-BY-PAGE SNAPSHOT REVIEW
        // ═════════════════════════════════════════════════════════════════════
        AddH1(body, "2  Page-by-Page Snapshot Review");
        HRule(body);
        AddPara(body,
            "Each page pair shows AFS 1 and AFS 2 side by side with yellow highlights on " +
            "changed lines. Pages are matched by content similarity, not page number. " +
            "The Auditor Note box after each snapshot is for findings.",
            halfPt: 18);
        body.AppendChild(new Paragraph());

        var paired = cmp.PageDiffs.Where(pd =>
            pd.PageAfs1.HasValue &&
            pd.PageAfs2.HasValue &&
            !pd.Unmatched &&
            !pd.SecondaryMatch).ToList();

        for (int k = 0; k < paired.Count; k++)
        {
            var    pd       = paired[k];
            int    p1i      = pd.PageAfs1!.Value - 1;
            int    p2i      = pd.PageAfs2!.Value - 1;
            int    issues   = pd.Changed + pd.Added + pd.Removed;
            string pairKey  = "page:" + k;
            string pairLbl  = $"AFS1 p{pd.PageAfs1} \u2194 AFS2 p{pd.PageAfs2}";
            string issuesTxt = issues == 0 ? "NO CHANGES" : $"{issues} CHANGE(S)";
            string statsText =
                $"Pair {k + 1}   {pairLbl}   |  Align: {pd.AlignSim:F3}" +
                $"   |  {pd.PctSame}% same" +
                $"   \u2713{pd.Same}  ~{pd.Changed}  +{pd.Added}  -{pd.Removed}" +
                $"   \u2190 {issuesTxt}";

            var statsP = body.AppendChild(new Paragraph());
            SetParaFill(statsP, C_DARK);
            statsP.AppendChild(MakeRun(statsText, bold: true, colorHex: C_WHITE, halfPt: 18));

            var hl1Raw = pd.Diffs
                .Where(d => (d.Status == "changed" || d.Status == "removed")
                         && !string.IsNullOrWhiteSpace(d.Line1))
                .Select(d => d.Line1).ToList();
            var hl2Raw = pd.Diffs
                .Where(d => (d.Status == "changed" || d.Status == "added")
                         && !string.IsNullOrWhiteSpace(d.Line2))
                .Select(d => d.Line2).ToList();

            string? img1b64 = null, img2b64 = null;
            try { if (pdf1Bytes != null) img1b64 = await _snapshots.RenderPageToBase64(pdf1Bytes, p1i, hl1Raw); }
            catch { /* best-effort */ }
            try { if (pdf2Bytes != null) img2b64 = await _snapshots.RenderPageToBase64(pdf2Bytes, p2i, hl2Raw); }
            catch { /* best-effort */ }

            AddImageRow(mainPart, body, img1b64, img2b64,
                $"AFS 1 \u2014 page {pd.PageAfs1}", "DBEAFE",
                $"AFS 2 \u2014 page {pd.PageAfs2}", "EDE9FE");

            // ▲ legend immediately below the snapshot
            var legP = body.AppendChild(new Paragraph());
            var legPPr = legP.AppendChild(new ParagraphProperties());
            legPPr.AppendChild(new SpacingBetweenLines { Before = "40", After = "40" });
            legP.AppendChild(MakeRun("\u25b2 Yellow highlights = changed / added / removed lines",
                halfPt: 14, colorHex: "D97706"));

            // ✏ AUDITOR COMMENT BOX — directly below snapshot
            AddCommentBox(body, pairKey, pairLbl, comments);

            // ── Line diff table (only if this page pair has changes) ──────────
            var diffRows = pd.Diffs
                .Where(d => d.Status is "changed" or "added" or "removed").ToList();

            if (diffRows.Count > 0)
            {
                var chHdr = body.AppendChild(new Paragraph());
                var chPPr = chHdr.AppendChild(new ParagraphProperties());
                chPPr.AppendChild(new SpacingBetweenLines { Before = "120", After = "60" });
                chHdr.AppendChild(MakeRun($"Section-by-Section Changes — {diffRows.Count} line-level difference(s) detected on this page pair:",
                    bold: true, halfPt: 16, colorHex: C_DARK));

                AddTable(body,
                    headers: new[] { "", "AFS 1 line  \u2192  AFS 2 line", "Num diff" },
                    rows: diffRows.Select((d, j) =>
                    {
                        string sym = (d.Status == "changed" ? "~"
                                    : d.Status == "added"   ? "+" : "-") + (j + 1);
                        string lineText = d.Status == "changed"
                            ? Trunc(d.Line1, 60) + "  →  " + Trunc(d.Line2, 60)
                            : d.Status == "added" ? "+  " + Trunc(d.Line2, 120)
                                                  : "−  " + Trunc(d.Line1, 120);
                        string numText = d.NumDiff.Count > 0
                            ? string.Join(", ", d.NumDiff) : "\u2014";
                        return new[] { sym, lineText, numText };
                    }).ToArray(),
                    headerFill: C_DARK,
                    colWidths: new[] { 567, 7933, 1138 }
                );
            }
            else if (issues > 0)
            {
                var chHdr2 = body.AppendChild(new Paragraph());
                chHdr2.AppendChild(MakeRun(
                    $"Note: {issues} change(s) detected on this page pair (see snapshot highlights above).",
                    halfPt: 16, colorHex: C_DARK, italic: true));
            }

            // ── Review Findings — always shown below every snapshot ──────────
            {
                bool readReady = string.Equals(
                    cmp.AiContentRead?.Status, "ready", StringComparison.OrdinalIgnoreCase);
                bool aiReady = string.Equals(
                    cmp.AiAssist?.Status, "ready", StringComparison.OrdinalIgnoreCase);

                // Source 1: Document Reading section diffs (semantic analysis)
                // Match by PairIndex first; fall back to matching by page numbers
                var pairSections = readReady
                    ? (cmp.AiContentRead!.SectionDiffs ?? new())
                        .Where(s =>
                            (s.PairIndex.HasValue && s.PairIndex == pd.PairIndex) ||
                            (!s.PairIndex.HasValue &&
                             ((s.PageAfs1.HasValue && s.PageAfs1 == pd.PageAfs1) ||
                              (s.PageAfs2.HasValue && s.PageAfs2 == pd.PageAfs2))))
                        .ToList()
                    : new();

                // Material findings — match by PairIndex OR page numbers as robust fallback
                var pairMaterialFindings = aiReady
                    ? (cmp.AiAssist!.MaterialFindings ?? new())
                        .Where(f =>
                            f.PairIndex == pd.PairIndex ||
                            (f.PageAfs1.HasValue && f.PageAfs1 == pd.PageAfs1) ||
                            (f.PageAfs2.HasValue && f.PageAfs2 == pd.PageAfs2))
                        .DistinctBy(f => f.Title)
                        .ToList()
                    : new();

                var pairFocus = aiReady
                    ? (cmp.AiAssist!.PageFocus ?? new())
                        .Where(f =>
                            f.PairIndex == pd.PairIndex ||
                            (f.PageAfs1.HasValue && f.PageAfs1 == pd.PageAfs1))
                        .ToList()
                    : new();

                bool hasAiFindings = pairSections.Count > 0 || pairMaterialFindings.Count > 0;

                // System-detected diffs — these are the exact lines highlighted in the snapshots
                var issueDiffsForPair = pd.Diffs
                    .Where(d => d.Status != "same" &&
                                d.Status != "same-cross-page" &&
                                d.Status != "same-different-page")
                    .ToList();
                var changedDiffs = issueDiffsForPair.Where(d => d.Status == "changed").ToList();
                var addedDiffs   = issueDiffsForPair.Where(d => d.Status == "added").ToList();
                var removedDiffs = issueDiffsForPair.Where(d => d.Status == "removed").ToList();
                bool hasSystemEvidence = issueDiffsForPair.Count > 0;

                // Header bar — always rendered
                var rvHdr = body.AppendChild(new Paragraph());
                var rvPPr = rvHdr.AppendChild(new ParagraphProperties());
                rvPPr.AppendChild(new SpacingBetweenLines { Before = "120", After = "40" });
                SetParaFill(rvHdr, C_DPURPLE);
                rvHdr.AppendChild(MakeRun(
                    $"●  Review Findings — {pairLbl}",
                    bold: true, colorHex: C_WHITE, halfPt: 18));

                if (!hasSystemEvidence && issues == 0)
                {
                    // Truly no changes on this page pair
                    var okP = body.AppendChild(new Paragraph());
                    SetParaFill(okP, C_GREEN);
                    okP.AppendChild(MakeRun(
                        "✓  No changes identified for this page pair.",
                        bold: true, colorHex: C_WHITE, halfPt: 16));
                }
                else
                {
                    // ── PART A: System-Detected Differences ──────────────────────
                    // Each row corresponds 1-to-1 with a highlighted line in the
                    // snapshot images above and with the Changed Lines, Added Lines,
                    // and Removed Lines tabs shown in the system UI. This is the
                    // factual, machine-generated record of what changed.
                    if (hasSystemEvidence)
                    {
                        var sysHdrP = body.AppendChild(new Paragraph());
                        var sysHdrPPr = sysHdrP.AppendChild(new ParagraphProperties());
                        sysHdrPPr.AppendChild(new SpacingBetweenLines { Before = "80", After = "20" });
                        SetParaFill(sysHdrP, C_DARK);
                        sysHdrP.AppendChild(MakeRun(
                            $"A  System-Detected Differences   " +
                            $"[~{changedDiffs.Count} changed   +{addedDiffs.Count} added   −{removedDiffs.Count} removed]",
                            bold: true, colorHex: C_WHITE, halfPt: 17));
                        AddPara(body,
                            "Each row below corresponds to a highlighted line in the snapshot images above. " +
                            "Changed-line rows show which specific words were altered (word diff).",
                            halfPt: 14);

                        if (changedDiffs.Count > 0)
                        {
                            var chSubP = body.AppendChild(new Paragraph());
                            var chSubPPr = chSubP.AppendChild(new ParagraphProperties());
                            chSubPPr.AppendChild(new SpacingBetweenLines { Before = "60", After = "20" });
                            chSubP.AppendChild(MakeRun(
                                $"~ Changed Lines ({changedDiffs.Count}) — word diff: −\"removed\" / +\"added\":",
                                bold: true, halfPt: 15, colorHex: "D97706"));
                            AddTable(body,
                                headers: new[] { "#", "AFS 1 — Before", "AFS 2 — After", "Changed Words", "Num Diff" },
                                rows: changedDiffs.Select((d, j) => new[]
                                {
                                    $"~{j + 1}",
                                    Trunc(d.Line1 ?? "", 130),
                                    Trunc(d.Line2 ?? "", 130),
                                    FormatWordDiff(d.WordDiff),
                                    d.NumDiff.Count > 0 ? string.Join(" | ", d.NumDiff.Take(6)) : "—",
                                }).ToArray(),
                                headerFill: "92400E",
                                colWidths: new[] { 454, 2757, 2757, 2484, 1186 });
                        }

                        if (addedDiffs.Count > 0)
                        {
                            var addSubP = body.AppendChild(new Paragraph());
                            var addSubPPr = addSubP.AppendChild(new ParagraphProperties());
                            addSubPPr.AppendChild(new SpacingBetweenLines { Before = "60", After = "20" });
                            addSubP.AppendChild(MakeRun(
                                $"+ Lines Added in AFS 2 ({addedDiffs.Count}) — present in AFS 2 but absent from AFS 1:",
                                bold: true, halfPt: 15, colorHex: "065F46"));
                            AddTable(body,
                                headers: new[] { "#", "Content Added in AFS 2" },
                                rows: addedDiffs.Select((d, j) => new[]
                                {
                                    $"+{j + 1}",
                                    Trunc(d.Line2 ?? "", 250),
                                }).ToArray(),
                                headerFill: C_GREEN,
                                colWidths: new[] { 454, CONTENT_DXA - 454 });
                        }

                        if (removedDiffs.Count > 0)
                        {
                            var remSubP = body.AppendChild(new Paragraph());
                            var remSubPPr = remSubP.AppendChild(new ParagraphProperties());
                            remSubPPr.AppendChild(new SpacingBetweenLines { Before = "60", After = "20" });
                            remSubP.AppendChild(MakeRun(
                                $"− Lines Removed from AFS 1 ({removedDiffs.Count}) — present in AFS 1 but absent from AFS 2:",
                                bold: true, halfPt: 15, colorHex: "991B1B"));
                            AddTable(body,
                                headers: new[] { "#", "Content Removed from AFS 1" },
                                rows: removedDiffs.Select((d, j) => new[]
                                {
                                    $"-{j + 1}",
                                    Trunc(d.Line1 ?? "", 250),
                                }).ToArray(),
                                headerFill: C_RED,
                                colWidths: new[] { 454, CONTENT_DXA - 454 });
                        }

                        body.AppendChild(new Paragraph());
                    }

                    // ── PART B: System Review Findings (AutoReview — always available) ──
                    var sysReviewFindings = (cmp.AutoReview ?? new())
                        .Where(f => f.PairIndex == pd.PairIndex)
                        .OrderBy(f => f.Severity == "high" ? 0 : f.Severity == "medium" ? 1 : 2)
                        .ToList();

                    if (sysReviewFindings.Count > 0)
                    {
                        var srHdrP = body.AppendChild(new Paragraph());
                        var srHdrPPr = srHdrP.AppendChild(new ParagraphProperties());
                        srHdrPPr.AppendChild(new SpacingBetweenLines { Before = "80", After = "20" });
                        SetParaFill(srHdrP, C_DPURPLE);
                        srHdrP.AppendChild(MakeRun(
                            $"B  System Review Findings  ({sysReviewFindings.Count} finding(s) — deterministic, no AI required)",
                            bold: true, colorHex: C_WHITE, halfPt: 17));

                        foreach (var sf in sysReviewFindings)
                        {
                            string sfc = (sf.Severity?.ToLowerInvariant() ?? "low") switch
                            { "high" => C_RED, "medium" => C_AMBER, _ => C_GREEN };

                            var sfHdr = body.AppendChild(new Paragraph());
                            var sfPPr = sfHdr.AppendChild(new ParagraphProperties());
                            sfPPr.AppendChild(new SpacingBetweenLines { Before = "60", After = "20" });
                            SetParaFill(sfHdr, sfc);
                            sfHdr.AppendChild(MakeRun(
                                $"  [{sf.Severity?.ToUpperInvariant() ?? "LOW"}]  {sf.Title}",
                                bold: true, colorHex: C_WHITE, halfPt: 17));

                            var sfMeta = body.AppendChild(new Paragraph());
                            sfMeta.AppendChild(MakeRun(
                                $"Category: {sf.Category}   |   Section: {sf.SectionGroup}   |   " +
                                $"AFS 1 p{sf.PageAfs1} ↔ AFS 2 p{sf.PageAfs2}   |   Confidence: {sf.Confidence}%",
                                halfPt: 13, colorHex: sfc, bold: true));

                            if (!string.IsNullOrWhiteSpace(sf.Description))
                                AddPara(body, sf.Description, halfPt: 15);

                            if (!string.IsNullOrWhiteSpace(sf.Before) || !string.IsNullOrWhiteSpace(sf.After))
                            {
                                AddTable(body,
                                    headers: new[] { "Before (AFS 1)", "After (AFS 2)" },
                                    rows: new[] { new[] { Trunc(sf.Before ?? "—", 200), Trunc(sf.After ?? "—", 200) } },
                                    headerFill: "374151",
                                    colWidths: new[] { CONTENT_DXA / 2, CONTENT_DXA / 2 });
                            }

                            if (sf.PctChange.HasValue && sf.PctChange > 0)
                            {
                                var changeP = body.AppendChild(new Paragraph());
                                changeP.AppendChild(MakeRun(
                                    $"Change: {sf.PctChange:F1}%{(sf.AbsChange.HasValue ? $"   |   Absolute: R {sf.AbsChange:N0}" : "")}",
                                    halfPt: 13, colorHex: C_AMBER, bold: true));
                            }

                            body.AppendChild(new Paragraph());
                        }
                    }
                    else if (!hasAiFindings && hasSystemEvidence)
                    {
                        var aiNaP = body.AppendChild(new Paragraph());
                        aiNaP.AppendChild(MakeRun(
                            "System review did not identify curated findings for this page pair. " +
                            "The raw differences in Part A above require independent auditor review.",
                            halfPt: 14, colorHex: "6B7280", italic: true));
                    }

                    // ── PART C: AI Professional Analysis (when available) ─────────
                    if (hasAiFindings)
                    {
                        var aiHdrP = body.AppendChild(new Paragraph());
                        var aiHdrPPr = aiHdrP.AppendChild(new ParagraphProperties());
                        aiHdrPPr.AppendChild(new SpacingBetweenLines { Before = "80", After = "20" });
                        SetParaFill(aiHdrP, C_NAVY);
                        aiHdrP.AppendChild(MakeRun(
                            "C  AI Professional Analysis",
                            bold: true, colorHex: C_WHITE, halfPt: 17));
                    }

                    if (hasAiFindings)
                    foreach (var mf in pairMaterialFindings)
                    {
                        string mfc = (mf.Severity?.ToLowerInvariant() ?? "low") switch
                        { "high" => C_RED, "medium" => C_AMBER, _ => C_GREEN };

                        var mfHdr = body.AppendChild(new Paragraph());
                        var mfPPr = mfHdr.AppendChild(new ParagraphProperties());
                        mfPPr.AppendChild(new SpacingBetweenLines { Before = "80", After = "20" });
                        SetParaFill(mfHdr, mfc);
                        mfHdr.AppendChild(MakeRun(
                            $"  {mf.Title}",
                            bold: true, colorHex: C_WHITE, halfPt: 18));

                        var mfChip = body.AppendChild(new Paragraph());
                        mfChip.AppendChild(MakeRun(
                            $"Severity: {mf.Severity?.ToUpperInvariant() ?? "LOW"}   |   " +
                            $"AFS 1 p{mf.PageAfs1 ?? pd.PageAfs1} ↔ AFS 2 p{mf.PageAfs2 ?? pd.PageAfs2}",
                            halfPt: 14, colorHex: mfc, bold: true));

                        if (!string.IsNullOrWhiteSpace(mf.AuditRationale))
                            AddPara(body, mf.AuditRationale, halfPt: 16);

                        if (!string.IsNullOrWhiteSpace(mf.RecommendedProcedure))
                        {
                            var rpP = body.AppendChild(new Paragraph());
                            rpP.AppendChild(MakeRun(
                                "Recommended Procedure:  ", bold: true, halfPt: 14, colorHex: C_DARK));
                            rpP.AppendChild(MakeRun(mf.RecommendedProcedure, halfPt: 14));
                        }

                        if (mf.ChangedNumbers?.Count > 0)
                        {
                            var cnP = body.AppendChild(new Paragraph());
                            cnP.AppendChild(MakeRun("Changed numbers:  ", bold: true, halfPt: 14));
                            cnP.AppendChild(MakeRun(
                                string.Join("   |   ", mf.ChangedNumbers),
                                halfPt: 14, colorHex: C_AMBER));
                        }

                        // Show the specific system-generated diff lines cited by the AI
                        if (mf.FocusDiffOrdinals?.Count > 0 && issueDiffsForPair.Count > 0)
                        {
                            var citedDiffs = mf.FocusDiffOrdinals
                                .Where(o => o > 0 && o <= issueDiffsForPair.Count)
                                .Select(o => (ordinal: o, diff: issueDiffsForPair[o - 1]))
                                .Take(6)
                                .ToList();

                            if (citedDiffs.Count > 0)
                            {
                                var evHdr = body.AppendChild(new Paragraph());
                                var evPPr = evHdr.AppendChild(new ParagraphProperties());
                                evPPr.AppendChild(new SpacingBetweenLines { Before = "60", After = "20" });
                                evHdr.AppendChild(MakeRun(
                                    "System Evidence — diff lines cited by AI:",
                                    bold: true, halfPt: 14, colorHex: C_DARK));

                                AddTable(body,
                                    headers: new[] { "#", "AFS 1 Text", "AFS 2 Text", "Type" },
                                    rows: citedDiffs.Select(x => new[]
                                    {
                                        x.ordinal.ToString(),
                                        Trunc(string.IsNullOrEmpty(x.diff.Line1) ? "—" : x.diff.Line1, 80),
                                        Trunc(string.IsNullOrEmpty(x.diff.Line2) ? "—" : x.diff.Line2, 80),
                                        x.diff.Status switch
                                        {
                                            "changed" => "~ Changed",
                                            "added"   => "+ Added",
                                            "removed" => "− Removed",
                                            _         => x.diff.Status,
                                        },
                                    }).ToArray(),
                                    headerFill: C_DARK,
                                    colWidths: new[] { 454, 3600, 3600, 984 });
                            }
                        }

                        body.AppendChild(new Paragraph());
                    }

                    // ── Source 2: Document Reading section diffs (semantic / narrative analysis)
                    foreach (var s in pairSections)
                    {
                        string sc = (s.Severity?.ToLowerInvariant() ?? "low") switch
                        { "high" => C_RED, "medium" => C_AMBER, _ => C_GREEN };

                        var sHdr = body.AppendChild(new Paragraph());
                        SetParaFill(sHdr, sc);
                        sHdr.AppendChild(MakeRun(
                            $"{s.SectionName}",
                            bold: true, colorHex: C_WHITE, halfPt: 18));

                        var chipP = body.AppendChild(new Paragraph());
                        chipP.AppendChild(MakeRun(
                            $"Change type: {s.ChangeType}   |   Severity: {s.Severity?.ToUpperInvariant()}",
                            halfPt: 14, colorHex: sc, bold: true));

                        if (!string.IsNullOrWhiteSpace(s.ChangeDescription))
                            AddPara(body, s.ChangeDescription, halfPt: 16);

                        if (!string.IsNullOrWhiteSpace(s.Afs1Summary) || !string.IsNullOrWhiteSpace(s.Afs2Summary))
                        {
                            AddTable(body,
                                headers: new[] { "AFS 1 — Before", "AFS 2 — After" },
                                rows: new[] { new[] {
                                    Trunc(s.Afs1Summary, 120),
                                    Trunc(s.Afs2Summary, 120),
                                }},
                                headerFill: C_DARK,
                                colWidths: new[] { CONTENT_DXA / 2, CONTENT_DXA / 2 });
                        }
                        if (s.ChangedValues?.Count > 0)
                        {
                            var cvP = body.AppendChild(new Paragraph());
                            cvP.AppendChild(MakeRun("Changed values:  ", bold: true, halfPt: 14));
                            cvP.AppendChild(MakeRun(string.Join("   |   ", s.ChangedValues), halfPt: 14, colorHex: C_AMBER));
                        }
                        body.AppendChild(new Paragraph());
                    }

                    // ── Source 3: Priority notes from Diff Analysis (supplementary)
                    foreach (var pf in pairFocus.Where(f => !string.IsNullOrWhiteSpace(f.Why)))
                    {
                        var pfP = body.AppendChild(new Paragraph());
                        pfP.AppendChild(MakeRun(
                            $"▸  Auditor Priority [{pf.Priority?.ToUpperInvariant()}]: {pf.Why}",
                            halfPt: 16, colorHex: C_AMBER));
                    }
                }
                body.AppendChild(new Paragraph());
            }

            if (k < paired.Count - 1)
                PageBreak(body);
        }
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §3  COMPARISON SUMMARY
        // ═════════════════════════════════════════════════════════════════════
        AddH1(body, "3  Comparison Summary");
        HRule(body);
        int total = Math.Max(cmp.Counts.Values.Sum(), 1);
        int csm   = cmp.Counts.GetValueOrDefault("same",    0);
        int cch   = cmp.Counts.GetValueOrDefault("changed", 0);
        int cad   = cmp.Counts.GetValueOrDefault("added",   0);
        int crm   = cmp.Counts.GetValueOrDefault("removed", 0);
        AddTable(body,
            headers: new[] { "Metric", "Count", "% of Total" },
            rows: new[] {
                new[] { "Same Lines",           csm.ToString(), Pct(csm, total) },
                new[] { "Changed Lines",        cch.ToString(), Pct(cch, total) },
                new[] { "Added Lines",          cad.ToString(), Pct(cad, total) },
                new[] { "Removed Lines",        crm.ToString(), Pct(crm, total) },
                new[] { "Number Similarity",    cmp.NumCmp.SimilarityPct + "%", "\u2014" },
                new[] { "Numbers only in AFS1", cmp.NumCmp.OnlyInAfs1.Count.ToString(), "\u2014" },
                new[] { "Numbers only in AFS2", cmp.NumCmp.OnlyInAfs2.Count.ToString(), "\u2014" },
            },
            headerFill: C_NAVY
        );
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §4  CHANGED LINES
        // ═════════════════════════════════════════════════════════════════════
        var changedList = cmp.ChangedLines.ToList();
        AddH1(body, $"4  Changed Lines  ({changedList.Count})");
        HRule(body);
        if (changedList.Count > 0)
            AddTable(body,
                headers: new[] { "#", "AFS 1 Line", "AFS 2 Line", "Num Diff" },
                rows: changedList.Take(200).Select((d, i) => new[] {
                    (i + 1).ToString(),
                    Trunc(d.Line1, 80),
                    Trunc(d.Line2, 80),
                    d.NumDiff.Count > 0 ? string.Join(", ", d.NumDiff) : "\u2014",
                }).ToArray(),
                headerFill: "1C1400",
                colWidths: new[] { 454, 3969, 3969, 1246 }
            );
        else
            AddPara(body, "No changed lines detected.");
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §5  ADDED LINES
        // ═════════════════════════════════════════════════════════════════════
        var addedList = cmp.AddedLines.ToList();
        AddH1(body, $"5  Added Lines  ({addedList.Count} in AFS 2)");
        HRule(body);
        if (addedList.Count > 0)
            AddTable(body,
                headers: new[] { "#", "Added Line (AFS 2)" },
                rows: addedList.Take(200).Select((d, i) => new[] {
                    (i + 1).ToString(), Trunc(d.Line2, 120),
                }).ToArray(),
                headerFill: "1E3A5F"
            );
        else
            AddPara(body, "No added lines.");
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §6  REMOVED LINES
        // ═════════════════════════════════════════════════════════════════════
        var removedList = cmp.RemovedLines.ToList();
        AddH1(body, $"6  Removed Lines  ({removedList.Count} from AFS 1)");
        HRule(body);
        if (removedList.Count > 0)
            AddTable(body,
                headers: new[] { "#", "Removed Line (AFS 1)" },
                rows: removedList.Take(200).Select((d, i) => new[] {
                    (i + 1).ToString(), Trunc(d.Line1, 120),
                }).ToArray(),
                headerFill: "5F1D1D"
            );
        else
            AddPara(body, "No removed lines.");
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §7  NUMBER COMPARISON
        // ═════════════════════════════════════════════════════════════════════
        AddH1(body, "7  Number Comparison");
        HRule(body);
        var nc = cmp.NumCmp;
        AddTable(body,
            headers: new[] { "Metric", "Value" },
            rows: new[] {
                new[] { "AFS 1 unique numbers", nc.CountAfs1.ToString() },
                new[] { "AFS 2 unique numbers", nc.CountAfs2.ToString() },
                new[] { "Numbers in both",      nc.InBoth.Count.ToString() },
                new[] { "Only in AFS 1",        nc.OnlyInAfs1.Count.ToString() },
                new[] { "Only in AFS 2",        nc.OnlyInAfs2.Count.ToString() },
                new[] { "Similarity",           nc.SimilarityPct + "%" },
            },
            headerFill: "3B0066",
            colWidths: new[] { 4536, 5102 }
        );
        if (nc.OnlyInAfs1.Count > 0)
            AddPara(body, "Only in AFS 1: " + string.Join(", ", nc.OnlyInAfs1.Take(50)), halfPt: 16);
        if (nc.OnlyInAfs2.Count > 0)
            AddPara(body, "Only in AFS 2: " + string.Join(", ", nc.OnlyInAfs2.Take(50)), halfPt: 16);
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §8  CONSOLIDATED AUDITOR COMMENTS
        // ═════════════════════════════════════════════════════════════════════
        AddH1(body, "8  Consolidated Auditor Comments");
        HRule(body);
        AddPara(body,
            "Digital comments saved before export. Handwritten notes should be transcribed here.",
            halfPt: 18);
        body.AppendChild(new Paragraph());

        var savedCmts = comments
            .Where(kv => !string.IsNullOrWhiteSpace(kv.Value))
            .OrderBy(kv => kv.Key)
            .ToList();

        if (savedCmts.Count > 0)
            AddTable(body,
                headers: new[] { "Reference", "Comment" },
                rows: savedCmts.Select(kv => new[] { kv.Key, kv.Value.Trim() }).ToArray(),
                headerFill: "78350F",
                colWidths: new[] { 2551, 7087 }
            );
        else
            AddPara(body, "No digital comments recorded.", halfPt: 18);

        body.AppendChild(new Paragraph());
        var concP = body.AppendChild(new Paragraph());
        concP.AppendChild(MakeRun("Overall Audit Conclusion / Sign-off Notes:",
            bold: true, halfPt: 18));
        AddCommentBox(body, "overall", "Overall conclusion", comments);
        PageBreak(body);

        // ═════════════════════════════════════════════════════════════════════
        // §9  MATERIALITY REVIEW FINDINGS
        // ═════════════════════════════════════════════════════════════════════
        AddH1(body, "9  Materiality Review Findings");
        HRule(body);
        AddPara(body,
            "This section presents the results of the independent document analysis " +
            "conducted by reading and comparing both AFS documents in full. " +
            "The analysis identifies material differences in disclosures, narrative, " +
            "and financial data. Findings are sourced from the Document Reading review " +
            "and are classified by severity (High / Medium / Low) and type " +
            "(Factual Update / Narrative / Numeric / Structural / Policy Change). " +
            "Findings are aligned with the Review tab results shown in the system and " +
            "must be considered alongside the page-by-page snapshot analysis in Section 2.",
            halfPt: 18);
        body.AppendChild(new Paragraph());

        var ai9   = cmp.AiAssist;
        var read9 = cmp.AiContentRead;
        bool aiReady9   = string.Equals(ai9?.Status,   "ready", StringComparison.OrdinalIgnoreCase);
        bool readReady9 = string.Equals(read9?.Status, "ready", StringComparison.OrdinalIgnoreCase);

        if (!aiReady9 && !readReady9)
        {
            var naBox = body.AppendChild(new Paragraph());
            SetParaFill(naBox, C_DARK);
            naBox.AppendChild(MakeRun(
                "Document review analysis was not available at the time of export. " +
                "Run the comparison, wait for the background analysis to complete, then re-export this working paper.",
                colorHex: C_WHITE, halfPt: 18));
        }
        else
        {
            // ── 9.1  Overall Assessment (Document Reading — primary) ──────────
            var secHdr91 = body.AppendChild(new Paragraph());
            secHdr91.AppendChild(MakeRun("9.1  Overall Assessment", bold: true, halfPt: 22, colorHex: C_DPURPLE));
            HRule(body);

            // Risk badge
            string riskSrc = readReady9 ? read9!.RiskLevel : ai9!.OverallRisk;
            if (!string.IsNullOrWhiteSpace(riskSrc))
            {
                string riskColor9 = riskSrc.ToLowerInvariant() switch
                { "high" => C_RED, "medium" => C_AMBER, _ => C_GREEN };
                var riskP9 = body.AppendChild(new Paragraph());
                SetParaFill(riskP9, riskColor9);
                riskP9.AppendChild(MakeRun(
                    "  OVERALL RISK:  " + riskSrc.ToUpperInvariant() + "  ",
                    bold: true, colorHex: C_WHITE, halfPt: 22));
                body.AppendChild(new Paragraph());
            }

            // Document Reading overall assessment
            if (readReady9 && !string.IsNullOrWhiteSpace(read9!.OverallAssessment))
            {
                var oaBox = body.AppendChild(new Paragraph());
                SetParaFill(oaBox, C_NAVY);
                oaBox.AppendChild(MakeRun(read9.OverallAssessment, colorHex: C_WHITE, halfPt: 18));
                body.AppendChild(new Paragraph());
            }
            else if (aiReady9 && !string.IsNullOrWhiteSpace(ai9!.ExecutiveSummary))
            {
                var oaBox = body.AppendChild(new Paragraph());
                SetParaFill(oaBox, C_NAVY);
                oaBox.AppendChild(MakeRun(ai9.ExecutiveSummary, colorHex: C_WHITE, halfPt: 18));
                body.AppendChild(new Paragraph());
            }

            // Policy changes, new/removed disclosures
            if (readReady9)
            {
                if (read9!.PolicyChanges?.Count > 0)
                {
                    var pcP = body.AppendChild(new Paragraph());
                    pcP.AppendChild(MakeRun("Policy Changes Noted:", bold: true, halfPt: 18, colorHex: C_DPURPLE));
                    foreach (var pc in read9.PolicyChanges)
                    {
                        var pcI = body.AppendChild(new Paragraph());
                        pcI.AppendChild(MakeRun("  • " + pc, halfPt: 16));
                    }
                    body.AppendChild(new Paragraph());
                }
                if (read9.NewDisclosures?.Count > 0)
                {
                    var ndP = body.AppendChild(new Paragraph());
                    ndP.AppendChild(MakeRun("New Disclosures in AFS 2:", bold: true, halfPt: 18, colorHex: C_GREEN));
                    foreach (var nd in read9.NewDisclosures)
                    {
                        var ndI = body.AppendChild(new Paragraph());
                        ndI.AppendChild(MakeRun("  + " + nd, halfPt: 16));
                    }
                    body.AppendChild(new Paragraph());
                }
                if (read9.RemovedDisclosures?.Count > 0)
                {
                    var rdP = body.AppendChild(new Paragraph());
                    rdP.AppendChild(MakeRun("Disclosures Removed from AFS 1:", bold: true, halfPt: 18, colorHex: C_RED));
                    foreach (var rd in read9.RemovedDisclosures)
                    {
                        var rdI = body.AppendChild(new Paragraph());
                        rdI.AppendChild(MakeRun("  − " + rd, halfPt: 16));
                    }
                    body.AppendChild(new Paragraph());
                }
            }

            // ── 9.2  Section-by-Section Changes (Document Reading — primary) ──
            PageBreak(body);
            var secHdr92 = body.AppendChild(new Paragraph());
            secHdr92.AppendChild(MakeRun("9.2  Section-by-Section Changes", bold: true, halfPt: 22, colorHex: C_DPURPLE));
            HRule(body);
            AddPara(body,
                "The following changes were identified by reading each section of both " +
                "documents and comparing their content. Each finding shows the section " +
                "name, change type, severity, and a side-by-side comparison of AFS 1 vs AFS 2.",
                halfPt: 16);
            body.AppendChild(new Paragraph());

            var allSectionDiffs = readReady9 ? (read9!.SectionDiffs ?? new()) : new();

            if (allSectionDiffs.Count == 0 && aiReady9)
            {
                // Fall back to AiAssist material findings when no content read diffs
                int fn = 0;
                foreach (var f in (ai9!.MaterialFindings ?? new()))
                {
                    fn++;
                    string sc = (f.Severity?.ToLowerInvariant() ?? "low") switch
                    { "high" => C_RED, "medium" => C_AMBER, _ => C_GREEN };
                    string pg = f.PageAfs1.HasValue ? $"  (AFS1 p{f.PageAfs1} ↔ AFS2 p{f.PageAfs2})" : "";
                    var fH = body.AppendChild(new Paragraph());
                    SetParaFill(fH, sc);
                    fH.AppendChild(MakeRun(
                        $"Finding {fn}: {f.Title}  [{f.Severity?.ToUpperInvariant()}]{pg}",
                        bold: true, colorHex: C_WHITE, halfPt: 18));
                    if (!string.IsNullOrWhiteSpace(f.AuditRationale))
                        AddPara(body, f.AuditRationale, halfPt: 16);
                    if (!string.IsNullOrWhiteSpace(f.RecommendedProcedure))
                        AddPara(body, "Recommended: " + f.RecommendedProcedure, halfPt: 16);
                    body.AppendChild(new Paragraph());
                }
            }
            else
            {
                int sdNum = 0;
                foreach (var s in allSectionDiffs)
                {
                    sdNum++;
                    string sc = (s.Severity?.ToLowerInvariant() ?? "low") switch
                    { "high" => C_RED, "medium" => C_AMBER, _ => C_GREEN };
                    string pg = s.PageAfs1.HasValue
                        ? $"  (AFS1 p{s.PageAfs1} ↔ AFS2 p{s.PageAfs2})" : "";

                    // Section heading bar
                    var sH = body.AppendChild(new Paragraph());
                    SetParaFill(sH, sc);
                    sH.AppendChild(MakeRun(
                        $"{sdNum}.  {s.SectionName}{pg}",
                        bold: true, colorHex: C_WHITE, halfPt: 20));

                    // Type + severity line
                    var stP = body.AppendChild(new Paragraph());
                    stP.AppendChild(MakeRun(
                        $"Change type: {s.ChangeType}   |   Severity: {s.Severity?.ToUpperInvariant()}",
                        bold: true, halfPt: 14, colorHex: sc));

                    // Description
                    if (!string.IsNullOrWhiteSpace(s.ChangeDescription))
                        AddPara(body, s.ChangeDescription, halfPt: 16);

                    // AFS 1 vs AFS 2 side-by-side
                    if (!string.IsNullOrWhiteSpace(s.Afs1Summary) || !string.IsNullOrWhiteSpace(s.Afs2Summary))
                    {
                        AddTable(body,
                            headers: new[] { "AFS 1 — Before", "AFS 2 — After" },
                            rows: new[] { new[] {
                                Trunc(s.Afs1Summary, 200),
                                Trunc(s.Afs2Summary, 200),
                            }},
                            headerFill: C_DARK,
                            colWidths: new[] { CONTENT_DXA / 2, CONTENT_DXA / 2 });
                    }

                    // Changed values
                    if (s.ChangedValues?.Count > 0)
                    {
                        var cvP = body.AppendChild(new Paragraph());
                        cvP.AppendChild(MakeRun("Changed values:  ", bold: true, halfPt: 14));
                        cvP.AppendChild(MakeRun(string.Join("   |   ", s.ChangedValues), halfPt: 14, colorHex: C_AMBER));
                    }
                    body.AppendChild(new Paragraph());
                }
            }

            // ── 9.3  Line-Item Financial Changes ─────────────────────────────
            if (readReady9 && (read9!.LineItemChanges?.Count ?? 0) > 0)
            {
                PageBreak(body);
                var secHdr93 = body.AppendChild(new Paragraph());
                secHdr93.AppendChild(MakeRun("9.3  Line-Item Financial Changes", bold: true, halfPt: 22, colorHex: C_DPURPLE));
                HRule(body);
                AddPara(body,
                    "The following individual financial line items were identified as changed " +
                    "between AFS 1 and AFS 2. The percentage change and a materiality note " +
                    "are included for each item.",
                    halfPt: 16);
                body.AppendChild(new Paragraph());
                AddTable(body,
                    headers: new[] { "Line Item", "AFS 1 Value", "AFS 2 Value", "Change %", "Materiality Note" },
                    rows: read9.LineItemChanges!.Take(50).Select(li => new[]
                    {
                        Trunc(li.LineItem, 50), li.ValueInAfs1, li.ValueInAfs2,
                        li.PctChange, Trunc(li.MaterialityNote, 100),
                    }).ToArray(),
                    headerFill: C_DPURPLE,
                    colWidths: new[] { 2551, 1418, 1418, 900, 3351 });
                body.AppendChild(new Paragraph());
            }

            // ── 9.4  Audit Conclusion ─────────────────────────────────────────
            string conclusionText = readReady9 && !string.IsNullOrWhiteSpace(read9!.AuditConclusion)
                ? read9.AuditConclusion
                : "";
            if (!string.IsNullOrWhiteSpace(conclusionText))
            {
                var secHdr94 = body.AppendChild(new Paragraph());
                secHdr94.AppendChild(MakeRun("9.4  Audit Conclusion", bold: true, halfPt: 22, colorHex: C_DPURPLE));
                HRule(body);
                var concBox = body.AppendChild(new Paragraph());
                SetParaFill(concBox, C_NAVY);
                concBox.AppendChild(MakeRun(conclusionText, colorHex: C_WHITE, halfPt: 18));
                body.AppendChild(new Paragraph());
            }

            // ── 9.5  System Review Findings (deterministic — no AI) ──────────
            if (cmp.AutoReview.Count > 0)
            {
                PageBreak(body);
                var secHdr95 = body.AppendChild(new Paragraph());
                secHdr95.AppendChild(MakeRun(
                    "9.5  System Review Findings",
                    bold: true, halfPt: 22, colorHex: C_DPURPLE));
                HRule(body);
                AddPara(body,
                    $"The following {cmp.AutoReview.Count} finding(s) were generated deterministically by the comparison " +
                    "engine without AI assistance. Every changed, added, or removed line detected is listed below. " +
                    "Snapshot references indicate the page pair number in Section 2 of this working paper.",
                    halfPt: 16);
                body.AppendChild(new Paragraph());

                int sfNum = 0;
                foreach (var f in cmp.AutoReview)
                {
                    sfNum++;
                    string sfColor = (f.Severity?.ToLowerInvariant() ?? "low") switch
                    { "high" => C_RED, "medium" => C_AMBER, _ => C_GREEN };

                    string pageRef = "";
                    if (f.PageAfs1.HasValue || f.PageAfs2.HasValue)
                    {
                        string p1 = f.PageAfs1.HasValue ? $"AFS 1 p.{f.PageAfs1}" : "";
                        string p2 = f.PageAfs2.HasValue ? $"AFS 2 p.{f.PageAfs2}" : "";
                        pageRef   = "  |  " + string.Join(" ↔ ", new[] { p1, p2 }.Where(s => s.Length > 0));
                        if (f.PairIndex.HasValue)
                            pageRef += $"  (refer to snapshot in §2, Page Pair {f.PairIndex.Value + 1})";
                    }

                    // Finding header
                    var sfH = body.AppendChild(new Paragraph());
                    var sfPPr = sfH.AppendChild(new ParagraphProperties());
                    sfPPr.AppendChild(new SpacingBetweenLines { Before = "60", After = "20" });
                    SetParaFill(sfH, sfColor);
                    sfH.AppendChild(MakeRun(
                        $"  {sfNum}.  {f.Title}  [{f.Severity?.ToUpperInvariant()}]{pageRef}  ",
                        bold: true, colorHex: C_WHITE, halfPt: 17));

                    // Category label + description in proper English
                    string catLabel = f.Category switch
                    {
                        "sign-change"      => "Sign Reversal",
                        "policy-change"    => "Policy Change",
                        "director-change"  => "Key Personnel Change",
                        "section-added"    => "Page / Section Added",
                        "section-removed"  => "Page / Section Removed",
                        "numeric-change"   => "Numeric Change",
                        "content-added"    => "Content Added",
                        "content-removed"  => "Content Removed",
                        _                  => "Text Change",
                    };

                    string descText = f.Description;
                    if (string.IsNullOrWhiteSpace(descText))
                        descText = $"The comparison engine detected a {catLabel.ToLower()} on page {f.PageAfs1 ?? f.PageAfs2}.";

                    var sfDesc = body.AppendChild(new Paragraph());
                    var sfDescPPr = sfDesc.AppendChild(new ParagraphProperties());
                    sfDescPPr.AppendChild(new SpacingBetweenLines { Before = "20", After = "20" });
                    sfDesc.AppendChild(MakeRun($"Category: {catLabel}.  ", bold: true, halfPt: 14));
                    sfDesc.AppendChild(MakeRun(descText, halfPt: 14));

                    // Before / After table for changed lines
                    if (!string.IsNullOrWhiteSpace(f.Before) || !string.IsNullOrWhiteSpace(f.After))
                    {
                        bool hasBoth = !string.IsNullOrWhiteSpace(f.Before) && !string.IsNullOrWhiteSpace(f.After);
                        if (hasBoth)
                        {
                            string pctStr = f.PctChange.HasValue && f.PctChange.Value > 0
                                ? $" ({f.PctChange.Value:F1}% change)" : "";
                            AddTable(body,
                                headers: new[] { $"AFS 1 — Before{pctStr}", "AFS 2 — After" },
                                rows: new[] { new[] { Trunc(f.Before ?? "", 200), Trunc(f.After ?? "", 200) } },
                                headerFill: C_DARK,
                                colWidths: new[] { CONTENT_DXA / 2, CONTENT_DXA / 2 });
                        }
                        else if (!string.IsNullOrWhiteSpace(f.Before))
                        {
                            var bfP = body.AppendChild(new Paragraph());
                            bfP.AppendChild(MakeRun("Removed from AFS 1: ", bold: true, halfPt: 13, colorHex: C_RED));
                            bfP.AppendChild(MakeRun(Trunc(f.Before ?? "", 200), halfPt: 13));
                        }
                        else
                        {
                            var bfP = body.AppendChild(new Paragraph());
                            bfP.AppendChild(MakeRun("Added in AFS 2: ", bold: true, halfPt: 13, colorHex: C_GREEN));
                            bfP.AppendChild(MakeRun(Trunc(f.After ?? "", 200), halfPt: 13));
                        }
                    }
                    body.AppendChild(new Paragraph());
                }
            }

            // ── Limitations ───────────────────────────────────────────────────
            var limits9 = (ai9?.Limitations ?? new()).Concat(read9?.Limitations ?? new()).Distinct().ToList();
            if (limits9.Count > 0)
            {
                var limP = body.AppendChild(new Paragraph());
                limP.AppendChild(MakeRun("Analysis Limitations:", bold: true, halfPt: 16, colorHex: "6B7280"));
                foreach (var lim in limits9)
                {
                    var limI = body.AppendChild(new Paragraph());
                    limI.AppendChild(MakeRun("  • " + lim, halfPt: 14, colorHex: "6B7280"));
                }
            }
        }

        // ─────────────────────────────────────────────────────────────────
        // SECTION PROPERTIES — must be the LAST element in Body (OOXML spec)
        // Placing it anywhere else causes Word to refuse to open the file.
        // ─────────────────────────────────────────────────────────────────
        body.AppendChild(new SectionProperties(
            new PageSize   { Width = 11906, Height = 16838 },
            new PageMargin { Top = 1134, Bottom = 1134, Left = 1134, Right = 1134 }));

        mainPart.Document.Save();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // COVER PAGE
    // ─────────────────────────────────────────────────────────────────────────

    private static void AddCoverPage(Body body, EngagementDetails eng, string ts)
    {
        var sub = body.AppendChild(new Paragraph());
        sub.AppendChild(MakeRun("SNG GRANT THORNTON \u2014 CAATs PLATFORM",
            colorHex: C_PURPLE, halfPt: 18, bold: true));

        var title = body.AppendChild(new Paragraph());
        title.AppendChild(MakeRun("AFS Comparison Working Paper",
            bold: true, colorHex: C_NAVY, halfPt: 56));

        var subT = body.AppendChild(new Paragraph());
        subT.AppendChild(MakeRun(
            "Automated Line \u00b7 Number \u00b7 Page Snapshot Analysis \u2014 v4.3.2",
            colorHex: "7B4FFF", halfPt: 22));

        body.AppendChild(new Paragraph());

        // Engagement info table — property names match Models.cs EngagementDetails
        var engRows = new (string Label, string? Val)[]
        {
            ("Parent Name",          eng.Parent),
            ("Client Name",          eng.Client),
            ("Engagement Name",      eng.EngagementName),
            ("Engagement No.",       eng.EngagementNumber),
            ("Financial Year Start", eng.FinancialYearStart),
            ("Financial Year End",   eng.FinancialYearEnd),
            ("Prepared by",          eng.PreparedBy),
            ("Director",             eng.Director),
            ("Manager",              eng.Manager),
            ("Generated",            ts),
        };

        var tbl = MakeTable(CONTENT_DXA, new[] { 3118, 6520 });

        var hRow = tbl.AppendChild(new TableRow());
        var hCell = MakeCell("Engagement Information",
            fill: C_DPURPLE, bold: true, colorHex: C_WHITE, width: CONTENT_DXA);
        hCell.TableCellProperties!.AppendChild(new GridSpan { Val = 2 });
        hRow.AppendChild(hCell);

        for (int i = 0; i < engRows.Length; i++)
        {
            string fill = i % 2 == 0 ? C_LTGRAY : C_WHITE;
            var row = tbl.AppendChild(new TableRow());
            row.AppendChild(MakeCell(engRows[i].Label, fill: fill, bold: true, width: 3118));
            row.AppendChild(MakeCell(engRows[i].Val ?? "\u2014", fill: fill, width: 6520));
        }
        body.AppendChild(tbl);
        body.AppendChild(new Paragraph());

        string objText = string.IsNullOrWhiteSpace(eng.Objective)
            ? "To independently compare two versions of the Annual Financial Statements (AFS) " +
              "through automated computer-assisted audit techniques (CAATs), identifying all " +
              "textual, numerical, and structural differences on a line-by-line and number-by-number " +
              "basis, in order to support audit evidence gathering and financial statement accuracy " +
              "assessments in accordance with ISA 500 (Audit Evidence)."
            : eng.Objective;

        var objTbl = MakeTable(CONTENT_DXA, new[] { CONTENT_DXA });
        var ohRow = objTbl.AppendChild(new TableRow());
        ohRow.AppendChild(MakeCell("Audit Objective",
            fill: C_DPURPLE, bold: true, colorHex: C_WHITE, width: CONTENT_DXA));
        var ovRow = objTbl.AppendChild(new TableRow());
        ovRow.AppendChild(MakeCell(objText, fill: C_LTGRAY, width: CONTENT_DXA));
        body.AppendChild(objTbl);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SIDE-BY-SIDE IMAGE ROW  (caption + snapshot locked in one fixed table)
    // ─────────────────────────────────────────────────────────────────────────

    private void AddImageRow(
        MainDocumentPart mainPart,
        Body body,
        string? b64Left,
        string? b64Right,
        string cap1, string fill1,
        string cap2, string fill2)
    {
        int half = CONTENT_DXA / 2;
        var tbl  = MakeTable(CONTENT_DXA, new[] { half, half });

        // ── Caption row (exact height, cannot split across pages) ─────────────
        var capRow  = tbl.AppendChild(new TableRow());
        var capTrPr = capRow.AppendChild(new TableRowProperties());
        capTrPr.AppendChild(new CantSplit());
        capTrPr.AppendChild(new TableRowHeight { Val = 370, HeightType = HeightRuleValues.Exact });

        foreach (var (cap, fill) in new[] { (cap1, fill1), (cap2, fill2) })
        {
            var cell = capRow.AppendChild(new TableCell());
            var tcPr = cell.AppendChild(new TableCellProperties());
            tcPr.AppendChild(new TableCellWidth
                { Width = half.ToString(), Type = TableWidthUnitValues.Dxa });
            tcPr.AppendChild(new Shading
                { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill });
            tcPr.AppendChild(new TableCellVerticalAlignment
                { Val = TableVerticalAlignmentValues.Center });
            var capP  = cell.AppendChild(new Paragraph());
            var capPr = capP.AppendChild(new ParagraphProperties());
            capPr.AppendChild(new SpacingBetweenLines { Before = "0", After = "0" });
            capP.AppendChild(MakeRun(cap, bold: true, halfPt: 16, colorHex: C_DARK));
        }

        // ── Image row (cannot split; zero-margin cells for pixel-perfect fit) ──
        var imgRow  = tbl.AppendChild(new TableRow());
        var imgTrPr = imgRow.AppendChild(new TableRowProperties());
        imgTrPr.AppendChild(new CantSplit());

        for (int side = 0; side < 2; side++)
        {
            string? b64  = side == 0 ? b64Left : b64Right;
            var     cell = imgRow.AppendChild(new TableCell());
            var     tcPr = cell.AppendChild(new TableCellProperties());
            tcPr.AppendChild(new TableCellWidth
                { Width = half.ToString(), Type = TableWidthUnitValues.Dxa });
            tcPr.AppendChild(new Shading
                { Val = ShadingPatternValues.Clear, Color = "auto", Fill = C_WHITE });
            tcPr.AppendChild(new TableCellVerticalAlignment
                { Val = TableVerticalAlignmentValues.Top });
            tcPr.AppendChild(new TableCellMargin(
                new TopMargin    { Width = "40", Type = TableWidthUnitValues.Dxa },
                new BottomMargin { Width = "40", Type = TableWidthUnitValues.Dxa },
                new LeftMargin   { Width = "40", Type = TableWidthUnitValues.Dxa },
                new RightMargin  { Width = "40", Type = TableWidthUnitValues.Dxa }));

            if (!string.IsNullOrEmpty(b64))
            {
                try
                {
                    byte[] imgBytes = Convert.FromBase64String(b64);
                    var imgPart = mainPart.AddImagePart(ImagePartType.Png);
                    using var ms = new MemoryStream(imgBytes);
                    imgPart.FeedData(ms);

                    // Width fills the cell exactly (cell width minus 40+40 DXA margins; 1 DXA = 635 EMU)
                    long imgWidthEmu = (long)(half - 80) * 635L;
                    long heightEmu;
                    using (var bmp = SKBitmap.Decode(imgBytes))
                        heightEmu = bmp.Width > 0 ? imgWidthEmu * bmp.Height / bmp.Width : imgWidthEmu * 4 / 3;

                    string relId   = mainPart.GetIdOfPart(imgPart);
                    uint   drawId  = _nextDrawingId++;   // unique ID per drawing
                    var    inline  = BuildInlineDrawing(relId, imgWidthEmu, heightEmu, drawId);
                    // CRITICAL: Inline must be wrapped in w:drawing inside the Run.
                    // Placing Inline directly in Run produces invalid OOXML that Word cannot open.
                    var    drawing = new Drawing(inline);

                    var imgPara = cell.AppendChild(new Paragraph());
                    var imgPPr  = imgPara.AppendChild(new ParagraphProperties());
                    imgPPr.AppendChild(new SpacingBetweenLines { Before = "0", After = "0" });
                    imgPPr.AppendChild(new Justification { Val = JustificationValues.Left });
                    imgPara.AppendChild(new Run(drawing));
                }
                catch
                {
                    var p = cell.AppendChild(new Paragraph());
                    p.AppendChild(MakeRun("[Image render error]", halfPt: 16, colorHex: "888888"));
                }
            }
            else
            {
                var p = cell.AppendChild(new Paragraph());
                p.AppendChild(MakeRun(
                    "[Image unavailable \u2014 Docnet.Core required]",
                    halfPt: 16, colorHex: "888888"));
            }
        }
        body.AppendChild(tbl);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // AUDITOR COMMENT BOX
    // ─────────────────────────────────────────────────────────────────────────

    private static void AddCommentBox(
        Body body, string key, string label,
        Dictionary<string, string> comments)
    {
        string existing = comments.TryGetValue(key, out var v) ? v.Trim() : "";

        var hdrP = body.AppendChild(new Paragraph());
        SetParaFill(hdrP, "78350F");
        hdrP.PrependChild(new ParagraphProperties(
            new SpacingBetweenLines { Before = "120" }));
        hdrP.AppendChild(MakeRun(
            "\u270f  AUDITOR NOTE   " + label,
            bold: true, colorHex: C_WHITE, halfPt: 18));

        var noteTbl = MakeTable(CONTENT_DXA, new[] { CONTENT_DXA });
        for (int i = 0; i < 5; i++)
        {
            var row  = noteTbl.AppendChild(new TableRow());
            var trPr = row.AppendChild(new TableRowProperties());
            trPr.AppendChild(new TableRowHeight
                { Val = 454, HeightType = HeightRuleValues.Exact });

            var cell = row.AppendChild(new TableCell());
            var tcPr = cell.AppendChild(new TableCellProperties());
            tcPr.AppendChild(new TableCellWidth
                { Width = CONTENT_DXA.ToString(), Type = TableWidthUnitValues.Dxa });
            tcPr.AppendChild(new Shading
                { Val = ShadingPatternValues.Clear, Color = "auto", Fill = C_CMT_BG });
            tcPr.AppendChild(new TableCellBorders(
                new BottomBorder
                {
                    Val   = BorderValues.Single,
                    Size  = 4,
                    Color = "D97706",
                    Space = 0,
                }));

            var lineP = cell.AppendChild(new Paragraph());
            if (i == 0 && !string.IsNullOrEmpty(existing))
            {
                string preview = existing.Length > 150 ? existing[..150] : existing;
                lineP.AppendChild(MakeRun(preview, halfPt: 16, colorHex: "92400E", italic: true));
            }
        }
        body.AppendChild(noteTbl);

        int sfW = CONTENT_DXA / 3;
        var sfTbl = MakeTable(CONTENT_DXA, new[] { sfW, sfW, CONTENT_DXA - sfW * 2 });
        var sfRow = sfTbl.AppendChild(new TableRow());
        foreach (string lbl in new[] {
            "Initials: ___________",
            "Date: _______________",
            "Ref: ________________" })
        {
            var c  = sfRow.AppendChild(new TableCell());
            var cp = c.AppendChild(new TableCellProperties());
            cp.AppendChild(new TableCellWidth
                { Width = sfW.ToString(), Type = TableWidthUnitValues.Dxa });
            cp.AppendChild(new Shading
                { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FEF3C7" });
            c.AppendChild(new Paragraph(MakeRun(lbl, halfPt: 16, colorHex: "6B7280")));
        }
        body.AppendChild(sfTbl);
        body.AppendChild(new Paragraph());
    }

    // ─────────────────────────────────────────────────────────────────────────
    // REUSABLE DOCUMENT BUILDERS
    // ─────────────────────────────────────────────────────────────────────────

    private static void AddH1(Body body, string text)
    {
        var p  = body.AppendChild(new Paragraph());
        p.AppendChild(new ParagraphProperties(
            new SpacingBetweenLines { Before = "240", After = "120" }));
        p.AppendChild(MakeRun(text, bold: true, colorHex: C_PURPLE, halfPt: 28));
    }

    private static void HRule(Body body)
    {
        var p  = body.AppendChild(new Paragraph());
        var pP = p.AppendChild(new ParagraphProperties());
        pP.AppendChild(new ParagraphBorders(
            new BottomBorder
            {
                Val   = BorderValues.Single,
                Size  = 8,
                Space = 1,
                Color = C_PURPLE,
            }));
        pP.AppendChild(new SpacingBetweenLines { After = "120" });
    }

    private static string FormatWordDiff(List<WordDiffToken> tokens)
    {
        if (tokens.Count == 0) return "—";
        var parts = tokens
            .Where(t => t.Tag != "same")
            .Select(t => t.Tag == "removed"
                ? $"-\"{Trunc(t.Word, 18)}\""
                : $"+\"{Trunc(t.Word, 18)}\"")
            .Take(12)
            .ToList();
        return parts.Count > 0 ? string.Join("  ", parts) : "—";
    }

    private static void AddPara(Body body, string text,
        int halfPt = 18, string colorHex = "000000", bool bold = false)
    {
        var p = body.AppendChild(new Paragraph());
        p.AppendChild(MakeRun(text, halfPt: halfPt, colorHex: colorHex, bold: bold));
    }



    private static void AddTable(Body body,
        string[] headers, string[][] rows,
        string headerFill = "1E2F5A",
        int[]? colWidths = null)
    {
        int nCols = headers.Length;
        if (colWidths == null)
        {
            colWidths = new int[nCols];
            int w = CONTENT_DXA / nCols;
            for (int i = 0; i < nCols; i++) colWidths[i] = w;
            colWidths[nCols - 1] += CONTENT_DXA - colWidths.Sum();
        }

        var tbl = MakeTable(CONTENT_DXA, colWidths);

        var hRow = tbl.AppendChild(new TableRow());
        for (int j = 0; j < nCols; j++)
            hRow.AppendChild(MakeCell(headers[j],
                fill: headerFill, bold: true,
                colorHex: C_WHITE, width: colWidths[j]));

        for (int i = 0; i < rows.Length; i++)
        {
            string fill = i % 2 == 0 ? C_LTGRAY : C_WHITE;
            var row = tbl.AppendChild(new TableRow());
            for (int j = 0; j < nCols; j++)
                row.AppendChild(MakeCell(
                    j < rows[i].Length ? rows[i][j] : "",
                    fill: fill, width: colWidths[j]));
        }
        body.AppendChild(tbl);
        body.AppendChild(new Paragraph());
    }

    private static void PageBreak(Body body)
        => body.AppendChild(new Paragraph(
               new Run(new Break { Type = BreakValues.Page })));

    // ─────────────────────────────────────────────────────────────────────────
    // LOW-LEVEL OpenXML PRIMITIVES
    // ─────────────────────────────────────────────────────────────────────────

    private static Table MakeTable(int totalWidth, int[] colWidths)
    {
        var tbl = new Table();
        tbl.AppendChild(new TableProperties(
            new TableWidth  { Width = totalWidth.ToString(), Type = TableWidthUnitValues.Dxa },
            new TableLayout { Type = TableLayoutValues.Fixed },
            new TableBorders(
                new TopBorder              { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new BottomBorder           { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new LeftBorder             { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new RightBorder            { Val = BorderValues.Single, Size = 4, Color = "AAAAAA" },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "CCCCCC" },
                new InsideVerticalBorder   { Val = BorderValues.Single, Size = 4, Color = "CCCCCC" })));
        var grid = tbl.AppendChild(new TableGrid());
        foreach (int w in colWidths)
            grid.AppendChild(new GridColumn { Width = w.ToString() });
        return tbl;
    }

    private static TableCell MakeCell(string text,
        string fill = "FFFFFF",
        bool bold = false,
        string colorHex = "000000",
        int width = 0)
    {
        var cell = new TableCell();
        var tcPr = cell.AppendChild(new TableCellProperties());
        tcPr.AppendChild(new Shading
            { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill });
        tcPr.AppendChild(new TableCellMargin(
            new TopMargin    { Width = "80",  Type = TableWidthUnitValues.Dxa },
            new BottomMargin { Width = "80",  Type = TableWidthUnitValues.Dxa },
            new LeftMargin   { Width = "100", Type = TableWidthUnitValues.Dxa },
            new RightMargin  { Width = "100", Type = TableWidthUnitValues.Dxa }));
        if (width > 0)
            tcPr.AppendChild(new TableCellWidth
                { Width = width.ToString(), Type = TableWidthUnitValues.Dxa });

        cell.AppendChild(new Paragraph(MakeRun(text,
            bold: bold, colorHex: colorHex, halfPt: 16)));
        return cell;
    }

    private static Run MakeRun(string text,
        bool bold = false, bool italic = false,
        string colorHex = "000000", int halfPt = 18)
    {
        var rPr = new RunProperties(
            new FontSize { Val = halfPt.ToString() },
            new Color    { Val = colorHex },
            new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" });
        if (bold)   rPr.AppendChild(new Bold());
        if (italic) rPr.AppendChild(new Italic());
        return new Run(rPr,
            new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static void SetParaFill(Paragraph p, string fill)
    {
        var pPr = p.GetFirstChild<ParagraphProperties>()
                  ?? p.PrependChild(new ParagraphProperties());
        pPr.AppendChild(new Shading
            { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill });
    }

    private static DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline
        BuildInlineDrawing(string relId, long widthEmu, long heightEmu, uint id)
    {
        var pic = new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                    { Id = id, Name = "AFS_Page_" + id },
                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                new DocumentFormat.OpenXml.Drawing.Blip { Embed = relId },
                new DocumentFormat.OpenXml.Drawing.Stretch(
                    new DocumentFormat.OpenXml.Drawing.FillRectangle())),
            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                new DocumentFormat.OpenXml.Drawing.Transform2D(
                    new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                    new DocumentFormat.OpenXml.Drawing.Extents
                        { Cx = widthEmu, Cy = heightEmu }),
                new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                    new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                    { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }));

        return new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent
                { Cx = widthEmu, Cy = heightEmu },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent
                { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                { Id = id, Name = "AFS_Page_" + id },
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks
                    { NoChangeAspect = true }),
            new DocumentFormat.OpenXml.Drawing.Graphic(
                new DocumentFormat.OpenXml.Drawing.GraphicData(pic)
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
        {
            DistanceFromTop    = 0,
            DistanceFromBottom = 0,
            DistanceFromLeft   = 0,
            DistanceFromRight  = 0,
        };
    }

    // ─────────────────────────────────────────────────────────────────────────
    // UTILITY
    // ─────────────────────────────────────────────────────────────────────────

    private static string Trunc(string? s, int maxLen)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s.Length <= maxLen ? s : s[..maxLen] + "\u2026";
    }

    private static string Pct(int n, int total)
        => Math.Round(n * 100.0 / Math.Max(total, 1), 1) + "%";
}
