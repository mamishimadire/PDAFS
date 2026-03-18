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
        string[] tocNums   = { "1","2","3","4","5","6","7","8" };
        string[] tocTitles = {
            "Documents Compared",
            "Page-by-Page Snapshot Review",
            "Comparison Summary",
            "Changed Lines",
            "Added Lines",
            "Removed Lines",
            "Number Comparison",
            "Consolidated Auditor Comments",
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

        var paired = cmp.PageDiffs.Where(pd => pd.PageAfs2.HasValue).ToList();

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
                .Where(d => d.Status is "changed" or "added" or "removed")
                .Take(20).ToList();

            if (diffRows.Count > 0)
            {
                var chHdr = body.AppendChild(new Paragraph());
                var chPPr = chHdr.AppendChild(new ParagraphProperties());
                chPPr.AppendChild(new SpacingBetweenLines { Before = "120", After = "60" });
                chHdr.AppendChild(MakeRun($"Line-by-line changes on this page pair ({diffRows.Count}):",
                    bold: true, halfPt: 16, colorHex: C_DARK));

                AddTable(body,
                    headers: new[] { "", "AFS 1 line  \u2192  AFS 2 line", "Num diff" },
                    rows: diffRows.Select((d, j) =>
                    {
                        string sym = (d.Status == "changed" ? "~"
                                    : d.Status == "added"   ? "+" : "-") + (j + 1);
                        string lineText = d.Status == "changed"
                            ? Trunc(d.Line1, 45) + "  \u2192  " + Trunc(d.Line2, 45)
                            : d.Status == "added" ? "+  " + Trunc(d.Line2, 90)
                                                  : "\u2212  " + Trunc(d.Line1, 90);
                        string numText = d.NumDiff.Count > 0
                            ? string.Join(", ", d.NumDiff) : "\u2014";
                        return new[] { sym, lineText, numText };
                    }).ToArray(),
                    headerFill: C_DARK,
                    colWidths: new[] { 567, 7933, 1138 }
                );
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
