// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3                                                               ║
// ║                                                                              ║
// ║  SERVICE — ExportService                                                     ║
// ║  Generates PDF, Word (.docx), Excel (.xlsx), and plain-text working papers. ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • QuestPDF — MIT-licensed C# PDF generation library                       ║
// ║     https://www.questpdf.com                                                 ║
// ║   • NPOI — Apache-licensed C# port of Apache POI for Word/Excel             ║
// ║     https://github.com/nissl-lab/npoi                                        ║
// ║   • ClosedXML — MIT-licensed Excel generation                               ║
// ║     https://github.com/ClosedXML/ClosedXML                                  ║
// ║   • Python equivalent: _build_pdf_working_paper(),                          ║
// ║     _build_word_working_paper(), _build_text_report() in the notebook       ║
// ║   • SNG purple: #5C2D91 (brand colour — replaces default blue throughout)   ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text;
using AfsPdfComparison.Models;
using ClosedXML.Excel;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;

namespace AfsPdfComparison.Services;

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 5 · EXPORT SERVICE
//
// Responsibility: convert a ComparisonResult + EngagementDetails + auditor
// comments into four archive formats:
//
//  5.1  PDF Working Paper  — QuestPDF A4 document with:
//         cover page (SNG template) | TOC | documents table | page snapshots
//         changed/added/removed lines | number comparison | auditor comments
//
//  5.2  Word Working Paper — NPOI .docx with the same section structure.
//         Auditor can edit directly after export.
//
//  5.3  Excel Workbook     — ClosedXML .xlsx with one sheet per data category:
//         Summary | Changed | Added | Removed | Number_Comparison
//         Page_Alignment | Auditor_Comments
//
//  5.4  Plain-text Report  — UTF-8 .txt audit trail, human-readable.
//
// SNG brand colour: #5C2D91 (purple).
// All section headings use THIS colour.  Navy (#0F1C3F) is used for table headers.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Generates all working-paper export formats from a completed comparison.
/// Registered as a singleton in DI.
/// </summary>
public class ExportService
{
    // SNG brand colours (hex strings used by QuestPDF / NPOI)
    private const string Purple = "#5C2D91";
    private const string Navy   = "#0F1C3F";
    private const string Dark   = "#1E2F5A";
    private const string Amber  = "#F59E0B";
    private const string Green  = "#10B981";
    private const string Red    = "#EF4444";
    private const string Slate  = "#CBD5E1";

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 5.1 · PDF WORKING PAPER
    // Uses QuestPDF fluent API.
    // QuestPDF licence mode is set to Community in Program.cs.
    // Reference: Python _build_pdf_working_paper() in the notebook.
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Generates a professional A4 PDF working paper and saves it to
    /// <paramref name="path"/>.
    /// </summary>
    public void BuildPdf(
        ComparisonResult cmp,
        EngagementDetails eng,
        Dictionary<string, string> comments,
        string path)
    {
        QuestPDF.Settings.License = LicenseType.Community;
        string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        QuestPDF.Fluent.Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(20, Unit.Millimetre);
                page.PageColor(Colors.White);
                page.DefaultTextStyle(x => x.FontSize(9).FontFamily("Arial"));

                // ── Header (every page) ────────────────────────────────────
                page.Header().Element(hdr =>
                {
                    hdr.Column(col =>
                    {
                        col.Item().Background(Navy).Padding(6).Row(r =>
                        {
                            r.RelativeItem().Text(t =>
                            {
                                t.Span("SNG GRANT THORNTON — CAATs PLATFORM  |  ")
                                 .FontColor(Purple).FontSize(8).Bold();
                                t.Span("AFS Comparison Working Paper v4.3")
                                 .FontColor(Colors.White).FontSize(8).Bold();
                            });
                            r.ConstantItem(120).AlignRight().Text(t =>
                            {
                                t.Span("CONFIDENTIAL").FontColor(Amber).FontSize(7).Bold();
                                t.Span("  |  Page ").FontColor(Slate).FontSize(7);
                                t.CurrentPageNumber().FontColor(Slate).FontSize(7);
                                t.Span(" of ").FontColor(Slate).FontSize(7);
                                t.TotalPages().FontColor(Slate).FontSize(7);
                            });
                        });
                        col.Item().BorderBottom(1).BorderColor(Purple).PaddingBottom(2)
                            .Text(t =>
                            {
                                t.Span(eng.Client + "  |  Eng: " + eng.EngagementNumber +
                                       "  |  FY: " + eng.FinancialYearEnd +
                                       "  |  " + ts)
                                 .FontColor(Slate).FontSize(7);
                            });
                    });
                });

                // ── Footer (every page) ────────────────────────────────────
                page.Footer().Background(Dark).Padding(4).Text(t =>
                {
                    t.Span("AFS 1: " + cmp.Report1.Filename +
                           "  |  AFS 2: " + cmp.Report2.Filename)
                     .FontColor(Slate).FontSize(7);
                    t.Span("  |  Prepared: " + eng.PreparedBy +
                           "  |  Manager: " + eng.Manager)
                     .FontColor(Slate).FontSize(7);
                });

                // ── Content ────────────────────────────────────────────────
                page.Content().Column(col =>
                {
                    // COVER PAGE
                    PdfCoverPage(col, eng, cmp, ts);

                    // SECTION 1 · Documents Compared
                    PdfDocumentsSection(col, cmp);

                    // SECTION 2 · Comparison Summary
                    PdfSummarySection(col, cmp);

                    // SECTION 3 · Changed Lines
                    PdfChangedSection(col, cmp, comments);

                    // SECTION 4 · Added Lines
                    PdfAddedSection(col, cmp);

                    // SECTION 5 · Removed Lines
                    PdfRemovedSection(col, cmp);

                    // SECTION 6 · Number Comparison
                    PdfNumberSection(col, cmp);

                    // SECTION 7 · Consolidated Auditor Comments
                    PdfCommentsSection(col, comments);
                });
            });
        }).GeneratePdf(path);
    }

    // ── Cover page ────────────────────────────────────────────────────────────
    private static void PdfCoverPage(
        ColumnDescriptor col, EngagementDetails eng,
        ComparisonResult cmp, string ts)
    {
        col.Item().PaddingTop(20).Column(c =>
        {
            c.Item().Text("SNG GRANT THORNTON — CAATs PLATFORM")
             .FontColor(Purple).FontSize(9).Bold();

            c.Item().PaddingTop(4).Text("AFS Comparison Working Paper")
             .FontColor(Navy).FontSize(22).Bold();

            c.Item().PaddingTop(2).Text(
                "Automated Line · Number · Page Snapshot Analysis — v4.3")
             .FontColor(Purple).FontSize(11);

            c.Item().PaddingTop(4).LineHorizontal(1.5f).LineColor(Purple);
            c.Item().PaddingTop(8);

            // Engagement table
            c.Item().Table(t =>
            {
                t.ColumnsDefinition(cd =>
                {
                    cd.RelativeColumn(3);  // label
                    cd.RelativeColumn(7);  // value
                });
                PdfTblHdrRow(t, "Engagement Information", Purple);
                PdfTblRow(t, "Parent Name",           eng.Parent);
                PdfTblRow(t, "Client Name",           eng.Client);
                PdfTblRow(t, "Engagement Name",       eng.EngagementName);
                PdfTblRow(t, "Engagement No.",        eng.EngagementNumber);
                PdfTblRow(t, "Financial Year Start",  eng.FinancialYearStart);
                PdfTblRow(t, "Financial Year End",    eng.FinancialYearEnd);
                PdfTblRow(t, "Prepared by",           eng.PreparedBy);
                PdfTblRow(t, "Director",              eng.Director);
                PdfTblRow(t, "Manager",               eng.Manager);
                PdfTblRow(t, "Generated",             ts);
            });

            c.Item().PaddingTop(6);

            // Audit objective
            c.Item().Table(t =>
            {
                t.ColumnsDefinition(cd => cd.RelativeColumn());
                PdfTblHdrRow(t, "Audit Objective", Purple);
                t.Cell().Background("#F1F5F9").Padding(6)
                    .Text(eng.Objective).FontSize(9);
            });

            c.Item().PageBreak();
        });
    }

    // ── Documents section ─────────────────────────────────────────────────────
    private static void PdfDocumentsSection(ColumnDescriptor col, ComparisonResult cmp)
    {
        col.Item().Column(c =>
        {
            PdfSectionHeading(c, "1  Documents Compared");
            c.Item().Table(t =>
            {
                t.ColumnsDefinition(cd =>
                {
                    cd.RelativeColumn(1);  // label
                    cd.RelativeColumn(4);  // filename
                    cd.RelativeColumn(1);  // type
                    cd.RelativeColumn(1);  // year
                    cd.RelativeColumn(1);  // pages
                    cd.RelativeColumn(1);  // words
                });
                void Row(PdfReport r)
                {
                    t.Cell().Background("#F1F5F9").Padding(4).Text(r.Label).FontSize(8);
                    t.Cell().Background("#FFFFFF").Padding(4).Text(r.Filename).FontSize(8);
                    t.Cell().Background("#F1F5F9").Padding(4).Text(r.DocType).FontSize(8);
                    t.Cell().Background("#FFFFFF").Padding(4).Text(r.PrimaryYear?.ToString() ?? "?").FontSize(8);
                    t.Cell().Background("#F1F5F9").Padding(4).Text(r.PageCount.ToString()).FontSize(8);
                    t.Cell().Background("#FFFFFF").Padding(4).Text(r.WordCount.ToString()).FontSize(8);
                }
                foreach (var h in new[]{"Label","Filename","Type","Year","Pages","Words"})
                    t.Cell().Background(Navy).Padding(4).Text(h).FontSize(8).Bold().FontColor(Colors.White);
                Row(cmp.Report1);
                Row(cmp.Report2);
            });
            c.Item().PageBreak();
        });
    }

    // ── Summary section ───────────────────────────────────────────────────────
    private static void PdfSummarySection(ColumnDescriptor col, ComparisonResult cmp)
    {
        int total = cmp.Counts.Values.Sum();
        if (total == 0) total = 1;

        col.Item().Column(c =>
        {
            PdfSectionHeading(c, "2  Comparison Summary");
            c.Item().Table(t =>
            {
                t.ColumnsDefinition(cd =>
                {
                    cd.RelativeColumn(5); cd.RelativeColumn(2); cd.RelativeColumn(2);
                });
                foreach (var h in new[]{"Metric","Count","% of Total"})
                    t.Cell().Background(Navy).Padding(4).Text(h).FontSize(8).Bold().FontColor(Colors.White);

                void SumRow(string label, string key, string color)
                {
                    int cnt = cmp.Counts.GetValueOrDefault(key);
                    t.Cell().Background("#F1F5F9").Padding(4).Text(label).FontSize(8);
                    t.Cell().Background("#FFFFFF").Padding(4).Text(cnt.ToString()).FontSize(8).FontColor(color);
                    t.Cell().Background("#F1F5F9").Padding(4)
                        .Text(Math.Round((double)cnt / total * 100, 1) + "%").FontSize(8);
                }
                SumRow("Same Lines",    "same",    Green);
                SumRow("Changed Lines", "changed", Amber);
                SumRow("Added Lines",   "added",   "#3B82F6");
                SumRow("Removed Lines", "removed", Red);
                t.Cell().Background("#F1F5F9").Padding(4).Text("Number Similarity").FontSize(8);
                t.Cell().Background("#FFFFFF").Padding(4)
                    .Text(cmp.NumCmp.SimilarityPct + "%").FontSize(8).FontColor(Purple);
                t.Cell().Background("#F1F5F9").Padding(4).Text("—").FontSize(8);
            });
            c.Item().PageBreak();
        });
    }

    // ── Changed lines section ─────────────────────────────────────────────────
    private static void PdfChangedSection(
        ColumnDescriptor col, ComparisonResult cmp,
        Dictionary<string, string> comments)
    {
        var changed = cmp.ChangedLines.Take(200).ToList();
        col.Item().Column(c =>
        {
            PdfSectionHeading(c, $"3  Changed Lines  ({changed.Count})");
            if (!changed.Any()) { c.Item().Text("No changed lines detected.").FontSize(9); }
            else
            {
                c.Item().Table(t =>
                {
                    t.ColumnsDefinition(cd =>
                    {
                        cd.ConstantColumn(20); cd.RelativeColumn(4);
                        cd.RelativeColumn(4);  cd.RelativeColumn(2);
                    });
                    foreach (var h in new[]{"#","AFS 1 Line","AFS 2 Line","Num Diff"})
                        t.Cell().Background(Navy).Padding(4).Text(h).FontSize(8).Bold().FontColor(Colors.White);

                    for (int i = 0; i < changed.Count; i++)
                    {
                        var d   = changed[i];
                        string bg = i % 2 == 0 ? "#FFFBEB" : "#FEF9E7";
                        t.Cell().Background(bg).Padding(3).Text((i+1).ToString()).FontSize(7).FontColor(Amber);
                        t.Cell().Background(bg).Padding(3).Text(d.Line1.Truncate(80)).FontSize(7);
                        t.Cell().Background(bg).Padding(3).Text(d.Line2.Truncate(80)).FontSize(7).FontColor("#3B82F6");
                        t.Cell().Background(bg).Padding(3)
                            .Text(d.NumDiff.Any() ? string.Join(", ", d.NumDiff) : "—")
                            .FontSize(7).FontColor(Amber);
                    }
                });
            }
            // Auditor comment boxes (first 10)
            c.Item().PaddingTop(8).Text("Auditor Comments — Changed Lines").FontSize(9).Bold().FontColor(Amber);
            for (int i = 0; i < Math.Min(10, changed.Count); i++)
            {
                string key = "changed:" + (i + 1);
                string val = comments.GetValueOrDefault(key, "");
                PdfCommentBox(c, key, "~" + (i+1) + "  " + changed[i].Line1.Truncate(50), val);
            }
            c.Item().PageBreak();
        });
    }

    // ── Added lines section ───────────────────────────────────────────────────
    private static void PdfAddedSection(ColumnDescriptor col, ComparisonResult cmp)
    {
        var added = cmp.AddedLines.Take(200).ToList();
        col.Item().Column(c =>
        {
            PdfSectionHeading(c, $"4  Added Lines  ({added.Count} in AFS 2)");
            if (!added.Any()) { c.Item().Text("No added lines.").FontSize(9); }
            else
            {
                c.Item().Table(t =>
                {
                    t.ColumnsDefinition(cd => { cd.ConstantColumn(20); cd.RelativeColumn(); });
                    foreach (var h in new[]{"#","Added Line (AFS 2)"})
                        t.Cell().Background(Dark).Padding(4).Text(h).FontSize(8).Bold().FontColor(Colors.White);
                    for (int i = 0; i < added.Count; i++)
                    {
                        string bg = i % 2 == 0 ? "#EFF6FF" : "#DBEAFE";
                        t.Cell().Background(bg).Padding(3).Text((i+1).ToString()).FontSize(7).FontColor("#3B82F6");
                        t.Cell().Background(bg).Padding(3).Text(added[i].Line2.Truncate(120)).FontSize(7).FontColor("#1D4ED8");
                    }
                });
            }
            c.Item().PageBreak();
        });
    }

    // ── Removed lines section ─────────────────────────────────────────────────
    private static void PdfRemovedSection(ColumnDescriptor col, ComparisonResult cmp)
    {
        var removed = cmp.RemovedLines.Take(200).ToList();
        col.Item().Column(c =>
        {
            PdfSectionHeading(c, $"5  Removed Lines  ({removed.Count} from AFS 1)");
            if (!removed.Any()) { c.Item().Text("No removed lines.").FontSize(9); }
            else
            {
                c.Item().Table(t =>
                {
                    t.ColumnsDefinition(cd => { cd.ConstantColumn(20); cd.RelativeColumn(); });
                    foreach (var h in new[]{"#","Removed Line (AFS 1)"})
                        t.Cell().Background("#5F1D1D").Padding(4).Text(h).FontSize(8).Bold().FontColor(Colors.White);
                    for (int i = 0; i < removed.Count; i++)
                    {
                        string bg = i % 2 == 0 ? "#FEF2F2" : "#FEE2E2";
                        t.Cell().Background(bg).Padding(3).Text((i+1).ToString()).FontSize(7).FontColor(Red);
                        t.Cell().Background(bg).Padding(3).Text(removed[i].Line1.Truncate(120)).FontSize(7).FontColor("#991B1B");
                    }
                });
            }
            c.Item().PageBreak();
        });
    }

    // ── Number comparison section ─────────────────────────────────────────────
    private static void PdfNumberSection(ColumnDescriptor col, ComparisonResult cmp)
    {
        var n = cmp.NumCmp;
        col.Item().Column(c =>
        {
            PdfSectionHeading(c, "6  Number Comparison");
            c.Item().Table(t =>
            {
                t.ColumnsDefinition(cd => { cd.RelativeColumn(5); cd.RelativeColumn(2); });
                foreach (var h in new[]{"Metric","Value"})
                    t.Cell().Background(Navy).Padding(4).Text(h).FontSize(8).Bold().FontColor(Colors.White);
                void Row(string l, string v) {
                    t.Cell().Background("#F1F5F9").Padding(4).Text(l).FontSize(8);
                    t.Cell().Background("#FFFFFF").Padding(4).Text(v).FontSize(8).FontColor(Purple);
                }
                Row("AFS 1 unique numbers",  n.CountAfs1.ToString());
                Row("AFS 2 unique numbers",  n.CountAfs2.ToString());
                Row("Numbers in both",       n.InBoth.Count.ToString());
                Row("Only in AFS 1",         n.OnlyInAfs1.Count.ToString());
                Row("Only in AFS 2",         n.OnlyInAfs2.Count.ToString());
                Row("Similarity",            n.SimilarityPct + "%");
            });
            if (n.OnlyInAfs1.Any())
                c.Item().PaddingTop(6).Text("Only in AFS 1: " + string.Join(", ", n.OnlyInAfs1.Take(50)))
                 .FontSize(8).FontFamily("Courier New").FontColor(Red);
            if (n.OnlyInAfs2.Any())
                c.Item().PaddingTop(4).Text("Only in AFS 2: " + string.Join(", ", n.OnlyInAfs2.Take(50)))
                 .FontSize(8).FontFamily("Courier New").FontColor("#3B82F6");
            c.Item().PageBreak();
        });
    }

    // ── Comments section ──────────────────────────────────────────────────────
    private static void PdfCommentsSection(
        ColumnDescriptor col, Dictionary<string, string> comments)
    {
        col.Item().Column(c =>
        {
            PdfSectionHeading(c, "7  Consolidated Auditor Comments");
            var saved = comments.Where(kv => !string.IsNullOrWhiteSpace(kv.Value)).ToList();
            if (!saved.Any())
            {
                c.Item().Text("No digital comments recorded.").FontSize(9);
            }
            else
            {
                c.Item().Table(t =>
                {
                    t.ColumnsDefinition(cd => { cd.RelativeColumn(2); cd.RelativeColumn(7); });
                    foreach (var h in new[]{"Reference","Comment"})
                        t.Cell().Background("#78350F").Padding(4).Text(h).FontSize(8).Bold().FontColor(Colors.White);
                    bool alt = false;
                    foreach (var (k, v) in saved.OrderBy(x => x.Key))
                    {
                        string bg = alt ? "#FFFBEB" : "#FEF3C7";
                        t.Cell().Background(bg).Padding(4).Text(k).FontSize(8).FontColor(Amber);
                        t.Cell().Background(bg).Padding(4).Text(v).FontSize(8);
                        alt = !alt;
                    }
                });
            }
            c.Item().PaddingTop(8);
            PdfCommentBox(c, "overall", "Overall Audit Conclusion / Sign-off", "");
        });
    }

    // ── Reusable helpers ──────────────────────────────────────────────────────
    private static void PdfSectionHeading(ColumnDescriptor col, string title)
    {
        col.Item().PaddingTop(8).Text(title).FontSize(13).Bold().FontColor(Purple);
        col.Item().LineHorizontal(1).LineColor(Purple);
        col.Item().PaddingBottom(4);
    }

    private static void PdfTblHdrRow(TableDescriptor t, string text, string bgColor)
    {
        t.Cell().ColumnSpan(10).Background(bgColor).Padding(5)
         .Text(text).FontSize(9).Bold().FontColor(Colors.White);
    }

    private static void PdfTblRow(TableDescriptor t, string label, string value)
    {
        t.Cell().Background("#F1F5F9").Padding(5).Text(label).FontSize(9).Bold();
        t.Cell().Background("#FFFFFF").Padding(5).Text(value ?? "—").FontSize(9);
    }

    private static void PdfCommentBox(
        ColumnDescriptor col, string key, string label, string existing)
    {
        col.Item().PaddingTop(4).Background("#78350F").Padding(5)
            .Text("✏  AUDITOR NOTE — " + label).FontSize(8).Bold().FontColor(Colors.White);
        col.Item().Background("#FFFBEB").Border(1).BorderColor(Amber).MinHeight(40)
            .Padding(4).Text(string.IsNullOrEmpty(existing) ? " " : existing)
            .FontSize(8).FontColor("#92400E").Italic();
        col.Item().Background("#FEF3C7").Padding(4).Row(r =>
        {
            r.RelativeItem().Text("Initials: ___________").FontSize(7).FontColor("#6B7280");
            r.RelativeItem().Text("Date: ________________").FontSize(7).FontColor("#6B7280");
            r.RelativeItem().Text("Ref: _________________").FontSize(7).FontColor("#6B7280");
        });
        col.Item().PaddingBottom(4);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 5.2 · WORD WORKING PAPER
    // Uses NPOI XWPF (Open XML Word format).
    // Reference: Python _build_word_working_paper() in the notebook.
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Generates an editable Word (.docx) working paper and saves it to
    /// <paramref name="path"/>.
    /// </summary>
    public void BuildWord(
        ComparisonResult cmp,
        EngagementDetails eng,
        Dictionary<string, string> comments,
        string path)
    {
        string ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var doc   = new XWPFDocument();

        // ── Cover heading ──────────────────────────────────────────────────
        WordHeading(doc, "SNG GRANT THORNTON — CAATs PLATFORM", 0, "5C2D91");
        WordHeading(doc, "AFS Comparison Working Paper", 1, "0F1C3F");
        WordBody(doc, "Automated Line · Number · Page Snapshot Analysis — v4.3", "7B4FFF");
        WordBody(doc, "Author: Mamishi Tonny Madire  |  Version 4.3  |  " + ts, "64748B");

        // ── Engagement table ───────────────────────────────────────────────
        WordSectionTitle(doc, "Engagement Information");
        var engRows = new[]
        {
            ("Parent Name",           eng.Parent),
            ("Client Name",           eng.Client),
            ("Engagement Name",       eng.EngagementName),
            ("Engagement No.",        eng.EngagementNumber),
            ("Financial Year Start",  eng.FinancialYearStart),
            ("Financial Year End",    eng.FinancialYearEnd),
            ("Prepared by",           eng.PreparedBy),
            ("Director",              eng.Director),
            ("Manager",               eng.Manager),
            ("Generated",             ts),
        };
        var engTbl = doc.CreateTable(engRows.Length + 1, 2);
        WordTableCell(engTbl.GetRow(0).GetCell(0), "Field", bold: true, bg: "3D1A6E", fg: "FFFFFF");
        WordTableCell(engTbl.GetRow(0).GetCell(1), "Value", bold: true, bg: "3D1A6E", fg: "FFFFFF");
        for (int i = 0; i < engRows.Length; i++)
        {
            string rowBg = i % 2 == 0 ? "F1F5F9" : "FFFFFF";
            WordTableCell(engTbl.GetRow(i + 1).GetCell(0), engRows[i].Item1, bold: true, bg: rowBg);
            WordTableCell(engTbl.GetRow(i + 1).GetCell(1), engRows[i].Item2 ?? "—", bg: rowBg);
        }

        // ── Audit objective ────────────────────────────────────────────────
        WordSectionTitle(doc, "Audit Objective");
        WordBody(doc, eng.Objective, "1A1A1A");

        doc.CreateParagraph(); // spacer

        // ── Section 1 · Documents Compared ────────────────────────────────
        WordSectionTitle(doc, "1  Documents Compared");
        var docTbl = doc.CreateTable(3, 6);
        string[] docHdrs = { "Label", "Filename", "Type", "Year", "Pages", "Words" };
        for (int j = 0; j < docHdrs.Length; j++)
            WordTableCell(docTbl.GetRow(0).GetCell(j), docHdrs[j], bold: true, bg: "1E2F5A", fg: "FFFFFF");
        void DocRow(int rowIdx, PdfReport r) {
            WordTableCell(docTbl.GetRow(rowIdx).GetCell(0), r.Label);
            WordTableCell(docTbl.GetRow(rowIdx).GetCell(1), r.Filename);
            WordTableCell(docTbl.GetRow(rowIdx).GetCell(2), r.DocType);
            WordTableCell(docTbl.GetRow(rowIdx).GetCell(3), r.PrimaryYear?.ToString() ?? "?");
            WordTableCell(docTbl.GetRow(rowIdx).GetCell(4), r.PageCount.ToString());
            WordTableCell(docTbl.GetRow(rowIdx).GetCell(5), r.WordCount.ToString());
        }
        DocRow(1, cmp.Report1);
        DocRow(2, cmp.Report2);

        // ── Section 2 · Comparison Summary ────────────────────────────────
        WordSectionTitle(doc, "2  Comparison Summary");
        int total = cmp.Counts.Values.Sum(); if (total == 0) total = 1;
        var sumTbl = doc.CreateTable(6, 3);
        string[] sumHdrs = { "Metric", "Count", "% of Total" };
        for (int j = 0; j < sumHdrs.Length; j++)
            WordTableCell(sumTbl.GetRow(0).GetCell(j), sumHdrs[j], bold: true, bg: "0F1C3F", fg: "FFFFFF");
        void SumRow2(int idx, string lbl, string key) {
            int cnt = cmp.Counts.GetValueOrDefault(key);
            WordTableCell(sumTbl.GetRow(idx).GetCell(0), lbl);
            WordTableCell(sumTbl.GetRow(idx).GetCell(1), cnt.ToString());
            WordTableCell(sumTbl.GetRow(idx).GetCell(2), Math.Round((double)cnt / total * 100, 1) + "%");
        }
        SumRow2(1, "Same Lines",    "same");
        SumRow2(2, "Changed Lines", "changed");
        SumRow2(3, "Added Lines",   "added");
        SumRow2(4, "Removed Lines", "removed");
        WordTableCell(sumTbl.GetRow(5).GetCell(0), "Number Similarity");
        WordTableCell(sumTbl.GetRow(5).GetCell(1), cmp.NumCmp.SimilarityPct + "%");
        WordTableCell(sumTbl.GetRow(5).GetCell(2), "—");

        // ── Section 3 · Changed Lines ──────────────────────────────────────
        var changed = cmp.ChangedLines.Take(200).ToList();
        WordSectionTitle(doc, $"3  Changed Lines  ({changed.Count})");
        if (changed.Any())
        {
            var chTbl = doc.CreateTable(changed.Count + 1, 4);
            foreach (var (h, j) in new[] { "#", "AFS 1 Line", "AFS 2 Line", "Num Diff" }.Select((h, j) => (h, j)))
                WordTableCell(chTbl.GetRow(0).GetCell(j), h, bold: true, bg: "1C1400", fg: "FFFFFF");
            for (int i = 0; i < changed.Count; i++)
            {
                string bg = i % 2 == 0 ? "FFFBEB" : "FEF9E7";
                WordTableCell(chTbl.GetRow(i+1).GetCell(0), (i+1).ToString(), bg: bg);
                WordTableCell(chTbl.GetRow(i+1).GetCell(1), changed[i].Line1.Truncate(80), bg: bg);
                WordTableCell(chTbl.GetRow(i+1).GetCell(2), changed[i].Line2.Truncate(80), bg: bg);
                WordTableCell(chTbl.GetRow(i+1).GetCell(3),
                    changed[i].NumDiff.Any() ? string.Join(", ", changed[i].NumDiff) : "—", bg: bg);
            }
        }
        else WordBody(doc, "No changed lines detected.", "10B981");

        // ── Section 4 · Added Lines ────────────────────────────────────────
        var added = cmp.AddedLines.Take(200).ToList();
        WordSectionTitle(doc, $"4  Added Lines  ({added.Count} in AFS 2)");
        if (added.Any())
        {
            var addTbl = doc.CreateTable(added.Count + 1, 2);
            WordTableCell(addTbl.GetRow(0).GetCell(0), "#", bold: true, bg: "1E3A5F", fg: "FFFFFF");
            WordTableCell(addTbl.GetRow(0).GetCell(1), "Added Line (AFS 2)", bold: true, bg: "1E3A5F", fg: "FFFFFF");
            for (int i = 0; i < added.Count; i++)
            {
                string bg = i % 2 == 0 ? "EFF6FF" : "DBEAFE";
                WordTableCell(addTbl.GetRow(i+1).GetCell(0), (i+1).ToString(), bg: bg);
                WordTableCell(addTbl.GetRow(i+1).GetCell(1), added[i].Line2.Truncate(120), bg: bg);
            }
        }
        else WordBody(doc, "No added lines.", "10B981");

        // ── Section 5 · Removed Lines ─────────────────────────────────────
        var removed = cmp.RemovedLines.Take(200).ToList();
        WordSectionTitle(doc, $"5  Removed Lines  ({removed.Count} from AFS 1)");
        if (removed.Any())
        {
            var remTbl = doc.CreateTable(removed.Count + 1, 2);
            WordTableCell(remTbl.GetRow(0).GetCell(0), "#", bold: true, bg: "5F1D1D", fg: "FFFFFF");
            WordTableCell(remTbl.GetRow(0).GetCell(1), "Removed Line (AFS 1)", bold: true, bg: "5F1D1D", fg: "FFFFFF");
            for (int i = 0; i < removed.Count; i++)
            {
                string bg = i % 2 == 0 ? "FEF2F2" : "FEE2E2";
                WordTableCell(remTbl.GetRow(i+1).GetCell(0), (i+1).ToString(), bg: bg);
                WordTableCell(remTbl.GetRow(i+1).GetCell(1), removed[i].Line1.Truncate(120), bg: bg);
            }
        }
        else WordBody(doc, "No removed lines.", "10B981");

        // ── Section 6 · Number Comparison ─────────────────────────────────
        WordSectionTitle(doc, "6  Number Comparison");
        var numRows = new[]
        {
            ("AFS 1 unique numbers", cmp.NumCmp.CountAfs1.ToString()),
            ("AFS 2 unique numbers", cmp.NumCmp.CountAfs2.ToString()),
            ("Numbers in both",      cmp.NumCmp.InBoth.Count.ToString()),
            ("Only in AFS 1",        cmp.NumCmp.OnlyInAfs1.Count.ToString()),
            ("Only in AFS 2",        cmp.NumCmp.OnlyInAfs2.Count.ToString()),
            ("Similarity",           cmp.NumCmp.SimilarityPct + "%"),
        };
        var numTbl = doc.CreateTable(numRows.Length + 1, 2);
        WordTableCell(numTbl.GetRow(0).GetCell(0), "Metric", bold: true, bg: "3B0066", fg: "FFFFFF");
        WordTableCell(numTbl.GetRow(0).GetCell(1), "Value",  bold: true, bg: "3B0066", fg: "FFFFFF");
        for (int i = 0; i < numRows.Length; i++)
        {
            string bg = i % 2 == 0 ? "F1F5F9" : "FFFFFF";
            WordTableCell(numTbl.GetRow(i+1).GetCell(0), numRows[i].Item1, bg: bg);
            WordTableCell(numTbl.GetRow(i+1).GetCell(1), numRows[i].Item2, bg: bg);
        }

        // ── Section 7 · Consolidated Auditor Comments ──────────────────────
        WordSectionTitle(doc, "7  Consolidated Auditor Comments");
        var savedComments = comments.Where(kv => !string.IsNullOrWhiteSpace(kv.Value))
                                    .OrderBy(kv => kv.Key).ToList();
        if (savedComments.Any())
        {
            var cmtTbl = doc.CreateTable(savedComments.Count + 1, 2);
            WordTableCell(cmtTbl.GetRow(0).GetCell(0), "Reference", bold: true, bg: "78350F", fg: "FFFFFF");
            WordTableCell(cmtTbl.GetRow(0).GetCell(1), "Comment",   bold: true, bg: "78350F", fg: "FFFFFF");
            for (int i = 0; i < savedComments.Count; i++)
            {
                string bg = i % 2 == 0 ? "FFFBEB" : "FEF3C7";
                WordTableCell(cmtTbl.GetRow(i+1).GetCell(0), savedComments[i].Key, bg: bg);
                WordTableCell(cmtTbl.GetRow(i+1).GetCell(1), savedComments[i].Value, bg: bg);
            }
        }
        else WordBody(doc, "No digital comments recorded.", "64748B");

        WordCommentBox(doc, "overall", "Overall Audit Conclusion / Sign-off", "");

        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        doc.Write(fs);
    }

    // ── Word helpers ──────────────────────────────────────────────────────────

    private static void WordHeading(XWPFDocument doc, string text, int level, string hexColor)
    {
        var p = doc.CreateParagraph();
        p.Style = level == 0 ? "Heading1" : "Heading" + (level + 1);
        var r = p.CreateRun();
        r.SetText(text);
        r.IsBold = true;
        r.SetColor(hexColor);
    }

    private static void WordSectionTitle(XWPFDocument doc, string text)
    {
        var p = doc.CreateParagraph();
        var r = p.CreateRun();
        r.SetText(text);
        r.IsBold = true;
        r.FontSize = 11;
        r.SetColor("5C2D91");
        p.SpacingBefore = 200;
    }

    private static void WordBody(XWPFDocument doc, string text, string hexColor = "000000")
    {
        var p = doc.CreateParagraph();
        var r = p.CreateRun();
        r.SetText(text);
        r.FontSize = 9;
        r.SetColor(hexColor);
    }

    private static void WordTableCell(
        XWPFTableCell cell, string text,
        bool bold = false, string bg = "FFFFFF", string fg = "000000")
    {
        cell.SetColor(bg);
        var p = cell.Paragraphs.Count > 0 ? cell.Paragraphs[0] : cell.AddParagraph();
        var r = p.CreateRun();
        r.SetText(text ?? "—");
        r.IsBold = bold;
        r.FontSize = 8;
        r.SetColor(fg);
    }

    private static void WordCommentBox(
        XWPFDocument doc, string key, string label, string existing)
    {
        var hdr = doc.CreateParagraph();
        var hdrCell = hdr.CreateRun();
        hdrCell.SetText("✏  AUDITOR NOTE — " + label);
        hdrCell.IsBold = true; hdrCell.FontSize = 9; hdrCell.SetColor("FFFFFF");
        // Note: NPOI does not support paragraph shading in a simple run — use table row
        var noteTbl = doc.CreateTable(5, 1);
        for (int i = 0; i < 5; i++)
        {
            noteTbl.GetRow(i).GetCell(0).SetColor("FFFBEB");
            if (i == 0 && !string.IsNullOrEmpty(existing))
            {
                var r = noteTbl.GetRow(0).GetCell(0).Paragraphs[0].CreateRun();
                r.SetText(existing.Truncate(150));
                r.FontSize = 8; r.IsItalic = true; r.SetColor("92400E");
            }
        }
        var sfTbl = doc.CreateTable(1, 3);
        void SfCell(int j, string txt) {
            sfTbl.GetRow(0).GetCell(j).SetColor("FEF3C7");
            var r = sfTbl.GetRow(0).GetCell(j).Paragraphs[0].CreateRun();
            r.SetText(txt); r.FontSize = 7; r.SetColor("6B7280");
        }
        SfCell(0, "Initials: ___________");
        SfCell(1, "Date: ________________");
        SfCell(2, "Ref: _________________");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 5.3 · EXCEL WORKBOOK
    // Uses ClosedXML.  One worksheet per data category.
    // Reference: Python on_save() Excel block in the notebook.
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Generates a multi-sheet Excel workbook and saves it to <paramref name="path"/>.
    /// </summary>
    public void BuildExcel(
        ComparisonResult cmp,
        EngagementDetails eng,
        Dictionary<string, string> comments,
        string path)
    {
        using var wb = new XLWorkbook();

        // ── Summary sheet ──────────────────────────────────────────────────
        var ws = wb.AddWorksheet("Summary");
        string[] sumHeaders = { "same", "changed", "added", "removed", "num_similarity_pct" };
        for (int j = 0; j < sumHeaders.Length; j++) ws.Cell(1, j+1).Value = sumHeaders[j];
        ws.Cell(2, 1).Value = cmp.Counts.GetValueOrDefault("same");
        ws.Cell(2, 2).Value = cmp.Counts.GetValueOrDefault("changed");
        ws.Cell(2, 3).Value = cmp.Counts.GetValueOrDefault("added");
        ws.Cell(2, 4).Value = cmp.Counts.GetValueOrDefault("removed");
        ws.Cell(2, 5).Value = cmp.NumCmp.SimilarityPct;
        StyleHeader(ws, 1, sumHeaders.Length);

        // ── Changed Lines sheet ────────────────────────────────────────────
        if (cmp.ChangedLines.Any())
        {
            var ws2 = wb.AddWorksheet("Changed_Lines");
            string[] ch = { "AFS_1_line", "AFS_2_line", "similarity", "number_changes", "auditor_comment" };
            for (int j = 0; j < ch.Length; j++) ws2.Cell(1, j+1).Value = ch[j];
            var changed = cmp.ChangedLines.ToList();
            for (int i = 0; i < changed.Count; i++)
            {
                ws2.Cell(i+2, 1).Value = changed[i].Line1;
                ws2.Cell(i+2, 2).Value = changed[i].Line2;
                ws2.Cell(i+2, 3).Value = changed[i].Similarity;
                ws2.Cell(i+2, 4).Value = string.Join(", ", changed[i].NumDiff);
                ws2.Cell(i+2, 5).Value = comments.GetValueOrDefault("changed:" + (i+1), "");
            }
            StyleHeader(ws2, 1, ch.Length);
        }

        // ── Added Lines sheet ──────────────────────────────────────────────
        if (cmp.AddedLines.Any())
        {
            var ws3 = wb.AddWorksheet("Added_Lines");
            ws3.Cell(1, 1).Value = "line";
            var added = cmp.AddedLines.ToList();
            for (int i = 0; i < added.Count; i++) ws3.Cell(i+2, 1).Value = added[i].Line2;
            StyleHeader(ws3, 1, 1);
        }

        // ── Removed Lines sheet ────────────────────────────────────────────
        if (cmp.RemovedLines.Any())
        {
            var ws4 = wb.AddWorksheet("Removed_Lines");
            ws4.Cell(1, 1).Value = "line";
            var removed = cmp.RemovedLines.ToList();
            for (int i = 0; i < removed.Count; i++) ws4.Cell(i+2, 1).Value = removed[i].Line1;
            StyleHeader(ws4, 1, 1);
        }

        // ── Number Comparison sheet ────────────────────────────────────────
        var ws5 = wb.AddWorksheet("Number_Comparison");
        ws5.Cell(1, 1).Value = "only_in_AFS1";
        ws5.Cell(1, 2).Value = "only_in_AFS2";
        ws5.Cell(1, 3).Value = "in_both";
        int maxRows = Math.Max(cmp.NumCmp.OnlyInAfs1.Count,
                     Math.Max(cmp.NumCmp.OnlyInAfs2.Count, cmp.NumCmp.InBoth.Count));
        for (int i = 0; i < maxRows; i++)
        {
            if (i < cmp.NumCmp.OnlyInAfs1.Count) ws5.Cell(i+2, 1).Value = cmp.NumCmp.OnlyInAfs1[i];
            if (i < cmp.NumCmp.OnlyInAfs2.Count) ws5.Cell(i+2, 2).Value = cmp.NumCmp.OnlyInAfs2[i];
            if (i < cmp.NumCmp.InBoth.Count)     ws5.Cell(i+2, 3).Value = cmp.NumCmp.InBoth[i];
        }
        StyleHeader(ws5, 1, 3);

        // ── Page Alignment sheet ───────────────────────────────────────────
        var ws6 = wb.AddWorksheet("Page_Alignment");
        string[] pgHdrs = { "pair", "afs1_page", "afs2_page", "align_similarity",
                            "pct_same", "same", "changed", "added", "removed", "auditor_comment" };
        for (int j = 0; j < pgHdrs.Length; j++) ws6.Cell(1, j+1).Value = pgHdrs[j];
        for (int i = 0; i < cmp.PageDiffs.Count; i++)
        {
            var pd = cmp.PageDiffs[i];
            ws6.Cell(i+2, 1).Value  = pd.PairIndex + 1;
            ws6.Cell(i+2, 2).Value  = pd.PageAfs1?.ToString() ?? "";
            ws6.Cell(i+2, 3).Value  = pd.PageAfs2?.ToString() ?? "";
            ws6.Cell(i+2, 4).Value  = pd.AlignSim;
            ws6.Cell(i+2, 5).Value  = pd.PctSame;
            ws6.Cell(i+2, 6).Value  = pd.Same;
            ws6.Cell(i+2, 7).Value  = pd.Changed;
            ws6.Cell(i+2, 8).Value  = pd.Added;
            ws6.Cell(i+2, 9).Value  = pd.Removed;
            ws6.Cell(i+2, 10).Value = comments.GetValueOrDefault("page:" + i, "");
        }
        StyleHeader(ws6, 1, pgHdrs.Length);

        // ── Auditor Comments sheet ─────────────────────────────────────────
        var savedCmts = comments.Where(kv => !string.IsNullOrWhiteSpace(kv.Value)).ToList();
        if (savedCmts.Any())
        {
            var ws7 = wb.AddWorksheet("Auditor_Comments");
            ws7.Cell(1, 1).Value = "key"; ws7.Cell(1, 2).Value = "comment";
            for (int i = 0; i < savedCmts.Count; i++)
            {
                ws7.Cell(i+2, 1).Value = savedCmts[i].Key;
                ws7.Cell(i+2, 2).Value = savedCmts[i].Value;
            }
            StyleHeader(ws7, 1, 2);
        }

        wb.SaveAs(path);
    }

    // Apply SNG purple header styling to the first row of a worksheet.
    private static void StyleHeader(IXLWorksheet ws, int row, int colCount)
    {
        var range = ws.Range(row, 1, row, colCount);
        range.Style.Fill.BackgroundColor   = XLColor.FromHtml("#5C2D91");
        range.Style.Font.FontColor         = XLColor.White;
        range.Style.Font.Bold              = true;
        range.Style.Alignment.WrapText     = true;
        ws.Columns().AdjustToContents();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SECTION 5.4 · PLAIN-TEXT REPORT
    // UTF-8 audit trail in a fixed-width readable format.
    // Reference: Python _build_text_report() in the notebook.
    // ─────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Builds a plain-text working-paper and returns it as a string.
    /// </summary>
    public string BuildText(
        ComparisonResult cmp,
        EngagementDetails eng,
        Dictionary<string, string> comments)
    {
        string ts  = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var sb     = new StringBuilder();
        string sep = new('=', 80);

        sb.AppendLine(sep);
        sb.AppendLine("AFS PDF COMPARISON WORKING PAPER  v4.3");
        sb.AppendLine("SNG Grant Thornton — CAATs Platform");
        sb.AppendLine("Author  : Mamishi Tonny Madire");
        sb.AppendLine("Date    : " + ts);
        sb.AppendLine(sep);
        sb.AppendLine();

        if (!string.IsNullOrEmpty(eng.Client))
        {
            sb.AppendLine("Client       : " + eng.Client);
            sb.AppendLine("Engagement   : " + eng.EngagementNumber);
            sb.AppendLine("FY End       : " + eng.FinancialYearEnd);
            sb.AppendLine("Prepared by  : " + eng.PreparedBy);
            sb.AppendLine("Manager      : " + eng.Manager);
            sb.AppendLine();
        }

        sb.AppendLine("DOCUMENTS");
        sb.AppendLine("  AFS 1: " + cmp.Report1.Filename + " | " + cmp.Report1.PageCount + " pages");
        sb.AppendLine("  AFS 2: " + cmp.Report2.Filename + " | " + cmp.Report2.PageCount + " pages");
        sb.AppendLine();

        var c = cmp.Counts;
        sb.AppendLine("SUMMARY");
        sb.AppendLine($"  Same: {c.GetValueOrDefault("same")}  " +
                      $"Changed: {c.GetValueOrDefault("changed")}  " +
                      $"Added: {c.GetValueOrDefault("added")}  " +
                      $"Removed: {c.GetValueOrDefault("removed")}");
        sb.AppendLine();

        sb.AppendLine("NUMBERS");
        sb.AppendLine($"  AFS1: {cmp.NumCmp.CountAfs1}  AFS2: {cmp.NumCmp.CountAfs2}" +
                      $"  Similarity: {cmp.NumCmp.SimilarityPct}%");
        sb.AppendLine();

        sb.AppendLine("-- CHANGED LINES --");
        int idx = 1;
        foreach (var d in cmp.ChangedLines)
        {
            sb.AppendLine($"[~{idx}] AFS1: {d.Line1}");
            sb.AppendLine($"       AFS2: {d.Line2}");
            if (d.NumDiff.Any()) sb.AppendLine("       NUM: " + string.Join(", ", d.NumDiff));
            string cmt = comments.GetValueOrDefault("changed:" + idx, "");
            if (!string.IsNullOrEmpty(cmt)) sb.AppendLine("       AUDITOR: " + cmt);
            sb.AppendLine();
            idx++;
        }

        if (cmp.AddedLines.Any())
        {
            sb.AppendLine("-- ADDED --");
            idx = 1;
            foreach (var d in cmp.AddedLines) sb.AppendLine($"[+{idx++}] {d.Line2}");
            sb.AppendLine();
        }

        if (cmp.RemovedLines.Any())
        {
            sb.AppendLine("-- REMOVED --");
            idx = 1;
            foreach (var d in cmp.RemovedLines) sb.AppendLine($"[-{idx++}] {d.Line1}");
            sb.AppendLine();
        }

        var savedCmts = comments.Where(kv => !string.IsNullOrWhiteSpace(kv.Value)).ToList();
        if (savedCmts.Any())
        {
            sb.AppendLine("-- AUDITOR COMMENTS --");
            foreach (var (k, v) in savedCmts.OrderBy(x => x.Key))
                sb.AppendLine($"[{k}] {v}");
            sb.AppendLine();
        }

        sb.AppendLine(sep);
        sb.AppendLine("END OF WORKING PAPER");
        sb.AppendLine(sep);
        return sb.ToString();
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 5-INTERNAL · EXTENSION: string.Truncate
// Utility to safely truncate long strings for table cells.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// String extension utility used in the export service.
/// </summary>
public static class StringExtensions
{
    /// <summary>
    /// Returns at most <paramref name="maxLength"/> characters followed by "…"
    /// if truncated.
    /// </summary>
    public static string Truncate(this string? s, int maxLength)
    {
        if (string.IsNullOrEmpty(s)) return string.Empty;
        return s.Length <= maxLength ? s : s[..maxLength] + "…";
    }
}
