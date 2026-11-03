// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  SERVICE — TableDiffService  (#3)                                            ║
// ║  Detects financial tables in AFS page text and diffs them cell-by-cell.      ║
// ║  Line-level diff treats whole rows as text units, causing misalignment when  ║
// ║  a column shifts. This service parses rows×cols and compares each cell,      ║
// ║  producing precise changes per financial line item and column.                ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Globalization;
using System.Text.RegularExpressions;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

public class TableDiffService
{
    // Matches a financial table row: a text label followed by ≥2 whitespace-
    // separated numeric columns (SA format: digits, spaces, commas, brackets, R).
    // e.g. "Property, plant and equipment  1 234 567  (987 654)"
    private static readonly Regex _tableRowRe = new(
        @"^(.{2,60}?)\s{2,}([\d\s,()A-Za-z$€£¥₹%.-]{2,}?)\s{2,}([\d\s,()A-Za-z$€£¥₹%.-]{2,})\s*$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex _numCellRe = new(
        @"^(?=.*\d)[\d\s,.()\-A-Z$€£¥₹%]+$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // ── Public entry point ────────────────────────────────────────────────────

    /// <summary>
    /// For each matched page pair, extracts tables and diffs them cell-by-cell.
    /// Returns a flat list of <see cref="TableCellDiff"/> covering every changed cell.
    /// </summary>
    public List<TableCellDiff> Diff(ComparisonResult cmp)
    {
        var results = new List<TableCellDiff>();

        foreach (var pd in cmp.PageDiffs.Where(p => !p.Unmatched && p.PageAfs1.HasValue && p.PageAfs2.HasValue))
        {
            string page1 = SafePage(cmp.Report1.Pages, pd.PageAfs1!.Value - 1);
            string page2 = SafePage(cmp.Report2.Pages, pd.PageAfs2!.Value - 1);

            var table1 = ExtractTable(page1);
            var table2 = ExtractTable(page2);
            if (table1.Count == 0 && table2.Count == 0) continue;

            results.AddRange(DiffTables(table1, table2, pd));
        }

        return results;
    }

    // ── Table extraction ──────────────────────────────────────────────────────

    private record TableRow(string Label, string Col1, string Col2);

    private static List<TableRow> ExtractTable(string pageText)
    {
        var rows = new List<TableRow>();
        foreach (var line in pageText.Split('\n'))
        {
            var m = _tableRowRe.Match(line.Trim());
            if (!m.Success) continue;

            string label = m.Groups[1].Value.Trim();
            string c1    = m.Groups[2].Value.Trim();
            string c2    = m.Groups[3].Value.Trim();

            // Both numeric columns must look like numbers
            if (!_numCellRe.IsMatch(c1) || !_numCellRe.IsMatch(c2)) continue;
            if (label.Length < 3 || label.Length > 60) continue;

            rows.Add(new TableRow(label, c1, c2));
        }
        return rows;
    }

    // ── Table diff ────────────────────────────────────────────────────────────

    private static List<TableCellDiff> DiffTables(
        List<TableRow> table1, List<TableRow> table2, PageDiffResult pd)
    {
        var diffs = new List<TableCellDiff>();

        // Index table2 rows by label for O(1) lookup
        var idx2 = new Dictionary<string, TableRow>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in table2) idx2.TryAdd(NormLabel(r.Label), r);

        int row = 0;
        foreach (var r1 in table1)
        {
            row++;
            string normLabel = NormLabel(r1.Label);

            if (!idx2.TryGetValue(normLabel, out var r2))
            {
                // Entire row removed
                diffs.Add(new TableCellDiff
                {
                    Row      = row,
                    Col      = 1,
                    Label    = r1.Label,
                    ColHeader = "Current Year",
                    ValueAfs1 = r1.Col1,
                    ValueAfs2 = "",
                    Status   = "removed",
                    PageAfs1 = pd.PageAfs1,
                    PageAfs2 = pd.PageAfs2,
                    PairIndex = pd.PairIndex,
                });
                continue;
            }

            // Compare column 1 (current year)
            var d1 = CellDiff(r1.Label, "Current Year", r1.Col1, r2.Col1, row, 1, pd);
            if (d1 != null) diffs.Add(d1);

            // Compare column 2 (prior year / comparative)
            var d2 = CellDiff(r1.Label, "Prior Year", r1.Col2, r2.Col2, row, 2, pd);
            if (d2 != null) diffs.Add(d2);
        }

        // Rows only in table2 = added
        var seen = new HashSet<string>(table1.Select(r => NormLabel(r.Label)), StringComparer.OrdinalIgnoreCase);
        int addRow = row;
        foreach (var r2 in table2.Where(r => !seen.Contains(NormLabel(r.Label))))
        {
            addRow++;
            diffs.Add(new TableCellDiff
            {
                Row       = addRow,
                Col       = 1,
                Label     = r2.Label,
                ColHeader = "Current Year",
                ValueAfs1 = "",
                ValueAfs2 = r2.Col1,
                Status    = "added",
                PageAfs1  = pd.PageAfs1,
                PageAfs2  = pd.PageAfs2,
                PairIndex = pd.PairIndex,
            });
        }

        return diffs;
    }

    private static TableCellDiff? CellDiff(
        string label, string colHeader,
        string v1, string v2,
        int row, int col, PageDiffResult pd)
    {
        string n1 = AuditTextNormalizer.NormalizeFinancialValue(v1);
        string n2 = AuditTextNormalizer.NormalizeFinancialValue(v2);
        if (AuditTextNormalizer.FinancialValuesEquivalent(v1, v2)) return null;

        double? pct = null;
        if (decimal.TryParse(n1.Split('|')[0].TrimStart('-'), NumberStyles.Number, CultureInfo.InvariantCulture, out decimal d1) &&
            decimal.TryParse(n2.Split('|')[0].TrimStart('-'), NumberStyles.Number, CultureInfo.InvariantCulture, out decimal d2) && d1 != 0)
            pct = Math.Round((double)Math.Abs(d2 - d1) / (double)Math.Abs(d1) * 100, 1);

        return new TableCellDiff
        {
            Row       = row,
            Col       = col,
            Label     = label,
            ColHeader = colHeader,
            ValueAfs1 = v1,
            ValueAfs2 = v2,
            Status    = "changed",
            PctChange = pct,
            PageAfs1  = pd.PageAfs1,
            PageAfs2  = pd.PageAfs2,
            PairIndex = pd.PairIndex,
        };
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string SafePage(List<string> pages, int idx)
        => idx >= 0 && idx < pages.Count ? pages[idx] : "";

    private static string NormLabel(string s)
        => Regex.Replace(s.ToLowerInvariant().Trim(), @"\s+", " ");

}
