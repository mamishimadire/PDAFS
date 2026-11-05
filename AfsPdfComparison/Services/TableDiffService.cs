using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

/// <summary>
/// Reconstructs table-like rows from extracted page text, matches rows by label
/// meaning and maps columns by period headers when available. The resulting cell
/// findings are consumed by the same snapshot renderer as paragraph findings.
/// </summary>
public class TableDiffService
{
    private static readonly Regex ColumnSeparator = new(@"\s{2,}", RegexOptions.Compiled);
    private static readonly Regex FinancialToken = new(
        @"(?<![\p{L}\p{N}])(?:\(?\s*(?:ZAR|USD|EUR|GBP|AUD|CAD|JPY|CHF|R|[$€£¥₹])\s*)?[-−+]?\s*\d+(?:[ ,]\d{3})*(?:[.,]\d+)?\s*%?\)?(?![\p{L}\p{N}])",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex YearToken = new(@"\b(?:19|20)\d{2}\b", RegexOptions.Compiled);
    private static readonly Regex ReferenceHeader = new(
        @"\b(?:page|pages|folio|index|contents)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex DotLeaderPageReference = new(
        @"(?:\.{2,}|…+|·{3,})\s*\d+\s*$", RegexOptions.Compiled);
    private static readonly Regex IntegerReference = new(@"^\s*\d+\s*$", RegexOptions.Compiled);
    private static readonly Regex AggregateLabel = new(
        @"\b(?:total|subtotal|net|balance)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex CurrencyToken = new(
        @"\b(?:ZAR|USD|EUR|GBP|AUD|CAD|JPY|CHF)\b|[$€£¥₹]|(?<![\p{L}])R(?=\s*[-−(\d])",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex WordToken = new(@"[\p{L}\p{N}]+", RegexOptions.Compiled);

    private sealed record TableRow(
        string Label,
        IReadOnlyList<string> Values,
        IReadOnlyList<string> ColumnHeaders,
        string SourceText,
        int SourceRow,
        string Heading,
        double ExtractionConfidence);

    private sealed record RowMatch(TableRow? First, TableRow? Second, double Confidence);

    private sealed record TablePagePair(
        string TableId,
        string Heading,
        PageDiffResult PagePair,
        IReadOnlyList<TableRow> FirstRows,
        IReadOnlyList<TableRow> SecondRows,
        IReadOnlyList<RowMatch> RowMatches);

    private sealed record UnmatchedTableRow(TablePagePair Table, TableRow Row);

    public List<TableCellDiff> Diff(ComparisonResult cmp)
    {
        var results = new List<TableCellDiff>();
        var tablePairs = new List<TablePagePair>();

        foreach (var pagePair in cmp.PageDiffs.Where(page =>
                     !page.Unmatched && page.PageAfs1.HasValue && page.PageAfs2.HasValue))
        {
            string firstText = SafePage(cmp.Report1.Pages, pagePair.PageAfs1!.Value - 1);
            string secondText = SafePage(cmp.Report2.Pages, pagePair.PageAfs2!.Value - 1);
            var firstRows = ExtractRows(firstText);
            var secondRows = ExtractRows(secondText);
            if (firstRows.Count == 0 && secondRows.Count == 0)
                continue;

            string heading = firstRows.FirstOrDefault()?.Heading ?? secondRows.FirstOrDefault()?.Heading ?? "";
            string tableId = BuildTableId(pagePair.PairIndex, heading);
            var rowMatches = MatchRows(firstRows, secondRows);
            tablePairs.Add(new TablePagePair(tableId, heading, pagePair, firstRows, secondRows, rowMatches));
            var pageFindings = DiffRows(tableId, rowMatches, pagePair).ToList();
            if (IsReferenceIndex(firstRows, secondRows))
            {
                int? pageOffset = GetSystematicReferenceOffset(rowMatches);
                string structuralType = pageOffset.HasValue ? "pagination-shift" : "reference-renumbering";
                foreach (var finding in pageFindings.Where(finding =>
                             IsPageReferenceValue(finding.ValueAfs1) || IsPageReferenceValue(finding.ValueAfs2)))
                {
                    finding.DifferenceType = "pagination-reference";
                    finding.StructuralChangeType = structuralType;
                    finding.SemanticClassification = "structural-change";
                    finding.HighlightEligible = false;
                    finding.Confidence = Math.Min(finding.Confidence, 0.95);
                }

                if (pageFindings.Any(finding => finding.DifferenceType == "pagination-reference") &&
                    pagePair.StructuralChangeType == "none")
                {
                    pagePair.StructuralChangeType = structuralType;
                    pagePair.StructuralPageOffset = pageOffset;
                }
            }

            results.AddRange(pageFindings);
        }

        ReconcileRowsAcrossPageBreaks(tablePairs, results);
        cmp.CalculationRelationships = ValidateCalculations(tablePairs, results);
        foreach (var finding in results)
        {
            finding.OriginalDocumentId = cmp.Report1.Filename;
            finding.RevisedDocumentId = cmp.Report2.Filename;
        }
        return results;
    }

    private static bool IsReferenceIndex(
        IReadOnlyList<TableRow> firstRows,
        IReadOnlyList<TableRow> secondRows)
    {
        var rows = firstRows.Concat(secondRows).ToList();
        if (rows.Count < 3 || rows.Any(row => row.ColumnHeaders.Count > 0))
            return false;

        var referenceRows = rows.Where(row => row.Values.Count > 0 &&
                                              row.Values.All(IsPageReferenceValue)).ToList();
        if (referenceRows.Count < 3 || referenceRows.Count / (double)rows.Count < 0.75)
            return false;

        bool explicitReferenceContext = rows.Any(row => ReferenceHeader.IsMatch(row.Heading));
        bool dotLeaderContext = rows.Any(row => DotLeaderPageReference.IsMatch(row.SourceText));
        return explicitReferenceContext || dotLeaderContext ||
               IsMostlyAscendingPageSequence(firstRows) && IsMostlyAscendingPageSequence(secondRows);
    }

    private static bool IsMostlyAscendingPageSequence(IReadOnlyList<TableRow> rows)
    {
        var values = rows.OrderBy(row => row.SourceRow)
            .Select(row => row.Values.Count == 1 && int.TryParse(row.Values[0], NumberStyles.None,
                CultureInfo.InvariantCulture, out int value) ? value : (int?)null)
            .Where(value => value.HasValue)
            .Select(value => value!.Value)
            .ToList();
        if (values.Count < 3)
            return false;

        int orderedSteps = values.Zip(values.Skip(1), (first, second) => second >= first ? 1 : 0).Sum();
        return orderedSteps >= Math.Ceiling((values.Count - 1) * 0.7);
    }

    private static bool IsPageReferenceValue(string value) =>
        IntegerReference.IsMatch(value ?? "") &&
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out int number) && number > 0;

    private static int? GetSystematicReferenceOffset(IReadOnlyList<RowMatch> rowMatches)
    {
        var offsets = rowMatches
            .Where(match => match.First != null && match.Second != null &&
                            match.First.Values.Count == match.Second.Values.Count &&
                            match.First.Values.All(IsPageReferenceValue) &&
                            match.Second.Values.All(IsPageReferenceValue))
            .SelectMany(match => Enumerable.Range(0, match.First!.Values.Count)
                .Select(index => int.TryParse(match.Second!.Values[index], NumberStyles.None, CultureInfo.InvariantCulture, out int revised) &&
                                 int.TryParse(match.First.Values[index], NumberStyles.None, CultureInfo.InvariantCulture, out int original)
                    ? revised - original
                    : (int?)null))
            .Where(offset => offset.HasValue)
            .Select(offset => offset!.Value)
            .ToList();
        if (offsets.Count < 2)
            return null;

        var dominant = offsets.GroupBy(offset => offset)
            .OrderByDescending(group => group.Count())
            .ThenBy(group => group.Key)
            .First();
        return dominant.Key != 0 && dominant.Count() / (double)offsets.Count >= 0.6
            ? dominant.Key
            : null;
    }

    private static List<TableRow> ExtractRows(string pageText)
    {
        string[] lines = (pageText ?? "").Split('\n');
        var candidates = new List<(TableRow Row, bool Strong)>();
        var amountOnlyLines = new List<(int SourceRow, string Heading, string SourceText, IReadOnlyList<string> Values)>();
        string heading = "";

        for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++)
        {
            string source = lines[lineIndex].Trim();
            if (source.Length == 0)
                continue;

            if (IsLikelyPeriodHeader(source))
            {
                heading = source;
                continue;
            }

            if (TryExtractSpacedRow(source, lineIndex + 1, heading, lines, out var spacedRow))
            {
                candidates.Add((spacedRow, true));
                continue;
            }

            if (TryExtractAmountOnlyValues(source, out var amountOnlyValues))
            {
                amountOnlyLines.Add((lineIndex + 1, heading, source, amountOnlyValues));
                continue;
            }

            if (TryExtractTrailingValues(source, lineIndex + 1, heading, lines, out var row, out bool strong))
            {
                candidates.Add((row, strong));
                continue;
            }

            if (LooksLikeHeading(source))
                heading = source;
        }

        // Single-number lines can also be narrative sentences. Retain them when
        // the extraction has clear column spacing/formatting or they are part of
        // a run of adjacent financial rows.
        var retained = new HashSet<TableRow>();
        foreach (var (row, strong) in candidates.Where(candidate => candidate.Strong))
            retained.Add(row);

        for (int i = 0; i < candidates.Count;)
        {
            int end = i + 1;
            while (end < candidates.Count && candidates[end].Row.SourceRow - candidates[end - 1].Row.SourceRow <= 2)
                end++;

            if (end - i >= 2)
            {
                for (int j = i; j < end; j++)
                    retained.Add(candidates[j].Row);
            }

            i = end;
        }

        var retainedRows = candidates.Select(candidate => candidate.Row)
            .Where(retained.Contains)
            .ToList();

        foreach (var amountLine in amountOnlyLines)
        {
            var priorRows = retainedRows
                .Where(row => row.SourceRow < amountLine.SourceRow)
                .OrderBy(row => row.SourceRow)
                .ToList();

            var subtotalHeaders = FindColumnHeaders(lines, amountLine.SourceRow - 1, amountLine.Values.Count);
            var sameHeadingRows = priorRows
                .Where(row => NormalizeLabel(row.Heading) == NormalizeLabel(amountLine.Heading) ||
                              string.IsNullOrWhiteSpace(row.Heading) ||
                              string.IsNullOrWhiteSpace(amountLine.Heading))
                .ToList();
            var components = FindAdditiveSubtotalComponents(sameHeadingRows, amountLine.Values, subtotalHeaders)
                ?? FindAdditiveSubtotalComponents(priorRows, amountLine.Values, subtotalHeaders);
            if (components == null || components.Count < 2)
                continue;

            string subtotalLabel = "Subtotal of " + string.Join(" + ", components.Select(row => row.Label));
            retainedRows.Add(new TableRow(
                subtotalLabel,
                amountLine.Values,
                subtotalHeaders,
                amountLine.SourceText,
                amountLine.SourceRow,
                amountLine.Heading,
                0.92));
        }

        return retainedRows.OrderBy(row => row.SourceRow).ToList();
    }

    private static List<TableRow>? FindAdditiveSubtotalComponents(
        IReadOnlyList<TableRow> priorRows,
        IReadOnlyList<string> subtotalValues,
        IReadOnlyList<string> subtotalHeaders)
    {
        for (int componentCount = 2; componentCount <= priorRows.Count; componentCount++)
        {
            var components = priorRows.Skip(priorRows.Count - componentCount).ToList();
            bool matchesEveryColumn = true;
            for (int column = 0; column < subtotalValues.Count; column++)
            {
                if (!TryParseAmount(subtotalValues[column], out decimal expected))
                {
                    matchesEveryColumn = false;
                    break;
                }

                decimal sum = 0;
                foreach (var component in components)
                {
                    int componentColumn = FindMatchingColumnIndex(component, column, subtotalValues.Count, subtotalHeaders);
                    if (componentColumn < 0 || !TryParseAmount(component.Values[componentColumn], out decimal value))
                    {
                        matchesEveryColumn = false;
                        break;
                    }
                    sum += value;
                }

                if (!matchesEveryColumn || sum != expected)
                {
                    matchesEveryColumn = false;
                    break;
                }
            }

            if (matchesEveryColumn)
                return components;
        }

        return null;
    }

    private static int FindMatchingColumnIndex(
        TableRow row,
        int subtotalColumn,
        int subtotalColumnCount,
        IReadOnlyList<string> subtotalHeaders)
    {
        if (subtotalHeaders.Count == subtotalColumnCount && row.ColumnHeaders.Count == row.Values.Count)
        {
            string targetHeader = subtotalHeaders[subtotalColumn];
            int matchedIndex = Enumerable.Range(0, row.ColumnHeaders.Count)
                .FirstOrDefault(index => string.Equals(
                    row.ColumnHeaders[index], targetHeader, StringComparison.OrdinalIgnoreCase), -1);
            if (matchedIndex >= 0)
                return matchedIndex;
        }

        return row.Values.Count == subtotalColumnCount && row.ColumnHeaders.Count != row.Values.Count
            ? subtotalColumn
            : -1;
    }

    private static bool TryExtractAmountOnlyValues(string source, out IReadOnlyList<string> values)
    {
        string text = source ?? "";
        values = Array.Empty<string>();
        var matches = FinancialToken.Matches(text).Cast<Match>().ToList();
        if (matches.Count < 2)
            return false;

        var remainder = new StringBuilder(text.Length);
        int last = 0;
        foreach (var match in matches)
        {
            remainder.Append(text, last, match.Index - last);
            remainder.Append(' ', match.Length);
            last = match.Index + match.Length;
        }
        remainder.Append(text, last, text.Length - last);
        if (!string.IsNullOrWhiteSpace(remainder.ToString()))
            return false;

        var parsed = matches.Select(match => match.Value.Trim()).ToList();
        if (parsed.Any(value => AuditTextNormalizer.ExtractNormalizedNumbers(value).Count != 1))
            return false;

        values = parsed;
        return true;
    }

    private static void ReconcileRowsAcrossPageBreaks(
        IReadOnlyList<TablePagePair> tables,
        List<TableCellDiff> findings)
    {
        var removedRows = tables.SelectMany(table => table.RowMatches
                .Where(match => match.First != null && match.Second == null)
                .Select(match => new UnmatchedTableRow(table, match.First!)))
            .ToList();
        var addedRows = tables.SelectMany(table => table.RowMatches
                .Where(match => match.First == null && match.Second != null)
                .Select(match => new UnmatchedTableRow(table, match.Second!)))
            .ToList();

        var candidates = (from removed in removedRows
                          from added in addedRows
                          let rowScore = RowSimilarity(removed.Row, added.Row)
                          let headingScore = PageAlignmentService.SequenceSimilarity(
                              NormalizeLabel(removed.Row.Heading), NormalizeLabel(added.Row.Heading))
                          let sameColumns = HaveCompatibleColumns(removed.Row, added.Row)
                          where sameColumns && rowScore >= 0.80 &&
                                (headingScore >= 0.70 || NormalizeLabel(removed.Row.Label) == NormalizeLabel(added.Row.Label))
                          select (Removed: removed, Added: added, Score: (0.8 * rowScore) + (0.2 * headingScore)))
            .ToList();

        var confident = candidates.Where(candidate =>
                candidate.Score - candidates.Where(other => ReferenceEquals(other.Removed.Row, candidate.Removed.Row) &&
                                                           !ReferenceEquals(other.Added.Row, candidate.Added.Row))
                    .Select(other => other.Score).DefaultIfEmpty(0).Max() >= 0.08 &&
                candidate.Score - candidates.Where(other => ReferenceEquals(other.Added.Row, candidate.Added.Row) &&
                                                           !ReferenceEquals(other.Removed.Row, candidate.Removed.Row))
                    .Select(other => other.Score).DefaultIfEmpty(0).Max() >= 0.08)
            .OrderByDescending(candidate => candidate.Score)
            .ToList();

        var consumedRemoved = new HashSet<string>(StringComparer.Ordinal);
        var consumedAdded = new HashSet<string>(StringComparer.Ordinal);
        foreach (var candidate in confident)
        {
            var removed = candidate.Removed;
            var added = candidate.Added;
            string removedKey = $"{removed.Table.TableId}|{removed.Row.SourceRow}";
            string addedKey = $"{added.Table.TableId}|{added.Row.SourceRow}";
            if (consumedRemoved.Contains(removedKey) || consumedAdded.Contains(addedKey))
                continue;
            consumedRemoved.Add(removedKey);
            consumedAdded.Add(addedKey);

            findings.RemoveAll(finding =>
                finding.TableId == removed.Table.TableId &&
                finding.Status == "removed" &&
                finding.SourceRowAfs1 == removed.Row.SourceRow);
            findings.RemoveAll(finding =>
                finding.TableId == added.Table.TableId &&
                finding.Status == "added" &&
                finding.SourceRowAfs2 == added.Row.SourceRow);

            var relocatedPair = new PageDiffResult
            {
                PairIndex = removed.Table.PagePair.PairIndex,
                PageAfs1 = removed.Table.PagePair.PageAfs1,
                PageAfs2 = added.Table.PagePair.PageAfs2,
                StructuralChangeType = "section-relocation",
            };
            var rowMatch = new RowMatch(removed.Row, added.Row, Math.Min(candidate.Score, 0.96));
            var matchedColumns = MatchColumns(removed.Row, added.Row)
                .Select(column => new
                {
                    Column = column,
                    FirstValue = column.FirstIndex.HasValue ? removed.Row.Values[column.FirstIndex.Value] : "",
                    SecondValue = column.SecondIndex.HasValue ? added.Row.Values[column.SecondIndex.Value] : "",
                })
                .ToList();
            var changedColumns = matchedColumns.Where(item =>
                    item.FirstValue.Length == 0 || item.SecondValue.Length == 0 ||
                    !AuditTextNormalizer.FinancialValuesEquivalent(item.FirstValue, item.SecondValue))
                .ToList();

            if (changedColumns.Count == 0)
            {
                var finding = CreateCellFinding(
                    removed.Table.TableId,
                    rowMatch,
                    1,
                    string.Join(" | ", removed.Row.Values),
                    string.Join(" | ", added.Row.Values),
                    "changed",
                    relocatedPair,
                    "row");
                finding.DifferenceType = "financial-row-relocated";
                finding.StructuralChangeType = "section-relocation";
                finding.SemanticClassification = "structural-change";
                finding.HighlightEligible = false;
                findings.Add(finding);
                continue;
            }

            foreach (var item in changedColumns)
            {
                var finding = CreateCellFinding(
                    removed.Table.TableId,
                    rowMatch,
                    item.Column.Index,
                    item.FirstValue,
                    item.SecondValue,
                    item.FirstValue.Length == 0 ? "added" : item.SecondValue.Length == 0 ? "removed" : "changed",
                    relocatedPair,
                    item.Column.Identity);
                finding.StructuralChangeType = "section-relocation";
                findings.Add(finding);
            }
        }
    }

    private static bool TryExtractSpacedRow(
        string source,
        int sourceRow,
        string heading,
        string[] lines,
        out TableRow row)
    {
        row = null!;
        string[] fields = ColumnSeparator.Split(source)
            .Select(field => field.Trim())
            .Where(field => field.Length > 0)
            .ToArray();
        if (fields.Length < 2)
            return false;

        int valueStart = fields.Length;
        while (valueStart > 1 && IsSingleFinancialCell(fields[valueStart - 1]))
            valueStart--;
        if (valueStart == fields.Length)
            return false;

        string label = string.Join(" ", fields.Take(valueStart));
        var values = fields.Skip(valueStart).ToList();
        if (!IsPlausibleRowLabel(label) || values.Any(value => !IsSingleFinancialCell(value)))
            return false;

        row = new TableRow(label, values, FindColumnHeaders(lines, sourceRow - 1, values.Count),
            source, sourceRow, heading, 0.99);
        return true;
    }

    private static bool TryExtractTrailingValues(
        string source,
        int sourceRow,
        string heading,
        string[] lines,
        out TableRow row,
        out bool strong)
    {
        row = null!;
        strong = false;
        string scanText = Regex.Replace(source,
            @"\b(?:note|ref(?:erence)?)\s+\d+(?:\.\d+)*\b",
            match => new string(' ', match.Length),
            RegexOptions.IgnoreCase);
        var matches = FinancialToken.Matches(scanText).Cast<Match>().ToList();
        if (matches.Count == 0)
            return false;

        int firstValue = matches.Count - 1;
        int end = source.Length;
        while (firstValue >= 0)
        {
            Match match = matches[firstValue];
            string between = source[(match.Index + match.Length)..end];
            if (between.Any(char.IsLetter) || AuditTextNormalizer.ExtractNormalizedNumbers(match.Value).Count != 1)
                break;

            end = match.Index;
            firstValue--;
        }
        firstValue++;
        if (firstValue >= matches.Count)
            return false;

        string label = source[..matches[firstValue].Index].Trim();
        var values = matches.Skip(firstValue)
            .Select(match => source.Substring(match.Index, match.Length).Trim())
            .ToList();
        if (!IsPlausibleRowLabel(label))
            return false;

        strong = values.Count >= 2 ||
                 values.Any(value => value.Any(ch => ch is ',' or '(' or ')' or '%' or '$' or '€' or '£' or '¥' or '₹')) ||
                 Regex.IsMatch(source, @"\s{2,}");
        row = new TableRow(label, values, FindColumnHeaders(lines, sourceRow - 1, values.Count),
            source, sourceRow, heading, 0.92);
        return true;
    }

    private static IReadOnlyList<string> FindColumnHeaders(string[] lines, int rowIndex, int valueCount)
    {
        if (valueCount < 2)
            return Array.Empty<string>();

        for (int index = Math.Min(rowIndex - 1, lines.Length - 1); index >= Math.Max(0, rowIndex - 8); index--)
        {
            var years = YearToken.Matches(lines[index]).Select(match => match.Value).ToList();
            if (years.Count == valueCount && years.Distinct(StringComparer.Ordinal).Count() == valueCount)
                return years;

            int referenceColumns = valueCount - years.Count;
            if (years.Count > 0 && years.Distinct(StringComparer.Ordinal).Count() == years.Count &&
                referenceColumns > 0 &&
                Regex.IsMatch(lines[index], @"\b(?:notes?|refs?|references?)\b", RegexOptions.IgnoreCase))
            {
                return Enumerable.Range(1, referenceColumns)
                    .Select(column => $"reference:{column}")
                    .Concat(years)
                    .ToList();
            }
        }

        return Array.Empty<string>();
    }

    private static List<RowMatch> MatchRows(IReadOnlyList<TableRow> firstRows, IReadOnlyList<TableRow> secondRows)
    {
        var matches = new List<RowMatch>();
        var usedSecond = new HashSet<int>();
        var firstUnmatched = new List<int>();

        for (int firstIndex = 0; firstIndex < firstRows.Count; firstIndex++)
        {
            string key = NormalizeLabel(firstRows[firstIndex].Label);
            int exact = Enumerable.Range(0, secondRows.Count)
                .FirstOrDefault(index => !usedSecond.Contains(index) && NormalizeLabel(secondRows[index].Label) == key, -1);
            if (exact >= 0)
            {
                usedSecond.Add(exact);
                matches.Add(new RowMatch(firstRows[firstIndex], secondRows[exact], 1.0));
            }
            else
            {
                firstUnmatched.Add(firstIndex);
            }
        }

        var candidates = (from firstIndex in firstUnmatched
                          from secondIndex in Enumerable.Range(0, secondRows.Count).Where(index => !usedSecond.Contains(index))
                          let score = RowSimilarity(firstRows[firstIndex], secondRows[secondIndex])
                          where score >= 0.72
                          select (First: firstIndex, Second: secondIndex, Score: score))
            .ToList();

        var confident = candidates.Where(candidate =>
            candidate.Score - candidates.Where(other => other.First == candidate.First && other.Second != candidate.Second)
                .Select(other => other.Score).DefaultIfEmpty(0).Max() >= 0.06 &&
            candidate.Score - candidates.Where(other => other.Second == candidate.Second && other.First != candidate.First)
                .Select(other => other.Score).DefaultIfEmpty(0).Max() >= 0.06)
            .OrderByDescending(candidate => candidate.Score)
            .ToList();

        var usedFirstFuzzy = new HashSet<int>();
        foreach (var candidate in confident)
        {
            if (usedFirstFuzzy.Contains(candidate.First) || usedSecond.Contains(candidate.Second))
                continue;
            usedFirstFuzzy.Add(candidate.First);
            usedSecond.Add(candidate.Second);
            matches.Add(new RowMatch(firstRows[candidate.First], secondRows[candidate.Second], candidate.Score));
        }

        foreach (int index in firstUnmatched.Where(index => !usedFirstFuzzy.Contains(index)))
            matches.Add(new RowMatch(firstRows[index], null, firstRows[index].ExtractionConfidence));
        foreach (int index in Enumerable.Range(0, secondRows.Count).Where(index => !usedSecond.Contains(index)))
            matches.Add(new RowMatch(null, secondRows[index], secondRows[index].ExtractionConfidence));

        return matches.OrderBy(match => match.First?.SourceRow ?? int.MaxValue)
            .ThenBy(match => match.Second?.SourceRow ?? int.MaxValue)
            .ToList();
    }

    private static bool HaveCompatibleColumns(TableRow first, TableRow second)
    {
        if (first.Values.Count != second.Values.Count)
            return false;

        var firstKeys = ColumnKeys(first);
        var secondKeys = ColumnKeys(second);
        bool firstIsPositional = firstKeys.All(key => key.StartsWith("column:", StringComparison.Ordinal));
        bool secondIsPositional = secondKeys.All(key => key.StartsWith("column:", StringComparison.Ordinal));
        return firstIsPositional || secondIsPositional ||
               firstKeys.SequenceEqual(secondKeys, StringComparer.OrdinalIgnoreCase);
    }

    private static double RowSimilarity(TableRow first, TableRow second)
    {
        var firstTokens = LabelTokens(first.Label);
        var secondTokens = LabelTokens(second.Label);
        if (firstTokens.Count == 0 || secondTokens.Count == 0)
            return 0;

        int intersection = firstTokens.Intersect(secondTokens, StringComparer.OrdinalIgnoreCase).Count();
        double jaccard = (double)intersection / firstTokens.Union(secondTokens, StringComparer.OrdinalIgnoreCase).Count();
        double containment = (double)intersection / Math.Min(firstTokens.Count, secondTokens.Count);
        double text = PageAlignmentService.SequenceSimilarity(NormalizeLabel(first.Label), NormalizeLabel(second.Label));
        double valueOverlap = 0;
        if (first.Values.Count > 0 && second.Values.Count > 0)
        {
            int shared = first.Values.Select(AuditTextNormalizer.NormalizeFinancialValue)
                .Intersect(second.Values.Select(AuditTextNormalizer.NormalizeFinancialValue), StringComparer.Ordinal)
                .Count();
            valueOverlap = (double)shared / Math.Max(first.Values.Count, second.Values.Count);
        }

        return (0.40 * jaccard) + (0.25 * containment) + (0.25 * text) + (0.10 * valueOverlap);
    }

    private static IEnumerable<TableCellDiff> DiffRows(
        string tableId,
        IReadOnlyList<RowMatch> rowMatches,
        PageDiffResult pagePair)
    {
        foreach (var match in rowMatches)
        {
            var first = match.First;
            var second = match.Second;
            if (first == null || second == null)
            {
                var row = first ?? second!;
                for (int index = 0; index < row.Values.Count; index++)
                {
                    yield return CreateCellFinding(tableId, match, index, first == null ? "" : row.Values[index],
                        second == null ? "" : row.Values[index], first == null ? "added" : "removed", pagePair);
                }
                continue;
            }

            var columnPairs = MatchColumns(first, second);
            foreach (var column in columnPairs)
            {
                string firstValue = column.FirstIndex.HasValue ? first.Values[column.FirstIndex.Value] : "";
                string secondValue = column.SecondIndex.HasValue ? second.Values[column.SecondIndex.Value] : "";
                if (firstValue.Length > 0 && secondValue.Length > 0 &&
                    AuditTextNormalizer.FinancialValuesEquivalent(firstValue, secondValue))
                    continue;

                string status = firstValue.Length == 0 ? "added" : secondValue.Length == 0 ? "removed" : "changed";
                yield return CreateCellFinding(tableId, match, column.Index, firstValue, secondValue, status, pagePair,
                    column.Identity);
            }
        }
    }

    private sealed record ColumnMatch(int Index, int? FirstIndex, int? SecondIndex, string Identity);

    private static List<ColumnMatch> MatchColumns(TableRow first, TableRow second)
    {
        var firstKeys = ColumnKeys(first);
        var secondKeys = ColumnKeys(second);
        if (firstKeys.Count > 0 && secondKeys.Count > 0 &&
            firstKeys.All(key => !key.StartsWith("column:", StringComparison.Ordinal)) &&
            secondKeys.All(key => !key.StartsWith("column:", StringComparison.Ordinal)))
        {
            var keys = firstKeys.Concat(secondKeys).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            return keys.Select((key, index) => new ColumnMatch(index + 1,
                firstKeys.FindIndex(item => string.Equals(item, key, StringComparison.OrdinalIgnoreCase)) is var i1 && i1 >= 0 ? i1 : null,
                secondKeys.FindIndex(item => string.Equals(item, key, StringComparison.OrdinalIgnoreCase)) is var i2 && i2 >= 0 ? i2 : null,
                key)).ToList();
        }

        return Enumerable.Range(0, Math.Max(first.Values.Count, second.Values.Count))
            .Select(index => new ColumnMatch(index + 1,
                index < first.Values.Count ? index : null,
                index < second.Values.Count ? index : null,
                $"column:{index + 1}"))
            .ToList();
    }

    private static List<string> ColumnKeys(TableRow row) => row.Values
        .Select((_, index) => index < row.ColumnHeaders.Count ? row.ColumnHeaders[index] : $"column:{index + 1}")
        .ToList();

    private static TableCellDiff CreateCellFinding(
        string tableId,
        RowMatch match,
        int columnIndex,
        string firstValue,
        string secondValue,
        string status,
        PageDiffResult pagePair,
        string? columnIdentity = null)
    {
        var firstRow = match.First;
        var secondRow = match.Second;
        var row = firstRow ?? secondRow!;
        string normalizedLabel = NormalizeLabel(row.Label);
        string columnId = columnIdentity ?? $"column:{columnIndex + 1}";
        string findingId = BuildFindingId(pagePair.PairIndex, firstRow?.SourceRow ?? 0,
            secondRow?.SourceRow ?? 0, columnId, normalizedLabel);
        var valueMetadata = GetValueMetadata(firstValue.Length > 0 ? firstValue : secondValue);

        return new TableCellDiff
        {
            FindingId = findingId,
            TableId = tableId,
            RowId = BuildStableId("row", normalizedLabel),
            ColumnId = columnId,
            Row = firstRow?.SourceRow ?? secondRow?.SourceRow ?? 0,
            Col = columnIndex + 1,
            SourceRowAfs1 = firstRow?.SourceRow ?? 0,
            SourceRowAfs2 = secondRow?.SourceRow ?? 0,
            Label = row.Label,
            ColHeader = columnIdentity ?? $"Column {columnIndex + 1}",
            ValueAfs1 = firstValue,
            ValueAfs2 = secondValue,
            SourceTextAfs1 = firstRow?.SourceText ?? "",
            SourceTextAfs2 = secondRow?.SourceText ?? "",
            Currency = valueMetadata.Currency,
            Unit = valueMetadata.Unit,
            DifferenceType = status == "changed" ? "financial-cell" : $"financial-row-{status}",
            SemanticClassification = match.Confidence >= 0.82 ? "confirmed-change" : "possible-change",
            Status = status,
            Confidence = Math.Min(match.Confidence, row.ExtractionConfidence),
            HighlightEligible = match.Confidence >= 0.82 && row.ExtractionConfidence >= 0.85,
            PctChange = CalculatePercentageChange(firstValue, secondValue),
            PageAfs1 = pagePair.PageAfs1,
            PageAfs2 = pagePair.PageAfs2,
            PairIndex = pagePair.PairIndex,
        };
    }

    private static List<FinancialCalculationRelationship> ValidateCalculations(
        IReadOnlyList<TablePagePair> tables,
        List<TableCellDiff> findings)
    {
        var relationships = new List<FinancialCalculationRelationship>();
        foreach (var table in tables)
        {
            var orderedRows = table.RowMatches.Where(match => match.First != null && match.Second != null)
                .OrderBy(match => match.First!.SourceRow)
                .ToList();

            for (int targetIndex = 0; targetIndex < orderedRows.Count; targetIndex++)
            {
                var target = orderedRows[targetIndex];
                if (!AggregateLabel.IsMatch(target.First!.Label) || !AggregateLabel.IsMatch(target.Second!.Label))
                    continue;

                int previousAggregate = targetIndex - 1;
                while (previousAggregate >= 0 && !AggregateLabel.IsMatch(orderedRows[previousAggregate].First!.Label))
                    previousAggregate--;
                int firstCandidateIndex = Math.Max(0, previousAggregate);
                var componentCandidates = new List<List<RowMatch>>();
                for (int componentCount = 2; componentCount <= targetIndex - firstCandidateIndex; componentCount++)
                {
                    for (int candidateStart = targetIndex - componentCount; candidateStart >= firstCandidateIndex; candidateStart--)
                        componentCandidates.Add(orderedRows.Skip(candidateStart).Take(componentCount).ToList());
                }

                if (componentCandidates.Count == 0)
                    continue;

                foreach (var targetColumn in MatchColumns(target.First, target.Second))
                {
                    if (!targetColumn.FirstIndex.HasValue || !targetColumn.SecondIndex.HasValue)
                        continue;
                    if (!TryParseAmount(target.First.Values[targetColumn.FirstIndex.Value], out decimal targetOld) ||
                        !TryParseAmount(target.Second.Values[targetColumn.SecondIndex.Value], out decimal targetNew))
                        continue;

                    List<RowMatch>? components = null;
                    decimal sumOld = 0;
                    decimal sumNew = 0;
                    List<TableCellDiff> componentChanges = new();
                    decimal reviewSumOld = 0;
                    decimal reviewSumNew = 0;
                    List<TableCellDiff> reviewComponentChanges = new();
                    bool hasReviewCandidate = false;
                    foreach (var candidateComponents in componentCandidates.Where(candidate => candidate.Count > 0))
                    {
                        decimal candidateOld = 0;
                        decimal candidateNew = 0;
                        var candidateChanges = new List<TableCellDiff>();
                        bool valid = true;
                        foreach (var component in candidateComponents)
                        {
                            var column = MatchColumns(component.First!, component.Second!)
                                .FirstOrDefault(item => string.Equals(item.Identity, targetColumn.Identity, StringComparison.OrdinalIgnoreCase));
                            if (column?.FirstIndex is not int oldIndex || column.SecondIndex is not int newIndex ||
                                !TryParseAmount(component.First!.Values[oldIndex], out decimal oldValue) ||
                                !TryParseAmount(component.Second!.Values[newIndex], out decimal newValue))
                            {
                                valid = false;
                                break;
                            }

                            candidateOld += oldValue;
                            candidateNew += newValue;
                            var componentFinding = findings.FirstOrDefault(finding =>
                                finding.TableId == table.TableId &&
                                finding.SourceRowAfs1 == component.First.SourceRow &&
                                finding.SourceRowAfs2 == component.Second.SourceRow &&
                                string.Equals(finding.ColumnId, column.Identity, StringComparison.OrdinalIgnoreCase));
                            if (componentFinding != null)
                                candidateChanges.Add(componentFinding);
                        }

                        if (valid && !hasReviewCandidate)
                        {
                            hasReviewCandidate = true;
                            reviewSumOld = candidateOld;
                            reviewSumNew = candidateNew;
                            reviewComponentChanges = candidateChanges;
                        }

                        if (valid && candidateOld == targetOld && candidateNew == targetNew)
                        {
                            components = candidateComponents;
                            sumOld = candidateOld;
                            sumNew = candidateNew;
                            componentChanges = candidateChanges;
                            break;
                        }
                    }

                    var targetFinding = findings.FirstOrDefault(finding =>
                        finding.TableId == table.TableId &&
                        finding.SourceRowAfs1 == target.First.SourceRow &&
                        finding.SourceRowAfs2 == target.Second.SourceRow &&
                        string.Equals(finding.ColumnId, targetColumn.Identity, StringComparison.OrdinalIgnoreCase));

                    if (components == null)
                    {
                        if (!hasReviewCandidate || (targetFinding == null && reviewComponentChanges.Count == 0))
                            continue;

                        string reviewRoot = reviewComponentChanges.FirstOrDefault()?.FindingId ?? targetFinding?.FindingId ?? "";
                        var reviewRelationship = new FinancialCalculationRelationship
                        {
                            Id = BuildStableId("calc-review", $"{table.TableId}|{target.First.SourceRow}|{targetColumn.Identity}"),
                            TableId = table.TableId,
                            TargetFindingId = targetFinding?.FindingId ?? "",
                            ComponentFindingIds = reviewComponentChanges.Select(change => change.FindingId).ToList(),
                            Relationship = "unverified-additive",
                            OriginalCalculatedValue = reviewSumOld,
                            RevisedCalculatedValue = reviewSumNew,
                            Confidence = 0.45,
                            RequiresReview = true,
                        };
                        relationships.Add(reviewRelationship);
                        if (targetFinding != null)
                        {
                            targetFinding.CalculationConfidence = reviewRelationship.Confidence;
                            targetFinding.DependencyFindingIds = reviewRelationship.ComponentFindingIds;
                            targetFinding.RootFindingId = reviewRoot;
                        }
                        foreach (var componentFinding in reviewComponentChanges)
                        {
                            componentFinding.RootFindingId = reviewRoot;
                            if (targetFinding != null)
                                componentFinding.DependencyFindingIds = new List<string> { targetFinding.FindingId };
                        }
                        continue;
                    }

                    if (targetFinding == null && componentChanges.Count == 0)
                        continue;

                    string rootFindingId = componentChanges.FirstOrDefault()?.FindingId ?? targetFinding?.FindingId ?? "";
                    var relationship = new FinancialCalculationRelationship
                    {
                        Id = BuildStableId("calc", $"{table.TableId}|{target.First.SourceRow}|{targetColumn.Identity}"),
                        TableId = table.TableId,
                        TargetFindingId = targetFinding?.FindingId ?? "",
                        ComponentFindingIds = componentChanges.Select(change => change.FindingId).ToList(),
                        Relationship = "sum",
                        OriginalCalculatedValue = sumOld,
                        RevisedCalculatedValue = sumNew,
                        Confidence = 0.92,
                    };
                    relationships.Add(relationship);

                    if (targetFinding != null)
                    {
                        targetFinding.CalculationValidated = true;
                        targetFinding.CalculationConfidence = relationship.Confidence;
                        targetFinding.RootFindingId = rootFindingId;
                        targetFinding.DependencyFindingIds = relationship.ComponentFindingIds;
                    }

                    foreach (var componentFinding in componentChanges)
                    {
                        componentFinding.RootFindingId = rootFindingId;
                        if (targetFinding != null)
                            componentFinding.DependencyFindingIds = new List<string> { targetFinding.FindingId };
                    }
                }
            }
        }

        return relationships;
    }

    private static bool IsSingleFinancialCell(string value) =>
        FinancialToken.IsMatch(value) && AuditTextNormalizer.ExtractNormalizedNumbers(value).Count == 1;

    private static bool IsPlausibleRowLabel(string label)
    {
        string normalized = label.Trim();
        return normalized.Length >= 2 && normalized.Length <= 180 && normalized.Any(char.IsLetter);
    }

    private static bool LooksLikeHeading(string line) =>
        line.Length is >= 3 and <= 140 && line.Any(char.IsLetter) &&
        AuditTextNormalizer.ExtractNormalizedNumbers(line).Count == 0;

    private static bool IsLikelyPeriodHeader(string line)
    {
        var years = YearToken.Matches(line);
        if (years.Count == 0)
            return false;

        return years.Count >= 2 ||
               Regex.IsMatch(line, @"\b(?:reporting|financial)\s+(?:year|period)|\b(?:year|period)\s+ended\b",
                   RegexOptions.IgnoreCase);
    }

    private static HashSet<string> LabelTokens(string label) => WordToken.Matches(NormalizeLabel(label))
        .Select(match => match.Value)
        .Where(token => token.Length > 1)
        .ToHashSet(StringComparer.OrdinalIgnoreCase);

    private static string NormalizeLabel(string label) => Regex.Replace(
        Regex.Replace(label.ToLowerInvariant(), @"\b(?:note|ref(?:erence)?)\s+\d+(?:\.\d+)*\b", " "),
        @"[^\p{L}\p{N}]+", " ").Trim();

    private static string BuildTableId(int pairIndex, string heading) =>
        BuildStableId("table", $"{pairIndex}|{NormalizeLabel(heading)}");

    private static string BuildFindingId(int pairIndex, int firstRow, int secondRow, string column, string label) =>
        BuildStableId("cell", $"{pairIndex}|{firstRow}|{secondRow}|{column}|{label}");

    private static string BuildStableId(string prefix, string key)
    {
        string hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(key)))[..16];
        return $"{prefix}-{hash}";
    }

    private static (string Currency, string Unit) GetValueMetadata(string value)
    {
        string currency = CurrencyToken.Match(value ?? "").Value.ToUpperInvariant();
        string unit = (value ?? "").Contains('%') ? "percent" : "amount";
        return (currency, unit);
    }

    private static double? CalculatePercentageChange(string first, string second)
    {
        if (!TryParseAmount(first, out decimal oldValue) || !TryParseAmount(second, out decimal newValue) || oldValue == 0)
            return null;
        return Math.Round((double)Math.Abs(newValue - oldValue) / (double)Math.Abs(oldValue) * 100, 1);
    }

    private static bool TryParseAmount(string value, out decimal amount)
    {
        var numbers = AuditTextNormalizer.ExtractNormalizedNumbers(value);
        if (numbers.Count != 1 || numbers.Single().StartsWith("date:", StringComparison.Ordinal))
        {
            amount = 0;
            return false;
        }

        if (!decimal.TryParse(numbers.Single(), NumberStyles.Number, CultureInfo.InvariantCulture, out amount))
            return false;
        if (value.Contains('(') || Regex.IsMatch(value, @"[-−]\s*\d"))
            amount = -Math.Abs(amount);
        return true;
    }

    private static List<string> ExtractCellVisualTokens(string value) =>
        Regex.Matches(value ?? "", @"\d+(?:[.,]\d+)?|[%$€£¥₹]|\b(?:ZAR|USD|EUR|GBP|AUD|CAD|JPY|CHF)\b",
                RegexOptions.IgnoreCase)
            .Select(match => match.Value)
            .ToList();

    private static string SafePage(List<string> pages, int index) =>
        index >= 0 && index < pages.Count ? pages[index] : "";

    // Exposed internally to the controller so cell findings can use the exact
    // source row and value tokens without asking the renderer to infer a diff.
    internal static List<string> GetVisualTokens(string value) => ExtractCellVisualTokens(value);
}
