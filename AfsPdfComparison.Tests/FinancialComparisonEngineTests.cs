using AfsPdfComparison.Models;
using AfsPdfComparison.Services;
using AfsPdfComparison.Controllers;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using SkiaSharp;
using UglyToad.PdfPig;

namespace AfsPdfComparison.Tests;

public class FinancialComparisonEngineTests
{
    [Fact]
    public void ChangedAmountProducesOneTraceableCellFinding()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nOperating reserve    842,766  840,000\nOther current resources    91,716  90,000",
            "Reporting periods 2025 2024\nOperating reserve    842,765  840,000\nOther current resources    91,716  90,000");

        var finding = Assert.Single(comparison.TableCellDiffs);
        Assert.Equal("changed", finding.Status);
        Assert.Equal("Operating reserve", finding.Label);
        Assert.Equal("2025", finding.ColumnId);
        Assert.True(finding.HighlightEligible);
        Assert.NotEmpty(finding.FindingId);
        Assert.Equal(0.99, finding.Confidence);
    }

    [Fact]
    public void OrderedPageReferenceRowsBecomeStructuralPaginationFindings()
    {
        var comparison = CompareTable(
            "Index\nPage\nGovernance oversight    3\nAudit services    8\nFinancial statements    12",
            "Index\nPage\nGovernance oversight    5\nAudit services    10\nFinancial statements    14");

        Assert.Equal(3, comparison.TableCellDiffs.Count);
        Assert.All(comparison.TableCellDiffs, finding =>
        {
            Assert.Equal("pagination-reference", finding.DifferenceType);
            Assert.Equal("pagination-shift", finding.StructuralChangeType);
            Assert.Equal("structural-change", finding.SemanticClassification);
            Assert.False(finding.HighlightEligible);
        });
        Assert.Equal("pagination-shift", comparison.PageDiffs[0].StructuralChangeType);
        Assert.Equal(2, comparison.PageDiffs[0].StructuralPageOffset);
    }

    [Fact]
    public void MultiPageReferenceRowsUseTheirSharedDynamicOffset()
    {
        var comparison = CompareTable(
            "Contents\nPage\nAudit report    5  7\nPolicy section    12  15\nNotes section    16  20",
            "Contents\nPage\nAudit report    6  8\nPolicy section    13  16\nNotes section    17  21");

        Assert.Equal(6, comparison.TableCellDiffs.Count);
        Assert.All(comparison.TableCellDiffs, finding =>
        {
            Assert.Equal("pagination-reference", finding.DifferenceType);
            Assert.False(finding.HighlightEligible);
        });
        Assert.Equal(1, comparison.PageDiffs[0].StructuralPageOffset);
    }

    [Fact]
    public void ChangedAmountAndMathematicallyValidatedSubtotalAreBothReported()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nOperating reserve    842,766  840,000\nOther current resources    91,716  90,000\nTotal current resources    934,482  930,000",
            "Reporting periods 2025 2024\nOperating reserve    842,765  840,000\nOther current resources    91,716  90,000\nTotal current resources    934,481  930,000");

        Assert.Equal(2, comparison.TableCellDiffs.Count);
        Assert.Contains(comparison.TableCellDiffs, finding => finding.Label == "Operating reserve");
        var subtotal = Assert.Single(comparison.TableCellDiffs.Where(finding => finding.Label == "Total current resources"));
        Assert.True(subtotal.CalculationValidated);
        Assert.NotNull(subtotal.RootFindingId);
        Assert.Single(comparison.CalculationRelationships);
    }

    [Fact]
    public void UnlabeledSubtotalIsReconstructedOnlyWhenItsColumnsValidateArithmetically()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nOpening resource    50,000  25,000\nComponent alpha    63,000  312,000\nComponent beta    607,531  840,725\nComponent gamma    534,044  293,217\n    1,204,575  1,445,942",
            "Reporting periods 2025 2024\nOpening resource    50,000  25,000\nComponent alpha    63,000  312,000\nComponent beta    607,532  840,725\nComponent gamma    534,044  293,217\n    1,204,576  1,445,942");

        Assert.Equal(2, comparison.TableCellDiffs.Count);
        Assert.Contains(comparison.TableCellDiffs, finding => finding.Label == "Component beta");
        var subtotal = Assert.Single(comparison.TableCellDiffs.Where(finding => finding.Label.StartsWith("Subtotal of ")));
        Assert.Equal("1,204,575", subtotal.ValueAfs1);
        Assert.Equal("1,204,576", subtotal.ValueAfs2);
        Assert.True(subtotal.CalculationValidated);
        Assert.NotNull(subtotal.RootFindingId);
    }

    [Fact]
    public void UnlabeledSubtotalMatchesFinancialColumnsWhenRowsAlsoContainReferenceColumns()
    {
        var comparison = CompareTable(
            "Figures in Rand Note 2025 2024\nNon-current assets\nComponent alpha    1  50,000  25,000\nCurrent assets\nComponent beta    2  91,716  13,835\nComponent gamma    3  842,766  1,027,525\n    934,482  1,041,360\nTotal Assets    992,269  1,128,446",
            "Figures in Rand Note 2025 2024\nNon-current assets\nComponent alpha    1  50,000  25,000\nCurrent assets\nComponent beta    2  91,716  13,835\nComponent gamma    3  842,765  1,027,525\n    934,481  1,041,360\nTotal Assets    992,268  1,128,446");

        Assert.Contains(comparison.TableCellDiffs, finding =>
            finding.Label == "Component gamma" && finding.ValueAfs1 == "842,766" && finding.ValueAfs2 == "842,765");
        var subtotal = Assert.Single(comparison.TableCellDiffs.Where(finding => finding.Label.StartsWith("Subtotal of ")));
        Assert.Equal("934,482", subtotal.ValueAfs1);
        Assert.Equal("934,481", subtotal.ValueAfs2);
        Assert.Contains(comparison.TableCellDiffs, finding =>
            finding.Label == "Total Assets" && finding.ValueAfs1 == "992,269" && finding.ValueAfs2 == "992,268");
    }

    [Fact]
    public void ChangedRowMovedAcrossPagesIsMatchedAndKeepsEachDocumentPage()
    {
        string[] originalPages =
        {
            "Schedule\nReporting periods 2025 2024\nCurrent reserve    100  90",
            "Schedule\nReporting periods 2025 2024\nBank balances    842,766  1,027,525",
        };
        string[] revisedPages =
        {
            "Schedule\nReporting periods 2025 2024\nBank balances    842,765  1,027,525",
            "Schedule\nReporting periods 2025 2024\nCurrent reserve    100  90",
        };
        var comparison = new ComparisonResult
        {
            Report1 = new PdfReport { Pages = originalPages.ToList(), FullText = string.Join("\n", originalPages) },
            Report2 = new PdfReport { Pages = revisedPages.ToList(), FullText = string.Join("\n", revisedPages) },
            PageDiffs = new()
            {
                new PageDiffResult { PairIndex = 0, PageAfs1 = 1, PageAfs2 = 1 },
                new PageDiffResult { PairIndex = 1, PageAfs1 = 2, PageAfs2 = 2 },
            },
        };

        comparison.TableCellDiffs = new TableDiffService().Diff(comparison);

        var bankChange = Assert.Single(comparison.TableCellDiffs.Where(finding => finding.Label == "Bank balances"));
        Assert.Equal("changed", bankChange.Status);
        Assert.Equal("842,766", bankChange.ValueAfs1);
        Assert.Equal("842,765", bankChange.ValueAfs2);
        Assert.Equal(2, bankChange.PageAfs1);
        Assert.Equal(1, bankChange.PageAfs2);
        Assert.True(bankChange.HighlightEligible);
        Assert.DoesNotContain(comparison.TableCellDiffs, finding =>
            finding.Label == "Bank balances" && finding.Status is "added" or "removed");

        var relocatedUnchangedRow = Assert.Single(comparison.TableCellDiffs.Where(finding =>
            finding.Label == "Current reserve"));
        Assert.Equal("section-relocation", relocatedUnchangedRow.StructuralChangeType);
        Assert.Equal("structural-change", relocatedUnchangedRow.SemanticClassification);
        Assert.False(relocatedUnchangedRow.HighlightEligible);
        Assert.Equal(1, relocatedUnchangedRow.PageAfs1);
        Assert.Equal(2, relocatedUnchangedRow.PageAfs2);
    }

    [Fact]
    public void ConsequentialTotalChangesRemainIndividualHighlights()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nOperating reserve    842,766  840,000\nOther current resources    91,716  90,000\nTotal current resources    934,482  930,000\nLong term resources    200,000  190,000\nTotal resources    1,134,482  1,120,000",
            "Reporting periods 2025 2024\nOperating reserve    842,765  840,000\nOther current resources    91,716  90,000\nTotal current resources    934,481  930,000\nLong term resources    200,000  190,000\nTotal resources    1,134,481  1,120,000");

        Assert.Equal(3, comparison.TableCellDiffs.Count);
        Assert.All(comparison.TableCellDiffs, finding => Assert.True(finding.HighlightEligible));
        Assert.Equal(2, comparison.CalculationRelationships.Count);
        Assert.All(comparison.CalculationRelationships, relationship => Assert.True(relationship.Confidence >= 0.9));
    }

    [Fact]
    public void InsertedPageProducesCalculatedPaginationShiftWithoutFinancialChanges()
    {
        string first = NarrativePage("position", "The entity records its resources, obligations, and accumulated balances at the end of each reporting period.");
        string second = NarrativePage("performance", "The entity records income, expenses, and operating results for the reporting period under the adopted policies.");
        string inserted = NarrativePage("governance", "The governing body approves the annual report after reviewing the submitted operating and financial records.");

        var result = BuildComparison(new[] { first, second }, new[] { inserted, first, second });

        Assert.Contains(result.PageDiffs, page => page.PageAfs1 == 1 && page.PageAfs2 == 2 && page.StructuralChangeType == "pagination-shift");
        Assert.Contains(result.PageDiffs, page => page.PageAfs1 == 2 && page.PageAfs2 == 3 && page.StructuralChangeType == "pagination-shift");
        Assert.Empty(result.TableCellDiffs);
        Assert.DoesNotContain(result.PageDiffs.SelectMany(page => page.Diffs), diff => diff.HighlightEligible && diff.Status == "changed");
    }

    [Fact]
    public void ReorderedSectionsAreMatchedAsRelocations()
    {
        string first = NarrativePage("resources", "The organization records its resources, short term balances, and obligations in the annual financial report.");
        string second = NarrativePage("performance", "The organization presents its annual operating results, income, and expenses in the performance report.");
        string third = NarrativePage("governance", "The governing body approves the annual statements after a review of supporting records and management explanations.");

        var alignment = new PageAlignmentService().BuildAlignment(
            new() { first, second, third },
            new() { second, first, third });

        Assert.Contains(alignment, entry => entry.MatchType == "relocated" && entry.I2 >= 0);
    }

    [Fact]
    public void NumberFormattingOnlyDoesNotCreateCellFinding()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nOperating reserve    1 500  1,250",
            "Reporting periods 2025 2024\nOperating reserve    1,500.00  1 250.00");

        Assert.Empty(comparison.TableCellDiffs);
    }

    [Fact]
    public void NegativeToPositiveAmountIsAChangedCell()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nNet adjustment    (150,000)  12,000",
            "Reporting periods 2025 2024\nNet adjustment    150,000  12,000");

        var finding = Assert.Single(comparison.TableCellDiffs);
        Assert.Equal("changed", finding.Status);
        Assert.Equal("2025", finding.ColumnId);
    }

    [Fact]
    public void ReorderedRowsAndPeriodColumnsMatchByMeaning()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nOperating reserve    842,766  840,000\nOther resources    91,716  90,000",
            "Reporting periods 2024 2025\nOther resources    90,000  91,716\nOperating reserve    840,000  842,766");

        Assert.Empty(comparison.TableCellDiffs);
    }

    [Fact]
    public void RemovedDisclosureIsReportedOnOriginalSide()
    {
        const string original = "The entity also reports a separate disclosure describing the basis used to estimate long term obligations.";
        var diffs = new LineComparatorService().CompareLines(
            original + "\nThe remaining policy information is unchanged and available in the supporting schedule.",
            "The remaining policy information is unchanged and available in the supporting schedule.");

        var removed = Assert.Single(diffs.Where(diff => diff.Status == "removed"));
        Assert.Equal(original, removed.Line1);
    }

    [Fact]
    public void NoteRenumberingWithMatchedHeadingIsClassifiedAsStructural()
    {
        var noteReference = new LineDiffResult
        {
            Status = "changed",
            Line1 = "The related terms are explained in Note 5.",
            Line2 = "The related terms are explained in Note 8.",
            Similarity = 0.9,
        };
        var comparison = new ComparisonResult
        {
            Report1 = new PdfReport { Pages = new() { "Note 5 Revenue recognition" } },
            Report2 = new PdfReport { Pages = new() { "Note 8 Revenue recognition" } },
            PageDiffs = new() { new PageDiffResult { PairIndex = 0, PageAfs1 = 1, PageAfs2 = 1, Diffs = new() { noteReference } } },
        };

        FindingClassificationService.Classify(comparison);

        Assert.Equal("reference-renumbering", noteReference.DifferenceType);
        Assert.Equal("structural-change", noteReference.SemanticClassification);
        Assert.False(noteReference.HighlightEligible);
    }

    [Fact]
    public void ExtractionUncertaintyRequiresReviewInsteadOfConfirmedHighlight()
    {
        var uncertain = new LineDiffResult
        {
            Status = "changed",
            Line1 = "Operating reserve 840,000",
            Line2 = "Operating reserve 84?000",
            Similarity = 0.9,
        };
        var comparison = new ComparisonResult
        {
            Report1 = new PdfReport { Pages = new() { "Operating reserve 840,000" } },
            Report2 = new PdfReport { Pages = new() { "Operating reserve 84?000" } },
            PageDiffs = new() { new PageDiffResult { PairIndex = 0, PageAfs1 = 1, PageAfs2 = 1, Diffs = new() { uncertain } } },
        };

        FindingClassificationService.Classify(comparison);

        Assert.Equal("extraction-error", uncertain.SemanticClassification);
        Assert.False(uncertain.HighlightEligible);
        Assert.True(uncertain.Confidence < 0.62);
    }

    [Fact]
    public void FinancialCellHighlightContractTargetsValueAndRequiresApproval()
    {
        var finding = new TableCellDiff
        {
            FindingId = "cell-finding",
            Label = "Operating reserve",
            ValueAfs1 = "842,766",
            ValueAfs2 = "842,765",
            SourceTextAfs1 = "Operating reserve    842,766",
            SourceTextAfs2 = "Operating reserve    842,765",
            Status = "changed",
            HighlightEligible = true,
            Confidence = 0.99,
        };
        var spec = InvokePrivateStatic<HighlightSpec>(
            typeof(AfsController), "BuildFinancialCellHighlightSpec", finding, 2);

        Assert.Equal(finding.FindingId, spec.FindingId);
        Assert.True(spec.HighlightEligible);
        Assert.True(spec.ExactRegion);
        Assert.Empty(spec.ChangedWords);
        Assert.Contains("842,765", spec.ChangedPhrases);
    }

    [Fact]
    public void PaginationReferenceRemainsStructuralButCanReceiveAReviewMarker()
    {
        var finding = new TableCellDiff
        {
            FindingId = "structural-reference",
            Label = "Governance oversight",
            ValueAfs1 = "8",
            ValueAfs2 = "9",
            SourceTextAfs1 = "Governance oversight    8",
            SourceTextAfs2 = "Governance oversight    9",
            Status = "changed",
            DifferenceType = "pagination-reference",
            StructuralChangeType = "pagination-shift",
            SemanticClassification = "structural-change",
            HighlightEligible = false,
            Confidence = 0.95,
        };

        var spec = InvokePrivateStatic<HighlightSpec>(
            typeof(AfsController), "BuildFinancialCellHighlightSpec", finding, 2);

        Assert.False(spec.HighlightEligible);
        Assert.True(spec.StructuralHighlightEligible);
        Assert.Equal("structural-change", finding.SemanticClassification);
        Assert.Equal("pagination-shift", spec.StructuralChangeType);
    }

    [Fact]
    public async Task NumericCellHighlightDoesNotCaptureMatchingNoteReference()
    {
        QuestPDF.Settings.License = LicenseType.Community;
        byte[] originalPdf = CreateFinancialPdfWithNoteColumn("11 296");
        byte[] revisedPdf = CreateFinancialPdfWithNoteColumn("11 295");
        string originalPath = Path.Combine(Path.GetTempPath(), $"afs-note-original-{Guid.NewGuid():N}.pdf");
        string revisedPath = Path.Combine(Path.GetTempPath(), $"afs-note-revised-{Guid.NewGuid():N}.pdf");

        try
        {
            await File.WriteAllBytesAsync(originalPath, originalPdf);
            await File.WriteAllBytesAsync(revisedPath, revisedPdf);
            var comparison = CompareTable(ExtractPageText(originalPdf), ExtractPageText(revisedPdf));
            var finding = Assert.Single(comparison.TableCellDiffs.Where(cell => cell.Status == "changed"));
            var pagePair = new PageDiffResult { PairIndex = 0, PageAfs1 = 1, PageAfs2 = 1 };
            var revisedSpec = InvokePrivateStatic<HighlightSpec>(
                typeof(AfsController), "BuildFinancialCellHighlightSpec", finding, 2);
            revisedSpec.OriginalValue = finding.ValueAfs1;
            revisedSpec.RevisedValue = finding.ValueAfs2;
            revisedSpec.TableId = finding.TableId;
            revisedSpec.ColumnId = finding.ColumnId;
            revisedSpec.SourceCoordinates = ResolveCoordinates(revisedPath, pagePair, revisedSpec, 2);

            Assert.NotEmpty(revisedSpec.SourceCoordinates);
            using var document = PdfDocument.Open(revisedPdf);
            var page = document.GetPage(1);
            var rowWords = page.GetWords()
                .Where(word => revisedSpec.SourceCoordinates.Any(rect =>
                    Math.Abs(rect.Y - (page.Height - word.BoundingBox.Top)) < 5 &&
                    Math.Abs(rect.Y + rect.Height - (page.Height - word.BoundingBox.Bottom)) < 5))
                .OrderBy(word => word.BoundingBox.Left)
                .ToList();
            var noteReference = rowWords.FirstOrDefault(word => word.Text.Trim() == "11");
            Assert.NotNull(noteReference);
            Assert.DoesNotContain(revisedSpec.SourceCoordinates, rect =>
                rect.X <= noteReference!.BoundingBox.Right && rect.X + rect.Width >= noteReference.BoundingBox.Left);
        }
        finally
        {
            File.Delete(originalPath);
            File.Delete(revisedPath);
        }
    }

    [Fact]
    public void MultipleChangedRowsAndTheirReportedTotalsRemainCellFindings()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nPrimary inflows 6    5 372 097  5 085 031\nSupplemental inflows 7    1 204 575  1 445 942\nOperating outflows 8    (6 705 724)  (6 635 514)\nOperating result    (129 052)  (104 541)\nFinance income 11    11 296  26 584\nPeriod result    (117 756)  (77 957)\nComprehensive result    (117 756)  (77 957)",
            "Reporting periods 2025 2024\nPrimary inflows 6    5 372 098  5 085 031\nSupplemental inflows 7    1 204 576  1 445 942\nOperating outflows 8    (6 705 726)  (6 635 514)\nOperating result    (129 052)  (104 541)\nFinance income 11    11 295  26 584\nPeriod result    (117 757)  (77 957)\nComprehensive result    (117 757)  (77 957)");

        var changedLabels = comparison.TableCellDiffs
            .Where(finding => finding.Status == "changed")
            .Select(finding => finding.Label)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains("Primary inflows 6", changedLabels);
        Assert.Contains("Supplemental inflows 7", changedLabels);
        Assert.Contains("Operating outflows 8", changedLabels);
        Assert.Contains("Finance income 11", changedLabels);
        Assert.Contains("Period result", changedLabels);
        Assert.Contains("Comprehensive result", changedLabels);
        Assert.Equal(6, changedLabels.Count);
    }

    [Fact]
    public async Task SyntheticPdfsRenderVerifiedCellHighlightsOnBothSides()
    {
        QuestPDF.Settings.License = LicenseType.Community;
        byte[] originalPdf = CreateFinancialPdf("842,766");
        byte[] revisedPdf = CreateFinancialPdf("842,765");
        string originalPath = Path.Combine(Path.GetTempPath(), $"afs-original-{Guid.NewGuid():N}.pdf");
        string revisedPath = Path.Combine(Path.GetTempPath(), $"afs-revised-{Guid.NewGuid():N}.pdf");

        try
        {
            await File.WriteAllBytesAsync(originalPath, originalPdf);
            await File.WriteAllBytesAsync(revisedPath, revisedPdf);
            string originalPage = ExtractPageText(originalPdf);
            string revisedPage = ExtractPageText(revisedPdf);
            var comparison = CompareTable(originalPage, revisedPage);
            var finding = Assert.Single(comparison.TableCellDiffs);
            var pagePair = new PageDiffResult { PairIndex = 0, PageAfs1 = 1, PageAfs2 = 1 };
            var originalSpec = CreateExactCellSpec(finding, originalSide: true);
            var revisedSpec = CreateExactCellSpec(finding, originalSide: false);
            originalSpec.SourceCoordinates = ResolveCoordinates(originalPath, pagePair, originalSpec, 1);
            revisedSpec.SourceCoordinates = ResolveCoordinates(revisedPath, pagePair, revisedSpec, 2);
            originalSpec.DocumentSide = 1;
            revisedSpec.DocumentSide = 2;
            AssertCoordinatesResolveToAmount(originalPdf, originalSpec.SourceCoordinates, "842766");
            AssertCoordinatesResolveToAmount(revisedPdf, revisedSpec.SourceCoordinates, "842765");
            Assert.NotEmpty(originalSpec.SourceCoordinates);
            Assert.NotEmpty(revisedSpec.SourceCoordinates);

            var renderer = new PageSnapshotService();
            string originalPlain = await renderer.RenderPageToBase64(originalPdf, 0, new List<string>());
            string revisedPlain = await renderer.RenderPageToBase64(revisedPdf, 0, new List<string>());
            string originalHighlighted = await renderer.RenderPageToBase64(originalPdf, 0, new List<HighlightSpec> { originalSpec });
            string revisedHighlighted = await renderer.RenderPageToBase64(revisedPdf, 0, new List<HighlightSpec> { revisedSpec });
            var structuralSpec = new HighlightSpec
            {
                FindingId = finding.FindingId,
                LineText = revisedSpec.LineText,
                ChangedWords = revisedSpec.ChangedWords,
                AnchorWords = revisedSpec.AnchorWords,
                ChangedPhrases = revisedSpec.ChangedPhrases,
                Confidence = revisedSpec.Confidence,
                HighlightEligible = false,
                StructuralHighlightEligible = true,
                StructuralChangeType = "pagination-shift",
                DifferenceType = "pagination-reference",
                ExactRegion = true,
                SourceCoordinates = revisedSpec.SourceCoordinates,
            };
            string revisedStructural = await renderer.RenderPageToBase64(revisedPdf, 0, new List<HighlightSpec> { structuralSpec });
            using var originalBase = SKBitmap.Decode(Convert.FromBase64String(originalPlain));
            using var revisedBase = SKBitmap.Decode(Convert.FromBase64String(revisedPlain));
            using var originalMarked = SKBitmap.Decode(Convert.FromBase64String(originalHighlighted));
            using var revisedMarked = SKBitmap.Decode(Convert.FromBase64String(revisedHighlighted));
            using var revisedStructuralMarked = SKBitmap.Decode(Convert.FromBase64String(revisedStructural));
            Assert.NotNull(originalBase);
            Assert.NotNull(revisedBase);
            Assert.NotNull(originalMarked);
            Assert.NotNull(revisedMarked);
            Assert.NotNull(revisedStructuralMarked);
            using var originalPdfDocument = PdfDocument.Open(originalPdf);
            using var revisedPdfDocument = PdfDocument.Open(revisedPdf);
            var originalPdfPage = originalPdfDocument.GetPage(1);
            var revisedPdfPage = revisedPdfDocument.GetPage(1);
            Assert.True(DarkPixelsInCoordinates(originalBase!, originalSpec.SourceCoordinates,
                originalPdfPage.Width, originalPdfPage.Height) > 0);
            Assert.True(DarkPixelsInCoordinates(revisedBase!, revisedSpec.SourceCoordinates,
                revisedPdfPage.Width, revisedPdfPage.Height) > 0);
            Assert.True(ChangedPixelsInCoordinates(originalBase!, originalMarked!, originalSpec.SourceCoordinates) > 20);
            Assert.True(ChangedPixelsInCoordinates(revisedBase!, revisedMarked!, revisedSpec.SourceCoordinates) > 20);
            Assert.True(ChangedPixelsInCoordinates(revisedBase!, revisedStructuralMarked!, structuralSpec.SourceCoordinates) > 20);
            Assert.True(YellowTintPixelsInCoordinates(revisedStructuralMarked!, structuralSpec.SourceCoordinates) > 0);

            using var sideBySide = new SKBitmap(originalMarked!.Width + revisedMarked!.Width,
                Math.Max(originalMarked.Height, revisedMarked.Height));
            using (var canvas = new SKCanvas(sideBySide))
            {
                canvas.Clear(SKColors.White);
                canvas.DrawBitmap(originalMarked, 0, 0);
                canvas.DrawBitmap(revisedMarked, originalMarked.Width, 0);
            }

            string visualArtifact = Path.Combine(Path.GetTempPath(), "afs-comparison-side-by-side.png");
            using var image = SKImage.FromBitmap(sideBySide);
            using var encoded = image.Encode(SKEncodedImageFormat.Png, 90);
            await File.WriteAllBytesAsync(visualArtifact, encoded.ToArray());
            Assert.True(new FileInfo(visualArtifact).Length > 10_000);
        }
        finally
        {
            File.Delete(originalPath);
            File.Delete(revisedPath);
        }
    }

    [Fact]
    public async Task RelocatedFinancialCellHighlightsItsIndependentPageOnEachPdf()
    {
        QuestPDF.Settings.License = LicenseType.Community;
        byte[] originalPdf = CreateTwoPageFinancialPdf("842,766", amountPage: 2);
        byte[] revisedPdf = CreateTwoPageFinancialPdf("842,765", amountPage: 1);
        string originalPath = Path.Combine(Path.GetTempPath(), $"afs-relocated-original-{Guid.NewGuid():N}.pdf");
        string revisedPath = Path.Combine(Path.GetTempPath(), $"afs-relocated-revised-{Guid.NewGuid():N}.pdf");

        try
        {
            await File.WriteAllBytesAsync(originalPath, originalPdf);
            await File.WriteAllBytesAsync(revisedPath, revisedPdf);
            var finding = new TableCellDiff
            {
                FindingId = "relocated-financial-cell",
                TableId = "matched-schedule",
                RowId = "row-bank-balances",
                ColumnId = "2025",
                Label = "Bank balances",
                ValueAfs1 = "842,766",
                ValueAfs2 = "842,765",
                SourceTextAfs1 = "Bank balances 842,766 1,027,525",
                SourceTextAfs2 = "Bank balances 842,765 1,027,525",
                Status = "changed",
                HighlightEligible = true,
                Confidence = 0.96,
                PageAfs1 = 2,
                PageAfs2 = 1,
            };
            var containingPair = new PageDiffResult { PairIndex = 0, PageAfs1 = 2, PageAfs2 = 2 };
            var originalSpec = CreateExactCellSpec(finding, originalSide: true);
            var revisedSpec = CreateExactCellSpec(finding, originalSide: false);
            originalSpec.OriginalPageNumber = 2;
            revisedSpec.RevisedPageNumber = 1;
            originalSpec.SourceCoordinates = ResolveCoordinates(originalPath, containingPair, originalSpec, 1);
            revisedSpec.SourceCoordinates = ResolveCoordinates(revisedPath, containingPair, revisedSpec, 2);

            Assert.All(originalSpec.SourceCoordinates, coordinate => Assert.Equal(2, coordinate.PageNumber));
            Assert.All(revisedSpec.SourceCoordinates, coordinate => Assert.Equal(1, coordinate.PageNumber));
            AssertCoordinatesResolveToAmount(originalPdf, originalSpec.SourceCoordinates, "842766");
            AssertCoordinatesResolveToAmount(revisedPdf, revisedSpec.SourceCoordinates, "842765");

            var renderer = new PageSnapshotService();
            string originalBase64 = await renderer.RenderPageToBase64(originalPdf, 1, new List<string>());
            string revisedBase64 = await renderer.RenderPageToBase64(revisedPdf, 0, new List<string>());
            string originalMarkedBase64 = await renderer.RenderPageToBase64(originalPdf, 1, new List<HighlightSpec> { originalSpec });
            string revisedMarkedBase64 = await renderer.RenderPageToBase64(revisedPdf, 0, new List<HighlightSpec> { revisedSpec });
            using var originalBase = SKBitmap.Decode(Convert.FromBase64String(originalBase64));
            using var revisedBase = SKBitmap.Decode(Convert.FromBase64String(revisedBase64));
            using var originalMarked = SKBitmap.Decode(Convert.FromBase64String(originalMarkedBase64));
            using var revisedMarked = SKBitmap.Decode(Convert.FromBase64String(revisedMarkedBase64));

            Assert.NotNull(originalBase);
            Assert.NotNull(revisedBase);
            Assert.NotNull(originalMarked);
            Assert.NotNull(revisedMarked);
            Assert.True(ChangedPixelsInCoordinates(originalBase!, originalMarked!, originalSpec.SourceCoordinates) > 20);
            Assert.True(ChangedPixelsInCoordinates(revisedBase!, revisedMarked!, revisedSpec.SourceCoordinates) > 20);
        }
        finally
        {
            File.Delete(originalPath);
            File.Delete(revisedPath);
        }
    }

    [Fact]
    public void TextHighlightsExcludeUnchangedNumbersFromTheSameLine()
    {
        QuestPDF.Settings.License = LicenseType.Community;
        const string originalLine = "Note 2025 2024";
        const string revisedLine = "Index 2025 2024";
        byte[] originalPdf = CreateTextPdf(originalLine);
        byte[] revisedPdf = CreateTextPdf(revisedLine);
        string originalPath = Path.Combine(Path.GetTempPath(), $"afs-text-original-{Guid.NewGuid():N}.pdf");
        string revisedPath = Path.Combine(Path.GetTempPath(), $"afs-text-revised-{Guid.NewGuid():N}.pdf");

        try
        {
            File.WriteAllBytes(originalPath, originalPdf);
            File.WriteAllBytes(revisedPath, revisedPdf);
            var diff = new LineDiffResult
            {
                Status = "changed",
                Line1 = originalLine,
                Line2 = revisedLine,
                WordDiff = new()
                {
                    new WordDiffToken { Word = "Note", Tag = "removed" },
                    new WordDiffToken { Word = "Index", Tag = "added" },
                    new WordDiffToken { Word = "2025", Tag = "same" },
                    new WordDiffToken { Word = "2024", Tag = "same" },
                },
            };
            var pagePair = new PageDiffResult { PairIndex = 0, PageAfs1 = 1, PageAfs2 = 1 };
            var originalSpec = InvokePrivateStatic<HighlightSpec>(typeof(AfsController), "BuildHighlightSpec", diff, 1);
            var revisedSpec = InvokePrivateStatic<HighlightSpec>(typeof(AfsController), "BuildHighlightSpec", diff, 2);
            originalSpec.OriginalValue = originalLine;
            originalSpec.RevisedValue = revisedLine;
            revisedSpec.OriginalValue = originalLine;
            revisedSpec.RevisedValue = revisedLine;

            var originalCoordinates = ResolveCoordinates(originalPath, pagePair, originalSpec, 1);
            var revisedCoordinates = ResolveCoordinates(revisedPath, pagePair, revisedSpec, 2);
            Assert.NotEmpty(originalCoordinates);
            Assert.NotEmpty(revisedCoordinates);
            Assert.Equal(new[] { "Note" }, WordsAtCoordinates(originalPdf, originalCoordinates));
            Assert.Equal(new[] { "Index" }, WordsAtCoordinates(revisedPdf, revisedCoordinates));
        }
        finally
        {
            File.Delete(originalPath);
            File.Delete(revisedPath);
        }
    }

    [Fact]
    public void ArithmeticThatCannotBeValidatedIsFlaggedForReview()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nPrimary resources    40  45\nOther resources    20  15\nTotal resources    70  60",
            "Reporting periods 2025 2024\nPrimary resources    42  45\nOther resources    20  15\nTotal resources    70  60");

        var relation = Assert.Single(comparison.CalculationRelationships);
        Assert.True(relation.RequiresReview);
        Assert.Equal("unverified-additive", relation.Relationship);
        Assert.Contains(comparison.TableCellDiffs, finding => finding.Label == "Primary resources");
    }

    [Fact]
    public void AddedFinancialRowIsReportedWithItsSourceEvidence()
    {
        var comparison = CompareTable(
            "Reporting periods 2025 2024\nPrimary resources    40  45",
            "Reporting periods 2025 2024\nPrimary resources    40  45\nNewly identified resources    12  10");

        var findings = comparison.TableCellDiffs.Where(cell => cell.Label == "Newly identified resources").ToList();
        Assert.Equal(2, findings.Count);
        Assert.All(findings, finding =>
        {
            Assert.Equal("added", finding.Status);
            Assert.Equal("financial-row-added", finding.DifferenceType);
            Assert.True(finding.HighlightEligible);
            Assert.Contains("Newly identified resources", finding.SourceTextAfs2);
        });
    }

    private static ComparisonResult CompareTable(string firstPage, string secondPage)
    {
        var comparison = new ComparisonResult
        {
            Report1 = new PdfReport { Pages = new() { firstPage }, FullText = firstPage },
            Report2 = new PdfReport { Pages = new() { secondPage }, FullText = secondPage },
            PageDiffs = new() { new PageDiffResult { PairIndex = 0, PageAfs1 = 1, PageAfs2 = 1, AlignSim = 1.0 } },
        };
        comparison.TableCellDiffs = new TableDiffService().Diff(comparison);
        return comparison;
    }

    private static ComparisonResult BuildComparison(string[] firstPages, string[] secondPages)
    {
        var lineComparator = new LineComparatorService();
        var service = new ComparisonService(
            new PageAlignmentService(),
            lineComparator,
            new SmartComparisonOrchestrator(),
            new DocumentIntelligenceService(lineComparator),
            new AutoReviewEngine(new LabeledNumberExtractor()),
            new TableDiffService());
        return service.BuildComparisonAsync(
            new PdfReport { Label = "Original", Pages = firstPages.ToList(), FullText = string.Join("\n", firstPages), PageCount = firstPages.Length },
            new PdfReport { Label = "Revised", Pages = secondPages.ToList(), FullText = string.Join("\n", secondPages), PageCount = secondPages.Length })
            .GetAwaiter().GetResult();
    }

    private static string NarrativePage(string section, string body) =>
        $"{section}\n{body} Additional explanatory context records the reporting policies, evidence, and conclusions used by the organization. The details remain available for independent review by the audit team.";

    private static byte[] CreateFinancialPdf(string amount) => Document.Create(document =>
        document.Page(page =>
        {
            page.Size(PageSizes.A4);
            page.Margin(40);
            page.DefaultTextStyle(style => style.FontSize(12));
            page.Content().Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    columns.RelativeColumn(3);
                    columns.RelativeColumn(1);
                    columns.RelativeColumn(1);
                });

                table.Header(header =>
                {
                    header.Cell().Padding(4).Text("Description");
                    header.Cell().Padding(4).AlignRight().Text("2025");
                    header.Cell().Padding(4).AlignRight().Text("2024");
                });
                table.Cell().Padding(4).Text("Operating reserve");
                table.Cell().Padding(4).AlignRight().Text(amount);
                table.Cell().Padding(4).AlignRight().Text("840,000");
                table.Cell().Padding(4).Text("Other current resources");
                table.Cell().Padding(4).AlignRight().Text("91,716");
                table.Cell().Padding(4).AlignRight().Text("90,000");
            });
        })).GeneratePdf();

    private static byte[] CreateTwoPageFinancialPdf(string amount, int amountPage) => Document.Create(document =>
    {
        for (int pageNumber = 1; pageNumber <= 2; pageNumber++)
        {
            document.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(40);
                page.DefaultTextStyle(style => style.FontSize(12));
                page.Content().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        columns.RelativeColumn(3);
                        columns.RelativeColumn(1);
                        columns.RelativeColumn(1);
                    });

                    table.Header(header =>
                    {
                        header.Cell().Padding(4).Text("Description");
                        header.Cell().Padding(4).AlignRight().Text("2025");
                        header.Cell().Padding(4).AlignRight().Text("2024");
                    });
                    if (pageNumber == amountPage)
                    {
                        table.Cell().Padding(4).Text("Bank balances");
                        table.Cell().Padding(4).AlignRight().Text(amount);
                        table.Cell().Padding(4).AlignRight().Text("1,027,525");
                    }
                    else
                    {
                        table.Cell().Padding(4).Text("Unchanged information");
                        table.Cell().Padding(4).AlignRight().Text("500");
                        table.Cell().Padding(4).AlignRight().Text("400");
                    }
                });
            });
        }
    }).GeneratePdf();

    private static byte[] CreateFinancialPdfWithNoteColumn(string amount) => Document.Create(document =>
        document.Page(page =>
        {
            page.Size(PageSizes.A4);
            page.Margin(40);
            page.DefaultTextStyle(style => style.FontSize(12));
            page.Content().Table(table =>
            {
                table.ColumnsDefinition(columns =>
                {
                    columns.RelativeColumn(4);
                    columns.RelativeColumn(1);
                    columns.RelativeColumn(2);
                    columns.RelativeColumn(2);
                });

                table.Header(header =>
                {
                    header.Cell().Padding(4).Text("Description");
                    header.Cell().Padding(4).AlignRight().Text("Note");
                    header.Cell().Padding(4).AlignRight().Text("2025");
                    header.Cell().Padding(4).AlignRight().Text("2024");
                });
                table.Cell().Padding(4).Text("Finance income");
                table.Cell().Padding(4).AlignRight().Text("11");
                table.Cell().Padding(4).AlignRight().Text(amount);
                table.Cell().Padding(4).AlignRight().Text("26 584");
            });
        })).GeneratePdf();

    private static byte[] CreateTextPdf(string text) => Document.Create(document =>
        document.Page(page =>
        {
            page.Size(PageSizes.A4);
            page.Margin(40);
            page.DefaultTextStyle(style => style.FontSize(12));
            page.Content().Text(text);
        })).GeneratePdf();

    private static string ExtractPageText(byte[] pdfBytes)
    {
        using var document = PdfDocument.Open(pdfBytes);
        var page = document.GetPage(1);
        var method = typeof(PageExtractorService).GetMethod(
            "ExtractPageText", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        Assert.NotNull(method);
        return (string)method!.Invoke(null, new object[] { page })!;
    }

    private static HighlightSpec CreateExactCellSpec(TableCellDiff finding, bool originalSide)
    {
        string value = originalSide ? finding.ValueAfs1 : finding.ValueAfs2;
        return new HighlightSpec
        {
            FindingId = finding.FindingId,
            LineText = originalSide ? finding.SourceTextAfs1 : finding.SourceTextAfs2,
            ChangedWords = System.Text.RegularExpressions.Regex.Matches(value, @"\d+(?:[.,]\d+)?")
                .Select(match => match.Value)
                .ToList(),
            AnchorWords = finding.Label.Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList(),
            ChangedPhrases = new() { value },
            OriginalValue = finding.ValueAfs1,
            RevisedValue = finding.ValueAfs2,
            Confidence = finding.Confidence,
            HighlightEligible = finding.HighlightEligible,
            ExactRegion = true,
            DifferenceType = finding.DifferenceType,
        };
    }

    private static List<PdfSourceCoordinates> ResolveCoordinates(
        string pdfPath,
        PageDiffResult pagePair,
        HighlightSpec spec,
        int side)
    {
        var method = typeof(AfsController).GetMethod(
            "ResolveHighlightCoordinates", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        Assert.NotNull(method);
        return (List<PdfSourceCoordinates>)method!.Invoke(null, new object[] { pdfPath, pagePair, spec, side })!;
    }

    private static int ChangedPixelsInCoordinates(
        SKBitmap baseline,
        SKBitmap highlighted,
        IReadOnlyList<PdfSourceCoordinates> coordinates)
    {
        int changed = 0;
        foreach (var rect in coordinates)
        {
            int left = Math.Clamp((int)(rect.X * baseline.Width / 595.0), 0, baseline.Width - 1);
            int top = Math.Clamp((int)(rect.Y * baseline.Height / 842.0), 0, baseline.Height - 1);
            int right = Math.Clamp((int)((rect.X + rect.Width) * baseline.Width / 595.0) + 1, left + 1, baseline.Width);
            int bottom = Math.Clamp((int)((rect.Y + rect.Height) * baseline.Height / 842.0) + 1, top + 1, baseline.Height);
            for (int y = top; y < bottom; y++)
            for (int x = left; x < right; x++)
            {
                var before = baseline.GetPixel(x, y);
                var after = highlighted.GetPixel(x, y);
                if (Math.Abs(before.Red - after.Red) + Math.Abs(before.Green - after.Green) + Math.Abs(before.Blue - after.Blue) > 30)
                    changed++;
            }
        }
        return changed;
    }

    private static int YellowTintPixelsInCoordinates(
        SKBitmap image,
        IReadOnlyList<PdfSourceCoordinates> coordinates)
    {
        int yellowPixels = 0;
        foreach (var rect in coordinates)
        {
            int left = Math.Clamp((int)(rect.X * image.Width / 595.0), 0, image.Width - 1);
            int top = Math.Clamp((int)(rect.Y * image.Height / 842.0), 0, image.Height - 1);
            int right = Math.Clamp((int)((rect.X + rect.Width) * image.Width / 595.0) + 1, left + 1, image.Width);
            int bottom = Math.Clamp((int)((rect.Y + rect.Height) * image.Height / 842.0) + 1, top + 1, image.Height);
            for (int y = top; y < bottom; y++)
            for (int x = left; x < right; x++)
            {
                SKColor color = image.GetPixel(x, y);
                if (color.Red > color.Blue + 20 && color.Green > color.Blue + 8)
                    yellowPixels++;
            }
        }

        return yellowPixels;
    }

    private static int DarkPixelsInCoordinates(
        SKBitmap bitmap,
        IReadOnlyList<PdfSourceCoordinates> coordinates,
        double pageWidth,
        double pageHeight)
    {
        int darkPixels = 0;
        foreach (var rect in coordinates)
        {
            int left = Math.Clamp((int)(rect.X * bitmap.Width / pageWidth), 0, bitmap.Width - 1);
            int top = Math.Clamp((int)(rect.Y * bitmap.Height / pageHeight), 0, bitmap.Height - 1);
            int right = Math.Clamp((int)((rect.X + rect.Width) * bitmap.Width / pageWidth) + 1, left + 1, bitmap.Width);
            int bottom = Math.Clamp((int)((rect.Y + rect.Height) * bitmap.Height / pageHeight) + 1, top + 1, bitmap.Height);
            for (int y = top; y < bottom; y++)
            for (int x = left; x < right; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.Red < 130 && pixel.Green < 130 && pixel.Blue < 130)
                    darkPixels++;
            }
        }

        return darkPixels;
    }

    private static void AssertCoordinatesResolveToAmount(
        byte[] pdfBytes,
        IReadOnlyList<PdfSourceCoordinates> coordinates,
        string expectedDigits)
    {
        string digits = System.Text.RegularExpressions.Regex.Replace(string.Join("", WordsAtCoordinates(pdfBytes, coordinates)), @"\D", "");
        Assert.Equal(expectedDigits, digits);
    }

    private static string[] WordsAtCoordinates(
        byte[] pdfBytes,
        IReadOnlyList<PdfSourceCoordinates> coordinates)
    {
        using var document = PdfDocument.Open(pdfBytes);
        var words = new List<string>();
        foreach (int pageNumber in coordinates.Select(rect => rect.PageNumber).Distinct())
        {
            if (pageNumber < 1 || pageNumber > document.NumberOfPages)
                continue;

            var page = document.GetPage(pageNumber);
            var pageCoordinates = coordinates.Where(rect => rect.PageNumber == pageNumber).ToList();
            words.AddRange(page.GetWords()
                .Where(word => pageCoordinates.Any(rect =>
                    rect.X <= word.BoundingBox.Right && rect.X + rect.Width >= word.BoundingBox.Left &&
                    rect.Y <= page.Height - word.BoundingBox.Bottom &&
                    rect.Y + rect.Height >= page.Height - word.BoundingBox.Top))
                .Select(word => word.Text ?? ""));
        }
        return words.ToArray();
    }

    private static T InvokePrivateStatic<T>(Type type, string methodName, params object?[] args)
    {
        var method = type.GetMethods(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)
            .First(candidate => candidate.Name == methodName && candidate.GetParameters().Length == args.Length);
        return (T)method.Invoke(null, args)!;
    }
}
