using System.Reflection;
using AfsPdfComparison.Controllers;
using AfsPdfComparison.Models;
using AfsPdfComparison.Services;

namespace AfsPdfComparison.Tests;

public class RegressionFixTests
{
    [Fact]
    public async Task ComparisonService_DefaultFastPath_SkipsDeferredInsights()
    {
        var lineComparator = new LineComparatorService();
        var service = new ComparisonService(
            new PageAlignmentService(),
            lineComparator,
            new SmartComparisonOrchestrator(),
            new DocumentIntelligenceService(lineComparator),
            new AutoReviewEngine(new LabeledNumberExtractor()),
            new TableDiffService());

        var report1 = new PdfReport
        {
            Label = "AFS 1",
            Filename = "afs1.pdf",
            PdfPath = "missing-afs1.pdf",
            Pages = new() { "Revenue 100\nExpenses 70" },
            FullText = "Revenue 100\nExpenses 70",
            PageCount = 1,
        };

        var report2 = new PdfReport
        {
            Label = "AFS 2",
            Filename = "afs2.pdf",
            PdfPath = "missing-afs2.pdf",
            Pages = new() { "Revenue 120\nExpenses 70" },
            FullText = "Revenue 120\nExpenses 70",
            PageCount = 1,
        };

        var fast = await service.BuildComparisonAsync(report1, report2);

        Assert.Null(fast.SmartComparison);
        Assert.Null(fast.DocumentIntelligence);
        Assert.NotEmpty(fast.PageDiffs);
        Assert.NotEmpty(fast.FullDiff);

        var deferred = await service.BuildComparisonAsync(
            report1,
            report2,
            cancellationToken: default,
            includeDeferredInsights: true);

        Assert.NotNull(deferred.SmartComparison);
    }

    [Fact]
    public async Task ComparisonService_FindsContentChangeOnLaterPage()
    {
        const string pageOne =
            "The board reviews the annual financial statements and confirms that all supporting " +
            "records are complete and available for inspection by the audit team.";
        const string oldPageTwo =
            "Management reviewed the supporting documentation for significant estimates and assumptions " +
            "used to prepare the annual financial statements and concluded that reported amounts fairly " +
            "represent the entity's financial position and performance at the reporting date.";
        const string newPageTwo =
            "Management ignored the supporting documentation for significant estimates and assumptions " +
            "used to prepare the annual financial statements and concluded that reported amounts fairly " +
            "represent the entity's financial position and performance at the reporting date.";
        var pages1 = new List<string> { pageOne, oldPageTwo };
        var pages2 = new List<string> { pageOne, newPageTwo };
        var lineComparator = new LineComparatorService();
        var service = new ComparisonService(
            new PageAlignmentService(),
            lineComparator,
            new SmartComparisonOrchestrator(),
            new DocumentIntelligenceService(lineComparator),
            new AutoReviewEngine(new LabeledNumberExtractor()),
            new TableDiffService());

        var result = await service.BuildComparisonAsync(
            new PdfReport
            {
                Label = "AFS 1",
                Pages = pages1,
                FullText = string.Join("\n", pages1),
                PageCount = pages1.Count,
            },
            new PdfReport
            {
                Label = "AFS 2",
                Pages = pages2,
                FullText = string.Join("\n", pages2),
                PageCount = pages2.Count,
            });

        var laterPage = Assert.Single(result.PageDiffs.Where(page => page.PageAfs1 == 2));
        Assert.Equal(2, laterPage.PageAfs2);
        Assert.Contains(laterPage.Diffs, diff => diff.Status == "changed");
    }

    [Fact]
    public void LineComparatorService_Gate1c_Treats_IdenticalWordSetsAndNumbers_AsSame()
    {
        var result = InvokePrivateStatic<string>(
            typeof(LineComparatorService),
            "DetermineStatus",
            "Straight line 6 years",
            "6 years straight line");

        Assert.Equal("same", result);
    }

    [Fact]
    public void LineComparisonService_Gate1c_Treats_IdenticalWordSetsAndNumbers_AsSame()
    {
        var result = InvokePrivateStatic<DiffStatus>(
            typeof(LineComparisonService),
            "DetermineStatus",
            "Straight line 6 years",
            "6 years straight line");

        Assert.Equal(DiffStatus.Same, result);
    }

    [Fact]
    public void LineComparatorService_HighlightsSingleWordEditInLongDisclosure()
    {
        const string oldText =
            "Management reviewed the supporting documentation for significant estimates and assumptions " +
            "used to prepare the annual financial statements and concluded that reported amounts fairly " +
            "represent the entity's financial position and performance for the year ended 2025";
        const string newText =
            "Management ignored the supporting documentation for significant estimates and assumptions " +
            "used to prepare the annual financial statements and concluded that reported amounts fairly " +
            "represent the entity's financial position and performance for the year ended 2025";

        var result = Assert.Single(new LineComparatorService()
            .CompareLines(oldText, newText)
            .Where(diff => diff.Status == "changed"));

        Assert.Contains(result.WordDiff, word => word.Word == "reviewed" && word.Tag == "removed");
        Assert.Contains(result.WordDiff, word => word.Word == "ignored" && word.Tag == "added");
    }

    [Fact]
    public void LineComparatorService_DoesNotIgnoreChangedParentheticalContent()
    {
        var result = InvokePrivateStatic<string>(
            typeof(LineComparatorService),
            "DetermineStatus",
            "Management evaluated the estimated useful lives (for buildings) before approval",
            "Management evaluated the estimated useful lives (for vehicles) before approval");

        Assert.Equal("changed", result);
    }

    [Fact]
    public void LineComparatorService_DoesNotIgnoreWordOrderChanges()
    {
        var result = InvokePrivateStatic<string>(
            typeof(LineComparatorService),
            "DetermineStatus",
            "The audited figures include cash and receivables at year end",
            "The audited figures include receivables and cash at year end");

        Assert.Equal("changed", result);
    }

    [Fact]
    public void LineComparatorService_DoesNotTreatPartialWordAsWrap()
    {
        var result = InvokePrivateStatic<string>(
            typeof(LineComparatorService),
            "DetermineStatus",
            "The directors reached a final conclusion after examining every assess",
            "The directors reached a final conclusion after examining every assessment");

        Assert.Equal("changed", result);
    }

    [Fact]
    public void LineComparatorService_WhitespaceAndNumberFormattingOnly_IsSame()
    {
        var result = InvokePrivateStatic<string>(
            typeof(LineComparatorService),
            "DetermineStatus",
            "Total assets were R 1 234 567 at year end",
            "Total   assets were R1,234,567 at year end");

        Assert.Equal("same", result);
    }

    [Fact]
    public void LineComparatorService_OptionalPluralFormattingDoesNotCreateAHighlight()
    {
        var result = Assert.Single(new LineComparatorService()
            .CompareLines("Figures in Rand  Note(s)  2025  2024", "Figures in Rand  Note  2025  2024"));

        Assert.Equal("same", result.Status);
        Assert.Empty(result.WordDiff.Where(word => word.Tag != "same"));
    }

    [Theory]
    [InlineData("Total assets 1 500", "total assets 1,500.00", "same")]
    [InlineData("Date: 06 October 1990", "Date: 6 October 1990", "same")]
    [InlineData("Amount: R 150", "Amount: ZAR 150", "same")]
    [InlineData("Amount: 150", "Amount: (150)", "changed")]
    [InlineData("Rate: 10%", "Rate: 10", "changed")]
    [InlineData("Rate: 10%", "Rate: 12%", "changed")]
    [InlineData("Amount: R 150", "Amount: $150", "changed")]
    [InlineData("Amount: 1.5", "Amount: 15", "changed")]
    [InlineData("Rate: 5%", "Rate: 6%", "changed")]
    [InlineData("Amount: 1 500,00", "Amount: 1500.00", "same")]
    public void LineComparatorService_PreservesFinancialMeaning(
        string first,
        string second,
        string expected)
    {
        var result = InvokePrivateStatic<string>(
            typeof(LineComparatorService),
            "DetermineStatus",
            first,
            second);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("Total assets 1 500", "total assets 1,500.00", DiffStatus.Same)]
    [InlineData("Date: 06 October 1990", "Date: 6 October 1990", DiffStatus.Same)]
    [InlineData("Amount: R 150", "Amount: ZAR 150", DiffStatus.Same)]
    [InlineData("Amount: 150", "Amount: (150)", DiffStatus.Changed)]
    [InlineData("Rate: 10%", "Rate: 10", DiffStatus.Changed)]
    [InlineData("Rate: 10%", "Rate: 12%", DiffStatus.Changed)]
    [InlineData("Amount: R 150", "Amount: $150", DiffStatus.Changed)]
    [InlineData("Amount: 1.5", "Amount: 15", DiffStatus.Changed)]
    [InlineData("Rate: 5%", "Rate: 6%", DiffStatus.Changed)]
    [InlineData("Amount: 1 500,00", "Amount: 1500.00", DiffStatus.Same)]
    public void LineComparisonService_PreservesFinancialMeaning(
        string first,
        string second,
        DiffStatus expected)
    {
        var result = InvokePrivateStatic<DiffStatus>(
            typeof(LineComparisonService),
            "DetermineStatus",
            first,
            second);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void AuditTextNormalizer_EquatesEquivalentNumberSeparators()
    {
        var first = AuditTextNormalizer.ExtractNormalizedNumbers("Total assets 1 500");
        var second = AuditTextNormalizer.ExtractNormalizedNumbers("total assets 1,500.00");

        Assert.True(first.SetEquals(second), $"First: {string.Join(", ", first)}; second: {string.Join(", ", second)}");

        var spacedCurrency = AuditTextNormalizer.ExtractNormalizedNumbers(
            "Total assets were R 1 234 567 at year end");
        var groupedCurrency = AuditTextNormalizer.ExtractNormalizedNumbers(
            "Total assets were R1,234,567 at year end");
        Assert.True(spacedCurrency.SetEquals(groupedCurrency),
            $"Spaced: {string.Join(", ", spacedCurrency)}; grouped: {string.Join(", ", groupedCurrency)}");
        Assert.False(AuditTextNormalizer.HasDifferentMeaningfulMarkers(
            "Total assets were R 1 234 567 at year end",
            "Total assets were R1,234,567 at year end"));
    }

    [Fact]
    public void LineComparatorService_OrderedWrapExtension_IsSame()
    {
        var result = InvokePrivateStatic<string>(
            typeof(LineComparatorService),
            "DetermineStatus",
            "These financial statements have been prepared in accordance with the applicable accounting standards",
            "These financial statements have been prepared in accordance with the applicable accounting standards and related interpretations");

        Assert.Equal("same", result);
    }

    [Fact]
    public void LineComparisonService_HighlightsSingleWordEditInLongDisclosure()
    {
        const string oldText =
            "Management reviewed the supporting documentation for significant estimates and assumptions " +
            "used to prepare the annual financial statements and concluded that reported amounts fairly " +
            "represent the entity's financial position and performance for the year ended 2025";
        const string newText =
            "Management ignored the supporting documentation for significant estimates and assumptions " +
            "used to prepare the annual financial statements and concluded that reported amounts fairly " +
            "represent the entity's financial position and performance for the year ended 2025";

        var result = Assert.Single(new LineComparisonService()
            .CompareLines(oldText, newText)
            .Where(diff => diff.Status == DiffStatus.Changed));

        Assert.Contains("reviewed", result.Line1);
        Assert.Contains("ignored", result.Line2);
    }

    [Fact]
    public void LineComparisonService_DoesNotIgnoreChangedParentheticalContent()
    {
        var result = InvokePrivateStatic<DiffStatus>(
            typeof(LineComparisonService),
            "DetermineStatus",
            "Management evaluated the estimated useful lives (for buildings) before approval",
            "Management evaluated the estimated useful lives (for vehicles) before approval");

        Assert.Equal(DiffStatus.Changed, result);
    }

    [Fact]
    public void LineComparisonService_DoesNotIgnoreWordOrderChanges()
    {
        var result = InvokePrivateStatic<DiffStatus>(
            typeof(LineComparisonService),
            "DetermineStatus",
            "The audited figures include cash and receivables at year end",
            "The audited figures include receivables and cash at year end");

        Assert.Equal(DiffStatus.Changed, result);
    }

    [Fact]
    public void LineComparisonService_DoesNotTreatPartialWordAsWrap()
    {
        var result = InvokePrivateStatic<DiffStatus>(
            typeof(LineComparisonService),
            "DetermineStatus",
            "The directors reached a final conclusion after examining every assess",
            "The directors reached a final conclusion after examining every assessment");

        Assert.Equal(DiffStatus.Changed, result);
    }

    [Fact]
    public void LineComparisonService_WhitespaceAndNumberFormattingOnly_IsSame()
    {
        var result = InvokePrivateStatic<DiffStatus>(
            typeof(LineComparisonService),
            "DetermineStatus",
            "Total assets were R 1 234 567 at year end",
            "Total   assets were R1,234,567 at year end");

        Assert.Equal(DiffStatus.Same, result);
    }

    [Fact]
    public void LineComparisonService_OrderedWrapExtension_IsSame()
    {
        var result = InvokePrivateStatic<DiffStatus>(
            typeof(LineComparisonService),
            "DetermineStatus",
            "These financial statements have been prepared in accordance with the applicable accounting standards",
            "These financial statements have been prepared in accordance with the applicable accounting standards and related interpretations");

        Assert.Equal(DiffStatus.Same, result);
    }

    [Fact]
    public void SmartDiffEngine_ReportsSingleWordEditsInParagraphs()
    {
        const string oldText =
            "Management reviewed the supporting documentation for significant estimates and assumptions " +
            "used to prepare the annual financial statements and concluded that reported amounts fairly " +
            "represent the entity's financial position and performance at the reporting date";
        const string newText =
            "Management ignored the supporting documentation for significant estimates and assumptions " +
            "used to prepare the annual financial statements and concluded that reported amounts fairly " +
            "represent the entity's financial position and performance at the reporting date";

        var diffs = new SmartDiffEngine().Compare(
            SemanticBlockBuilder.Build(oldText, 1),
            SemanticBlockBuilder.Build(newText, 1));

        Assert.Contains(diffs, diff => diff.ChangeType == SmartChangeType.NarrativeRewrite);
    }

    [Fact]
    public void SmartDiffEngine_ReportsWordOrderChangesInParagraphs()
    {
        const string oldText =
            "The audited figures include cash and receivables at year end and provide a clear " +
            "summary of the organization's liquidity position and financial performance";
        const string newText =
            "The audited figures include receivables and cash at year end and provide a clear " +
            "summary of the organization's liquidity position and financial performance";

        var diffs = new SmartDiffEngine().Compare(
            SemanticBlockBuilder.Build(oldText, 1),
            SemanticBlockBuilder.Build(newText, 1));

        Assert.Contains(diffs, diff => diff.ChangeType == SmartChangeType.NarrativeRewrite);
    }

    [Fact]
    public void SmartDiffEngine_IgnoresWhitespaceOnlyParagraphDifferences()
    {
        const string oldText =
            "The company recognizes revenue when control of a promised good or service transfers to a customer.";
        const string newText =
            "The company recognizes   revenue when control of a promised good or service transfers to a customer.";

        var diffs = new SmartDiffEngine().Compare(
            SemanticBlockBuilder.Build(oldText, 1),
            SemanticBlockBuilder.Build(newText, 1));

        Assert.Empty(diffs);
    }

    [Fact]
    public void PageSnapshotService_GateZero_Suppresses_ExactMatchFormattingArtifact()
    {
        const string lineText = "The Company classifies its financial instruments into the following categories";
        string lineNorm = InvokePrivateStatic<string>(typeof(PageSnapshotService), "SnapNorm", lineText);
        var spec = new HighlightSpec
        {
            LineText = lineText,
            ChangedWords = new() { "Company", "categories" },
        };

        var shouldMatch = InvokePrivateStatic<bool>(
            typeof(PageSnapshotService),
            "ShouldMatchSnapshotSpec",
            lineNorm,
            spec);

        Assert.False(shouldMatch);
    }

    [Fact]
    public void PageSnapshotService_Allows_ExactMatch_For_WholeLineMode()
    {
        const string lineText = "Statement of Financial Position 17";
        string lineNorm = InvokePrivateStatic<string>(typeof(PageSnapshotService), "SnapNorm", lineText);
        var spec = new HighlightSpec
        {
            LineText = lineText,
            WholeLineMode = true,
            AnchorWords = new() { "Statement", "Financial", "Position" },
            ChangedPhrases = new() { lineText },
        };

        var shouldMatch = InvokePrivateStatic<bool>(
            typeof(PageSnapshotService),
            "ShouldMatchSnapshotSpec",
            lineNorm,
            spec);

        Assert.True(shouldMatch);
    }

    [Fact]
    public void PageSnapshotService_PerWordMatching_Strips_Punctuation_And_Spaces()
    {
        var changedSet = InvokePrivateStatic<HashSet<string>>(
            typeof(PageSnapshotService),
            "BuildChangedTokenSet",
            (object)new[] { "(92 451)" });

        var hit = InvokePrivateStatic<bool>(
            typeof(PageSnapshotService),
            "SnapshotWordMatchesChangedSet",
            "(92451)",
            changedSet);

        Assert.True(hit);
    }

    [Fact]
    public void AfsController_PageNumberOnlyDiff_UsesWholeLineMode_But_Keeps_Anchors()
    {
        var diff = new LineDiffResult
        {
            Status = "changed",
            Line1 = "pages 31 32",
            Line2 = "pages 40 41",
            WordDiff = new()
            {
                new WordDiffToken { Word = "pages", Tag = "same" },
                new WordDiffToken { Word = "31", Tag = "removed" },
                new WordDiffToken { Word = "32", Tag = "removed" },
                new WordDiffToken { Word = "40", Tag = "added" },
                new WordDiffToken { Word = "41", Tag = "added" },
            }
        };

        var spec = InvokePrivateStatic<HighlightSpec>(
            typeof(AfsController),
            "BuildHighlightSpec",
            diff,
            1);

        Assert.Empty(spec.ChangedWords);
        Assert.True(spec.WholeLineMode);
        Assert.NotEmpty(spec.AnchorWords);
        Assert.NotEmpty(spec.ChangedPhrases);
    }

    private static T InvokePrivateStatic<T>(Type type, string methodName, params object?[] args)
    {
        var methods = type
            .GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .Where(m => string.Equals(m.Name, methodName, StringComparison.Ordinal))
            .ToList();

        var method = methods
            .FirstOrDefault(m =>
            {
                var parameters = m.GetParameters();
                if (parameters.Length != args.Length)
                    return false;

                for (int i = 0; i < parameters.Length; i++)
                {
                    var arg = args[i];
                    if (arg is null)
                    {
                        if (parameters[i].ParameterType.IsValueType)
                            return false;

                        continue;
                    }

                    if (!parameters[i].ParameterType.IsAssignableFrom(arg.GetType()))
                        return false;
                }

                return true;
            })
            ?? methods.FirstOrDefault(m => m.GetParameters().Length == args.Length);

        Assert.NotNull(method);
        var result = method!.Invoke(null, args);
        return Assert.IsType<T>(result);
    }
}
