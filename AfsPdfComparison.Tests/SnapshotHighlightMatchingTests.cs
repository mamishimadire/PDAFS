using System.Reflection;
using AfsPdfComparison.Controllers;
using AfsPdfComparison.Models;
using AfsPdfComparison.Services;

namespace AfsPdfComparison.Tests;

public class SnapshotHighlightMatchingTests
{
    private const string UnchangedResponsibilitiesParagraph =
        "The Provincial Executive Committee (PEC) are required by the SADTU Constitution as amended " +
        "in September 2024, to maintain adequate accounting records and are responsible for the content " +
        "and integrity of the special purpose financial statements and related financial information " +
        "included in this report.";

    [Fact]
    public void ControllerRenderer_DoesNotMatchUnchangedParagraphToUnrelatedDiff()
    {
        var spec = new HighlightSpec
        {
            LineText = "The external auditors examined the special purpose financial statements and " +
                       "presented their report on pages 5 to 8.",
            ChangedWords = new() { "8" },
            AnchorWords = new() { "examined", "presented", "pages" },
            ChangedPhrases = new()
            {
                "examined the special purpose financial statements and presented their report"
            },
            WholeLineMode = true,
        };

        Assert.False(ControllerMatch(UnchangedResponsibilitiesParagraph, spec));
    }

    [Fact]
    public void ControllerRenderer_MatchesExactWholeLineDiff()
    {
        const string changedLine = "Statement of Cash Flows 12";
        var spec = new HighlightSpec
        {
            LineText = changedLine,
            AnchorWords = new() { "Statement", "Cash", "Flows" },
            WholeLineMode = true,
        };

        Assert.True(ControllerMatch(changedLine, spec));
    }

    [Fact]
    public void SnapshotFallback_DoesNotMatchByHighUnorderedWordOverlap()
    {
        const string pageLine =
            "The special purpose financial statements were prepared in accordance with the accounting " +
            "policies described in Note 1 and present fairly the results for the reporting period ended " +
            "31 December 2024.";
        var spec = new HighlightSpec
        {
            LineText = pageLine.Replace("2024", "2025", StringComparison.Ordinal),
            ChangedWords = new() { "2025" },
            AnchorWords = new() { "period", "ended", "December" },
        };

        string lineNorm = InvokePrivateStatic<string>(
            typeof(PageSnapshotService),
            "SnapNorm",
            pageLine);

        Assert.False(InvokePrivateStatic<bool>(
            typeof(PageSnapshotService),
            "ShouldMatchSnapshotSpec",
            lineNorm,
            spec));
    }

    private static bool ControllerMatch(string lineText, HighlightSpec spec)
    {
        string lineNorm = InvokePrivateStatic<string>(typeof(AfsController), "SnapNorm", lineText);
        string specNorm = InvokePrivateStatic<string>(
            typeof(AfsController),
            "SnapNorm",
            spec.LineText);
        var lineTokens = lineText
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Select(token => InvokePrivateStatic<string>(
                typeof(AfsController),
                "SnapTokenNormRaw",
                token))
            .Where(token => token.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        return InvokePrivateStatic<bool>(
            typeof(AfsController),
            "HasStrongSnapshotLineMatch",
            lineNorm,
            specNorm,
            spec,
            lineTokens,
            0.0);
    }

    private static T InvokePrivateStatic<T>(Type type, string methodName, params object?[] args)
    {
        var method = type
            .GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .First(candidate =>
            {
                if (!string.Equals(candidate.Name, methodName, StringComparison.Ordinal))
                    return false;

                var parameters = candidate.GetParameters();
                return parameters.Length == args.Length &&
                       parameters.Select((parameter, index) =>
                               args[index] is null ||
                               parameter.ParameterType.IsAssignableFrom(args[index]!.GetType()))
                           .All(matches => matches);
            });

        return (T)method.Invoke(null, args)!;
    }
}
