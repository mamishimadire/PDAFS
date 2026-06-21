using System.Text.Json;
using AfsPdfComparison.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace AfsPdfComparison.Services;

/// <summary>
/// Runs Gemini, Groq, and OpenRouter in parallel on the same comparison.
/// Merges their findings — changes spotted by multiple models get promoted
/// to higher severity; items any model flags that others missed are still kept.
/// </summary>
public class MultiAiOrchestrator
{
    private readonly AiComparisonAssistantService   _gemini;
    private readonly OpenAiCompatibleAiService      _openAiAdapter;
    private readonly IConfiguration                 _config;
    private readonly ILogger<MultiAiOrchestrator>  _logger;

    public MultiAiOrchestrator(
        AiComparisonAssistantService gemini,
        OpenAiCompatibleAiService    openAiAdapter,
        IConfiguration               config,
        ILogger<MultiAiOrchestrator> logger)
    {
        _gemini        = gemini;
        _openAiAdapter = openAiAdapter;
        _config        = config;
        _logger        = logger;
    }

    public async Task<AiAssistResult> AnalyzeAsync(
        ComparisonResult comparison,
        string           geminiPromptOverride = "",
        CancellationToken cancellationToken = default)
    {
        // ── Kick off all models in parallel ──────────────────────────────────
        var geminiTask = _gemini.AnalyzeAsync(comparison, cancellationToken);

        var groqTask   = RunGroqAsync(comparison, cancellationToken);
        var orTask     = RunOpenRouterAsync(comparison, cancellationToken);

        await Task.WhenAll(geminiTask, groqTask, orTask);

        var allResults = new[] { geminiTask.Result, groqTask.Result, orTask.Result };
        var results = allResults
            .Where(r => r.Status == "ready")
            .ToList();

        _logger.LogInformation(
            "[MultiAI] {Count}/3 models returned ready results.", results.Count);

        if (results.Count == 0)
        {
            _logger.LogWarning(
                "[MultiAI] No ready results. Gemini={GeminiStatus}, Groq={GroqStatus}, OpenRouter={OpenRouterStatus}.",
                geminiTask.Result.Status,
                groqTask.Result.Status,
                orTask.Result.Status);

            return AiUserFacingMessages.CreateUnavailableReviewResult(allResults);
        }

        return Merge(results);
    }

    // ── Groq ─────────────────────────────────────────────────────────────────
    private async Task<AiAssistResult> RunGroqAsync(
        ComparisonResult cmp, CancellationToken ct)
    {
        string? key = _config["GROQ_API_KEY"]
            ?? Environment.GetEnvironmentVariable("GROQ_API_KEY");
        if (string.IsNullOrWhiteSpace(key))
            return Skipped("Groq", "GROQ_API_KEY not set.");

        string model    = _config["AiAssist:GroqModel"] ?? "llama-3.3-70b-versatile";
        string endpoint = _config["AiAssist:GroqUrl"]   ?? "https://api.groq.com/openai/v1/chat/completions";
        string prompt   = BuildOpenAiPrompt(cmp);

        return await _openAiAdapter.AnalyzeAsync(endpoint, key, model, "Groq", prompt, 60, ct);
    }

    // ── OpenRouter ────────────────────────────────────────────────────────────
    private async Task<AiAssistResult> RunOpenRouterAsync(
        ComparisonResult cmp, CancellationToken ct)
    {
        string? key = _config["OPENROUTER_API_KEY"]
            ?? Environment.GetEnvironmentVariable("OPENROUTER_API_KEY");
        if (string.IsNullOrWhiteSpace(key))
            return Skipped("OpenRouter", "OPENROUTER_API_KEY not set.");

        string model    = _config["AiAssist:OpenRouterModel"] ?? "deepseek/deepseek-chat";
        string endpoint = _config["AiAssist:OpenRouterUrl"]   ?? "https://openrouter.ai/api/v1/chat/completions";
        string prompt   = BuildOpenAiPrompt(cmp);

        return await _openAiAdapter.AnalyzeAsync(endpoint, key, model, "OpenRouter", prompt, 90, ct);
    }

    // ── Prompt for chat-completion models (no JSON schema enforcement) ────────
    private static string BuildOpenAiPrompt(ComparisonResult cmp)
    {
        var pagePairs = cmp.PageDiffs
            .Where(p => p.IssueCount > 0)
            .OrderByDescending(p => p.IssueCount)
            .Take(18)
            .Select((p, idx) => new
            {
                pairIndex  = p.PairIndex,
                pageAfs1   = p.PageAfs1,
                pageAfs2   = p.PageAfs2,
                pctSame    = p.PctSame,
                changed    = p.Changed,
                added      = p.Added,
                removed    = p.Removed,
                diffs      = p.Diffs
                    .Where(d => d.Status != "same" &&
                                d.Status != "same-cross-page" &&
                                d.Status != "same-different-page")
                    .Take(10)
                    .Select((d, i) => new
                    {
                        ordinal = i + 1,
                        status  = d.Status,
                        afs1    = Trim(d.Line1, 300),
                        afs2    = Trim(d.Line2, 300),
                        nums    = d.NumDiff.Take(6).ToList(),
                    })
            });

        // Include page text for the most-changed pages
        var pageContent = cmp.PageDiffs
            .Where(p => p.IssueCount > 0)
            .OrderByDescending(p => p.IssueCount)
            .Take(4)
            .Select(p => new
            {
                pairIndex = p.PairIndex,
                textAfs1  = p.PageAfs1.HasValue && p.PageAfs1.Value - 1 < cmp.Report1.Pages.Count
                    ? Trim(cmp.Report1.Pages[p.PageAfs1.Value - 1], 1400) : null,
                textAfs2  = p.PageAfs2.HasValue && p.PageAfs2.Value - 1 < cmp.Report2.Pages.Count
                    ? Trim(cmp.Report2.Pages[p.PageAfs2.Value - 1], 1400) : null,
            });

        var ctx = new
        {
            task = "You are an expert financial auditor AI at SNG Grant Thornton. " +
                   "Compare two versions of an Annual Financial Statement (AFS). " +
                   "Read the ACTUAL page content samples carefully — do not rely only on the diff output. " +
                   "Identify ALL material changes including: directorate changes (appointments/resignations/date changes), " +
                   "numeric changes, policy changes, added/removed disclosures, and any other differences.",
            report1 = new { cmp.Report1.Label, cmp.Report1.Filename, pages = cmp.Report1.PageCount },
            report2 = new { cmp.Report2.Label, cmp.Report2.Filename, pages = cmp.Report2.PageCount },
            counts  = new
            {
                same    = cmp.Counts.GetValueOrDefault("same", 0),
                changed = cmp.Counts.GetValueOrDefault("changed", 0),
                added   = cmp.Counts.GetValueOrDefault("added", 0),
                removed = cmp.Counts.GetValueOrDefault("removed", 0),
            },
            pageContent,
            pagePairs,
        };

        string schema = """
{
  "executiveSummary": "string",
  "overallRisk": "High|Medium|Low",
  "materialFindings": [
    {
      "title": "string",
      "severity": "High|Medium|Low",
      "pairIndex": 0,
      "pageAfs1": 0,
      "pageAfs2": 0,
      "focusDiffOrdinals": [1],
      "auditRationale": "string",
      "recommendedProcedure": "string",
      "changedNumbers": ["string"]
    }
  ],
  "pageFocus": [
    { "pairIndex": 0, "pageAfs1": 0, "pageAfs2": 0, "priority": "High|Medium|Low", "why": "string", "focusDiffOrdinals": [1] }
  ],
  "highlightGuidance": ["string"],
  "limitations": ["string"]
}
""";

        return $"Analyse the AFS comparison below. Return ONLY valid JSON matching this exact schema:\n{schema}\n\n" +
               JsonSerializer.Serialize(ctx, new JsonSerializerOptions { WriteIndented = false });
    }

    // ── Merge multiple model results into one ─────────────────────────────────
    private static AiAssistResult Merge(List<AiAssistResult> results)
    {
        // Executive summary: longest one wins (most detail)
        var best = results.OrderByDescending(r => r.ExecutiveSummary?.Length ?? 0).First();

        // Risk: take the highest across models
        static int RiskRank(string? r) => (r ?? "").ToLower() switch
        {
            "high" => 2, "medium" => 1, _ => 0
        };
        var topRisk = results.OrderByDescending(r => RiskRank(r.OverallRisk)).First().OverallRisk;

        // Merge material findings — group by pairIndex+title similarity
        var allFindings = results.SelectMany(r => r.MaterialFindings ?? new()).ToList();
        var mergedFindings = MergeFindings(allFindings, results.Count);

        // Merge pageFocus — take union, promote priority if multiple models agree
        var allFocus = results.SelectMany(r => r.PageFocus ?? new()).ToList();
        var mergedFocus = MergeFocus(allFocus, results.Count);

        // Merge highlight guidance — union distinct items
        var guidance = results
            .SelectMany(r => r.HighlightGuidance ?? new())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(10)
            .ToList();

        // Limitations
        var providers = string.Join(", ", results.Select(r => r.Provider));
        var limitations = new List<string>
            { $"Analysis from {results.Count} model(s): {providers}. Results merged by consensus." };
        limitations.AddRange(results
            .Where(r => r.Limitations?.Any() == true)
            .SelectMany(r => r.Limitations!)
            .Distinct()
            .Take(4));

        return new AiAssistResult
        {
            Enabled          = true,
            Provider         = string.Join(" + ", results.Select(r => r.Provider).Distinct()),
            Model            = string.Join(", ", results.Select(r => r.Model).Distinct()),
            Status           = "ready",
            Message          = $"Consensus analysis from {results.Count} model(s).",
            ExecutiveSummary = best.ExecutiveSummary,
            OverallRisk      = topRisk,
            MaterialFindings = mergedFindings,
            PageFocus        = mergedFocus,
            HighlightGuidance = guidance,
            Limitations      = limitations,
            GeneratedAt      = DateTime.UtcNow,
        };
    }

    private static List<AiMaterialFinding> MergeFindings(
        List<AiMaterialFinding> all, int modelCount)
    {
        if (!all.Any()) return new();

        // Group by pairIndex — findings on the same page pair from different models
        var grouped = all
            .GroupBy(f => f.PairIndex ?? -1)
            .SelectMany(g =>
            {
                // Within a page pair, try to cluster by title similarity
                var items = g.ToList();
                if (items.Count == 1) return items;

                // Merge clusters: if 2+ models flagged same pair, boost severity
                var merged = new List<AiMaterialFinding>();
                var used   = new HashSet<int>();
                for (int i = 0; i < items.Count; i++)
                {
                    if (used.Contains(i)) continue;
                    var cluster = new List<AiMaterialFinding> { items[i] };
                    used.Add(i);
                    for (int j = i + 1; j < items.Count; j++)
                    {
                        if (used.Contains(j)) continue;
                        if (AreSimilarFindings(items[i], items[j]))
                        {
                            cluster.Add(items[j]);
                            used.Add(j);
                        }
                    }
                    merged.Add(MergeCluster(cluster, modelCount));
                }
                return merged;
            })
            .OrderByDescending(f => SevRank(f.Severity))
            .ThenBy(f => f.PairIndex ?? 999)
            .Take(20)
            .ToList();

        return grouped;
    }

    private static AiMaterialFinding MergeCluster(
        List<AiMaterialFinding> cluster, int totalModels)
    {
        var best = cluster.OrderByDescending(f => SevRank(f.Severity))
                          .ThenByDescending(f => f.AuditRationale?.Length ?? 0)
                          .First();

        // Boost severity if multiple models agree
        string severity = best.Severity;
        if (cluster.Count >= 2 && SevRank(severity) < 2)
            severity = SevRank(severity) == 1 ? "High" : "Medium";

        // Union of ordinals from all models in the cluster
        var ordinals = cluster
            .SelectMany(f => f.FocusDiffOrdinals ?? new())
            .Distinct().OrderBy(x => x).ToList();

        // Union of changed numbers
        var nums = cluster
            .SelectMany(f => f.ChangedNumbers ?? new())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Take(10)
            .ToList();

        int agreeCount = cluster.Count;
        string consensusNote = agreeCount > 1
            ? $" [{agreeCount}/{totalModels} models agreed on this finding]"
            : string.Empty;

        return new AiMaterialFinding
        {
            Title                = best.Title,
            Severity             = severity,
            PairIndex            = best.PairIndex,
            PageAfs1             = best.PageAfs1,
            PageAfs2             = best.PageAfs2,
            FocusDiffOrdinals    = ordinals,
            AuditRationale       = (best.AuditRationale ?? "") + consensusNote,
            RecommendedProcedure = best.RecommendedProcedure,
            ChangedNumbers       = nums,
        };
    }

    private static List<AiPageFocus> MergeFocus(
        List<AiPageFocus> all, int modelCount)
    {
        return all
            .GroupBy(f => f.PairIndex)
            .Select(g =>
            {
                var best     = g.OrderByDescending(f => SevRank(f.Priority)).First();
                var ordinals = g.SelectMany(f => f.FocusDiffOrdinals ?? new()).Distinct().OrderBy(x => x).ToList();
                // If multiple models flagged this page, promote priority
                if (g.Count() >= 2 && SevRank(best.Priority) < 2)
                    best = new AiPageFocus
                    {
                        PairIndex         = best.PairIndex,
                        PageAfs1          = best.PageAfs1,
                        PageAfs2          = best.PageAfs2,
                        Priority          = SevRank(best.Priority) == 1 ? "High" : "Medium",
                        Why               = best.Why + $" [{g.Count()}/{modelCount} models]",
                        FocusDiffOrdinals = ordinals,
                    };
                else
                    best.FocusDiffOrdinals = ordinals;
                return best;
            })
            .OrderByDescending(f => SevRank(f.Priority))
            .Take(15)
            .ToList();
    }

    private static bool AreSimilarFindings(AiMaterialFinding a, AiMaterialFinding b)
    {
        if (a.PairIndex != b.PairIndex) return false;
        // If same page pair and title shares ≥2 significant words, treat as same finding
        var wordsA = Tokenise(a.Title);
        var wordsB = Tokenise(b.Title);
        int common = wordsA.Intersect(wordsB, StringComparer.OrdinalIgnoreCase).Count();
        return common >= 2;
    }

    private static HashSet<string> Tokenise(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return new();
        return text.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                   .Where(w => w.Length >= 4)
                   .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static int SevRank(string? s) => (s ?? "").ToLower() switch
    {
        "high" => 2, "medium" => 1, _ => 0
    };

    private static string Trim(string? s, int max)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        string t = s.Trim();
        return t.Length <= max ? t : t[..max] + "...";
    }

    private static AiAssistResult Skipped(string provider, string msg) => new()
    {
        Enabled = false, Provider = provider, Status = "disabled", Message = msg
    };
}
