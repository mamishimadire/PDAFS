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
        int readyResultTarget = Math.Max(1, AiConfiguration.GetInt(
            _config,
            1,
            "AiAssist:ReadyResultTarget",
            "AI_READY_RESULT_TARGET"));
        int consensusWindowMs = Math.Max(0, AiConfiguration.GetInt(
            _config,
            0,
            "AiAssist:ConsensusWindowMs",
            "AI_CONSENSUS_WINDOW_MS"));

        var attempts = new[]
        {
            new AiProviderAttempt<AiAssistResult>("Gemini",      ct => _gemini.AnalyzeAsync(comparison, ct)),
            new AiProviderAttempt<AiAssistResult>("Groq",        ct => RunGroqAsync(comparison, ct)),
            new AiProviderAttempt<AiAssistResult>("OpenRouter",  ct => RunOpenRouterAsync(comparison, ct)),
        };

        AiFailoverOutcome<AiAssistResult> outcome = await AiFailoverExecutor.ExecuteAsync(
            attempts,
            result => result.Status,
            readyResultTarget,
            TimeSpan.FromMilliseconds(consensusWindowMs),
            cancellationToken);

        var allResults = outcome.CompletedResults.ToList();
        var results    = outcome.ReadyResults.ToList();

        _logger.LogInformation(
            "[MultiAI] {Count}/3 models returned ready results within the failover window ({WindowMs} ms).",
            results.Count,
            consensusWindowMs);

        if (results.Count == 0)
        {
            string geminiStatus      = allResults.FirstOrDefault(r => string.Equals(r.Provider, "Gemini",      StringComparison.OrdinalIgnoreCase))?.Status ?? "not_started";
            string groqStatus        = allResults.FirstOrDefault(r => string.Equals(r.Provider, "Groq",        StringComparison.OrdinalIgnoreCase))?.Status ?? "not_started";
            string openRouterStatus  = allResults.FirstOrDefault(r => string.Equals(r.Provider, "OpenRouter",  StringComparison.OrdinalIgnoreCase))?.Status ?? "not_started";

            _logger.LogWarning(
                "[MultiAI] All models failed. Gemini={GeminiStatus}, Groq={GroqStatus}, OpenRouter={OpenRouterStatus}. Retrying with lean prompt.",
                geminiStatus, groqStatus, openRouterStatus);

            // ── Retry once with a leaner prompt (fewer diffs, less text) ──────
            // This handles the case where the standard prompt was too large for the model.
            bool anyKeyConfigured = allResults.Any(r =>
                !string.Equals(r.Status, "disabled",    StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(r.Status, "missing_key", StringComparison.OrdinalIgnoreCase));

            if (anyKeyConfigured && !cancellationToken.IsCancellationRequested)
            {
                var retryAttempts = new[]
                {
                    new AiProviderAttempt<AiAssistResult>("Gemini-Lean",     ct => _gemini.AnalyzeAsync(comparison, ct, lean: true)),
                    new AiProviderAttempt<AiAssistResult>("Groq-Lean",       ct => RunGroqAsync(comparison, ct, lean: true)),
                    new AiProviderAttempt<AiAssistResult>("OpenRouter-Lean", ct => RunOpenRouterAsync(comparison, ct, lean: true)),
                };

                var retryOutcome = await AiFailoverExecutor.ExecuteAsync(
                    retryAttempts,
                    result => result.Status,
                    readyResultTarget: 1,
                    waitAfterFirstReady: TimeSpan.Zero,
                    cancellationToken);

                var retryReady = retryOutcome.ReadyResults.ToList();
                if (retryReady.Count > 0)
                {
                    _logger.LogInformation("[MultiAI] Lean-prompt retry succeeded via {Provider}.", retryReady[0].Provider);
                    return retryReady[0];
                }

                allResults.AddRange(retryOutcome.CompletedResults);
            }

            return AiUserFacingMessages.CreateUnavailableReviewResult(allResults);
        }

        return results.Count == 1 ? results[0] : Merge(results);
    }

    // ── Groq ─────────────────────────────────────────────────────────────────
    private async Task<AiAssistResult> RunGroqAsync(
        ComparisonResult cmp, CancellationToken ct, bool lean = false)
    {
        string? key = AiConfiguration.GetOptional(_config, "AiAssist:GroqApiKey", "GROQ_API_KEY");
        if (string.IsNullOrWhiteSpace(key))
            return Skipped("Groq", "GROQ_API_KEY not set.");

        string model = AiConfiguration.GetString(
            _config,
            "llama-3.3-70b-versatile",
            "AiAssist:GroqModel",
            "GROQ_MODEL");
        string endpoint = AiConfiguration.GetString(
            _config,
            "https://api.groq.com/openai/v1/chat/completions",
            "AiAssist:GroqUrl",
            "GROQ_URL");
        string prompt   = BuildOpenAiPrompt(cmp, lean);

        return await _openAiAdapter.AnalyzeAsync(endpoint, key, model, "Groq", prompt, 60, ct);
    }

    // ── OpenRouter ────────────────────────────────────────────────────────────
    private async Task<AiAssistResult> RunOpenRouterAsync(
        ComparisonResult cmp, CancellationToken ct, bool lean = false)
    {
        string? key = AiConfiguration.GetOptional(_config, "AiAssist:OpenRouterApiKey", "OPENROUTER_API_KEY");
        if (string.IsNullOrWhiteSpace(key))
            return Skipped("OpenRouter", "OPENROUTER_API_KEY not set.");

        string model = AiConfiguration.GetString(
            _config,
            "deepseek/deepseek-chat",
            "AiAssist:OpenRouterModel",
            "OPENROUTER_MODEL");
        string endpoint = AiConfiguration.GetString(
            _config,
            "https://openrouter.ai/api/v1/chat/completions",
            "AiAssist:OpenRouterUrl",
            "OPENROUTER_URL");
        string prompt   = BuildOpenAiPrompt(cmp, lean);

        return await _openAiAdapter.AnalyzeAsync(endpoint, key, model, "OpenRouter", prompt, 90, ct);
    }

    // ── Prompt for chat-completion models (no JSON schema enforcement) ────────
    private static string BuildOpenAiPrompt(ComparisonResult cmp, bool lean = false)
    {
        int pairCap = lean ? 6  : 18;
        int diffCap = lean ? 6  : 10;
        int textCap = lean ? 800 : 1400;

        var pagePairs = cmp.PageDiffs
            .Where(p => p.IssueCount > 0)
            .OrderByDescending(p => p.IssueCount)
            .Take(pairCap)
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
                    .Take(diffCap)
                    .Select((d, i) => new
                    {
                        ordinal  = i + 1,
                        status   = d.Status,
                        afs1     = Trim(d.Line1, 250),
                        afs2     = Trim(d.Line2, 250),
                        wordDiff = d.WordDiff.Select(w => new { word = w.Word, tag = w.Tag }).Take(15).ToList(),
                        nums     = d.NumDiff.Take(6).ToList(),
                    })
            });

        // Page text for the most-changed pages
        var pageContent = cmp.PageDiffs
            .Where(p => p.IssueCount > 0)
            .OrderByDescending(p => p.IssueCount)
            .Take(lean ? 2 : 4)
            .Select(p => new
            {
                pairIndex = p.PairIndex,
                textAfs1  = p.PageAfs1.HasValue && p.PageAfs1.Value - 1 < cmp.Report1.Pages.Count
                    ? Trim(cmp.Report1.Pages[p.PageAfs1.Value - 1], textCap) : null,
                textAfs2  = p.PageAfs2.HasValue && p.PageAfs2.Value - 1 < cmp.Report2.Pages.Count
                    ? Trim(cmp.Report2.Pages[p.PageAfs2.Value - 1], textCap) : null,
            });

        // Compact diff report — covers every change with ~4x better token efficiency than JSON
        int lineCap = lean ? 50 : int.MaxValue;
        string compactDiffReport = BuildCompactDiffReport(cmp, lineCap);

        var pageAlignmentSummary = cmp.Alignment.Select((a, i) => new
        {
            afs1Page = a.I1 + 1,
            afs2Page = a.I2 >= 0 ? (int?)(a.I2 + 1) : null,
            sim      = Math.Round(a.Sim, 3),
        }).ToList();

        var systemReviewFindings = cmp.AutoReview.Take(lean ? 40 : int.MaxValue).Select(f => new
        {
            f.Category,
            f.Title,
            f.Severity,
            f.PageAfs1,
            f.PageAfs2,
            f.PairIndex,
            f.Before,
            f.After,
            pctChange = f.PctChange.HasValue ? Math.Round(f.PctChange.Value, 1) : (double?)null,
        }).ToList();

        var ctx = new
        {
            task = "You are an expert financial auditor AI at SNG Grant Thornton. " +
                   "You have: (1) compactDiffReport — every line changed/added/removed, " +
                   "and (2) systemReviewFindings — deterministic findings already categorised by the comparison engine. " +
                   "Start from systemReviewFindings: convert each into a materialFinding with a professional auditRationale " +
                   "and recommendedProcedure. Then check compactDiffReport for any genuine changes the system missed. " +
                   "Use pageAfs1/pageAfs2 from systemReviewFindings for correct snapshot references. " +
                   "No cap on findings — report ALL genuine material changes.",
            report1 = new { cmp.Report1.Label, cmp.Report1.Filename, pages = cmp.Report1.PageCount },
            report2 = new { cmp.Report2.Label, cmp.Report2.Filename, pages = cmp.Report2.PageCount },
            counts  = new
            {
                same    = cmp.Counts.GetValueOrDefault("same", 0),
                changed = cmp.Counts.GetValueOrDefault("changed", 0),
                added   = cmp.Counts.GetValueOrDefault("added", 0),
                removed = cmp.Counts.GetValueOrDefault("removed", 0),
            },
            systemReviewFindings,
            compactDiffReport,
            pageAlignmentSummary,
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

    private static string BuildCompactDiffReport(ComparisonResult cmp, int lineCap)
    {
        var sb = new System.Text.StringBuilder(8192);

        var lineToPage = new Dictionary<LineDiffResult, (int? p1, int? p2)>(
            ReferenceEqualityComparer.Instance);
        foreach (var pd in cmp.PageDiffs)
            foreach (var d in pd.Diffs)
                lineToPage.TryAdd(d, (pd.PageAfs1, pd.PageAfs2));

        // Changed
        var changed = cmp.ChangedLines.Take(lineCap).ToList();
        int total   = cmp.ChangedLines.Count();
        sb.AppendLine($"=== CHANGED LINES ({total} total" +
                      (changed.Count < total ? $", showing {changed.Count}" : "") + ") ===");
        int ord = 0;
        foreach (var d in changed)
        {
            ord++;
            lineToPage.TryGetValue(d, out var pg);
            string pageRef = pg.p1.HasValue || pg.p2.HasValue ? $"p{pg.p1}→p{pg.p2}" : "?";
            string suffix = "";
            if (d.NumDiff.Count > 0)  suffix += $" | nums:[{string.Join(",", d.NumDiff.Take(6))}]";
            var rem = d.WordDiff.Where(w => w.Tag == "removed").Select(w => w.Word).Take(5).ToList();
            var add = d.WordDiff.Where(w => w.Tag == "added").Select(w => w.Word).Take(5).ToList();
            if (rem.Count > 0) suffix += $" | -[{string.Join(",", rem)}]";
            if (add.Count > 0) suffix += $" | +[{string.Join(",", add)}]";
            sb.AppendLine($"CHG-{ord:D3} [{pageRef}]: \"{Trim(d.Line1, 120)}\" → \"{Trim(d.Line2, 120)}\"{suffix}");
        }

        // Added
        var added    = cmp.AddedLines.Take(lineCap).ToList();
        int addTotal = cmp.AddedLines.Count();
        sb.AppendLine();
        sb.AppendLine($"=== LINES ADDED IN AFS 2 ({addTotal} total" +
                      (added.Count < addTotal ? $", showing {added.Count}" : "") + ") ===");
        ord = 0;
        foreach (var d in added)
        {
            ord++;
            lineToPage.TryGetValue(d, out var pg);
            string pageRef = pg.p2.HasValue ? $"AFS2-p{pg.p2}" : "AFS2-?";
            sb.AppendLine($"ADD-{ord:D3} [{pageRef}]: \"{Trim(d.Line2, 150)}\"");
        }

        // Removed
        var removed  = cmp.RemovedLines.Take(lineCap).ToList();
        int remTotal = cmp.RemovedLines.Count();
        sb.AppendLine();
        sb.AppendLine($"=== LINES REMOVED FROM AFS 1 ({remTotal} total" +
                      (removed.Count < remTotal ? $", showing {removed.Count}" : "") + ") ===");
        ord = 0;
        foreach (var d in removed)
        {
            ord++;
            lineToPage.TryGetValue(d, out var pg);
            string pageRef = pg.p1.HasValue ? $"AFS1-p{pg.p1}" : "AFS1-?";
            sb.AppendLine($"REM-{ord:D3} [{pageRef}]: \"{Trim(d.Line1, 150)}\"");
        }

        return sb.ToString();
    }

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
