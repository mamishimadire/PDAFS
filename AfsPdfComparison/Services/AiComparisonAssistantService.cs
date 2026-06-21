using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using AfsPdfComparison.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace AfsPdfComparison.Services;

public class AiComparisonAssistantService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private readonly IHttpClientFactory _httpClientFactory;
    private readonly IConfiguration _configuration;
    private readonly ILogger<AiComparisonAssistantService> _logger;

    public AiComparisonAssistantService(
        IHttpClientFactory httpClientFactory,
        IConfiguration configuration,
        ILogger<AiComparisonAssistantService> logger)
    {
        _httpClientFactory = httpClientFactory;
        _configuration = configuration;
        _logger = logger;
    }

    public async Task<AiAssistResult> AnalyzeAsync(ComparisonResult comparison, CancellationToken cancellationToken = default)
    {
        bool enabled = GetBool("AiAssist:Enabled", defaultValue: true);
        string provider = _configuration["AiAssist:Provider"] ?? "Gemini";
        string model = _configuration["AiAssist:GeminiModel"] ?? "gemini-2.5-flash";

        if (!enabled)
        {
            return Unavailable(provider, model, "disabled", enabled: false);
        }

        string apiKey =
            _configuration["AiAssist:GeminiApiKey"] ??
            _configuration["GEMINI_API_KEY"] ??
            Environment.GetEnvironmentVariable("GEMINI_API_KEY") ??
            string.Empty;

        if (string.IsNullOrWhiteSpace(apiKey))
        {
            return Unavailable(provider, model, "missing_key", enabled: false);
        }

        try
        {
            string modelPath = model.StartsWith("models/", StringComparison.OrdinalIgnoreCase)
                ? model
                : "models/" + model;

            string endpoint = $"https://generativelanguage.googleapis.com/v1beta/{modelPath}:generateContent";
            string prompt = BuildPrompt(comparison);
            object requestBody = new
            {
                contents = new[]
                {
                    new
                    {
                        parts = new[]
                        {
                            new { text = prompt }
                        }
                    }
                },
                generationConfig = new
                {
                    temperature = 0.2,
                    responseMimeType = "application/json",
                    responseJsonSchema = BuildResponseSchema(),
                }
            };

            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(TimeSpan.FromSeconds(90));

            using var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(requestBody, JsonOptions),
                    Encoding.UTF8,
                    "application/json")
            };
            request.Headers.Add("x-goog-api-key", apiKey);

            var client = _httpClientFactory.CreateClient();
            using HttpResponseMessage response = await client.SendAsync(request, cts.Token);
            string raw = await response.Content.ReadAsStringAsync(cts.Token);

            if (!response.IsSuccessStatusCode)
            {
                _logger.LogWarning("Gemini AI review failed: HTTP {StatusCode} {Body}", (int)response.StatusCode, Trim(raw, 600));
                return Unavailable(provider, model, "error");
            }

            string jsonText = ExtractModelText(raw);
            if (string.IsNullOrWhiteSpace(jsonText))
            {
                _logger.LogWarning("Gemini AI review returned no structured analysis.");
                return Unavailable(provider, model, "error");
            }

            string normalizedJson = StripCodeFences(jsonText);
            var aiResult = JsonSerializer.Deserialize<AiAssistResult>(normalizedJson, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            }) ?? new AiAssistResult();

            aiResult.Enabled = true;
            aiResult.Provider = provider;
            aiResult.Model = model;
            aiResult.Status = "ready";
            aiResult.Message = "AI review completed.";
            aiResult.GeneratedAt = DateTime.UtcNow;
            Normalize(aiResult);
            return aiResult;
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("Gemini AI review timed out after {TimeoutSeconds}s.", 90);
            return Unavailable(provider, model, "error");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Gemini AI review failed unexpectedly.");
            return Unavailable(provider, model, "error");
        }
    }

    private string BuildPrompt(ComparisonResult comparison)
    {
        // Build page content samples for the most-changed pages so Gemini can
        // read the actual document text, not just the diff statistics.
        var topChangedPairs = comparison.PageDiffs
            .Where(p => p.IssueCount > 0)
            .OrderByDescending(p => p.IssueCount)
            .Take(6)
            .ToList();

        var pageContentSamples = topChangedPairs.Select(p => new
        {
            pairIndex = p.PairIndex,
            pageAfs1 = p.PageAfs1,
            pageAfs2 = p.PageAfs2,
            contentAfs1 = p.PageAfs1.HasValue && p.PageAfs1.Value - 1 < comparison.Report1.Pages.Count
                ? Trim(comparison.Report1.Pages[p.PageAfs1.Value - 1], 1200)
                : null,
            contentAfs2 = p.PageAfs2.HasValue && p.PageAfs2.Value - 1 < comparison.Report2.Pages.Count
                ? Trim(comparison.Report2.Pages[p.PageAfs2.Value - 1], 1200)
                : null,
        }).ToList();

        // Also include near-identical page pairs (pctSame ≥ 98%) so AI can
        // validate whether the single detected diff is a false positive or real.
        var nearIdenticalPairs = comparison.PageDiffs
            .Where(p => p.IssueCount > 0 && p.PctSame >= 98)
            .Select(p => new
            {
                pairIndex = p.PairIndex,
                p.PageAfs1,
                p.PageAfs2,
                p.PctSame,
                p.IssueCount,
                issueDiffs = p.Diffs
                    .Where(d => d.Status != "same" && d.Status != "same-cross-page" && d.Status != "same-different-page")
                    .Select((d, i) => new { ordinal = i + 1, status = d.Status, line1 = Trim(d.Line1, 220), line2 = Trim(d.Line2, 220) })
                    .ToList()
            })
            .Take(10)
            .ToList();

        var context = new
        {
            task = "You are an audit-review AI. Read the actual page content samples from both AFS documents AND the deterministic diff output. " +
                   "Identify material changes, prioritise page pairs for review, flag which diff ordinals deserve highlighting, and explain WHAT changed and WHY it matters. " +
                   "IMPORTANT: also validate whether detected differences on near-identical pages (pctSame ≥ 98%) are real material changes or false positives (e.g. whitespace, punctuation, page numbering artefacts).",
            rules = new[]
            {
                "Read the actual page content (contentAfs1 / contentAfs2) to verify and enrich the diff findings.",
                "Do not invent facts or numbers not present in the input.",
                "pairIndex is zero-based. focusDiffOrdinals are 1-based ordinals within issueDiffs for that pair.",
                "Use concise audit language. Prioritise numeric changes, policy rewrites, and removed/added disclosures.",
                "For each material finding, explain the business impact (e.g. 'Revenue decreased by R2.4M, which exceeds 5% materiality threshold').",
                "Severity 'High' = numeric change >5% or policy/disclosure change. 'Medium' = 1-5% or minor narrative. 'Low' = immaterial.",
                "For nearIdenticalPairs: if the diff is whitespace, punctuation, or a page-number artefact, mark severity 'Low' and note it is likely a false positive in auditRationale.",
                "For nearIdenticalPairs where pages appear truly identical in content, explicitly say 'Likely false positive — no material difference' in the finding title.",
            },
            report1 = new
            {
                comparison.Report1.Label,
                comparison.Report1.Filename,
                pages = comparison.Report1.PageCount,
                words = comparison.Report1.WordCount,
                year = comparison.Report1.PrimaryYear,
            },
            report2 = new
            {
                comparison.Report2.Label,
                comparison.Report2.Filename,
                pages = comparison.Report2.PageCount,
                words = comparison.Report2.WordCount,
                year = comparison.Report2.PrimaryYear,
            },
            counts = new
            {
                same    = comparison.Counts.GetValueOrDefault("same", 0),
                changed = comparison.Counts.GetValueOrDefault("changed", 0),
                added   = comparison.Counts.GetValueOrDefault("added", 0),
                removed = comparison.Counts.GetValueOrDefault("removed", 0),
            },
            numberComparison = new
            {
                similarityPct = comparison.NumCmp.SimilarityPct,
                onlyInAfs1    = comparison.NumCmp.OnlyInAfs1.Take(30).ToList(),
                onlyInAfs2    = comparison.NumCmp.OnlyInAfs2.Take(30).ToList(),
            },
            smartComparison = comparison.SmartComparison == null ? null : new
            {
                totalChanges       = comparison.SmartComparison.TotalChanges,
                materialCount      = comparison.SmartComparison.MaterialCount,
                numericChanges     = comparison.SmartComparison.MaterialNumeric.Count,
                narrativeRewrites  = comparison.SmartComparison.NarrativeRewrites.Count,
            },
            documentIntelligence = comparison.DocumentIntelligence == null ? null : new
            {
                summary         = Trim(comparison.DocumentIntelligence.Summary, 800),
                materialSections = comparison.DocumentIntelligence.Sections.Count(s => s.IsMaterial),
            },
            // Actual page text for the most-changed pages — Gemini reads these directly
            pageContentSamples,
            // Near-identical pages with at least 1 detected diff — validate if real or false positive
            nearIdenticalPairs,
            pagePairs = comparison.PageDiffs
                .Where(p => p.IssueCount > 0)
                .OrderByDescending(p => p.IssueCount)
                .ThenByDescending(p => p.Changed)
                .Take(20)
                .Select(p => new
                {
                    pairIndex = p.PairIndex,
                    p.PageAfs1,
                    p.PageAfs2,
                    p.SecondaryMatch,
                    p.Unmatched,
                    p.AlignSim,
                    p.PctSame,
                    p.Same,
                    p.Changed,
                    p.Added,
                    p.Removed,
                    issueDiffs = p.Diffs
                        .Where(d => d.Status != "same" &&
                                    d.Status != "same-cross-page" &&
                                    d.Status != "same-different-page")
                        .Take(10)
                        .Select((d, idx) => new
                        {
                            ordinal        = idx + 1,
                            status         = d.Status,
                            line1          = Trim(d.Line1, 280),
                            line2          = Trim(d.Line2, 280),
                            changedNumbers = d.NumDiff.Take(10).ToList(),
                            crossPageNote  = Trim(d.CrossPageNote, 140),
                        })
                        .ToList()
                })
                .ToList()
        };

        return
            "You are an expert financial audit AI assistant inside a PDF comparison tool for SNG Grant Thornton. " +
            "Read BOTH the actual page content samples AND the deterministic diff output provided below. " +
            "Return only JSON matching the requested schema. " +
            "Use the page content to verify numeric changes, identify accounting policy shifts, " +
            "and flag disclosures that were added or removed between the two AFS versions.\n\n" +
            JsonSerializer.Serialize(context, JsonOptions);
    }

    private static object BuildResponseSchema()
    {
        return new
        {
            type = "object",
            properties = new Dictionary<string, object>
            {
                ["executiveSummary"] = new { type = "string" },
                ["overallRisk"] = new { type = "string" },
                ["materialFindings"] = new
                {
                    type = "array",
                    items = new
                    {
                        type = "object",
                        properties = new Dictionary<string, object>
                        {
                            ["title"] = new { type = "string" },
                            ["severity"] = new { type = "string" },
                            ["pairIndex"] = new { type = new[] { "integer", "null" } },
                            ["pageAfs1"] = new { type = new[] { "integer", "null" } },
                            ["pageAfs2"] = new { type = new[] { "integer", "null" } },
                            ["focusDiffOrdinals"] = new
                            {
                                type = "array",
                                items = new { type = "integer" }
                            },
                            ["auditRationale"] = new { type = "string" },
                            ["recommendedProcedure"] = new { type = "string" },
                            ["changedNumbers"] = new
                            {
                                type = "array",
                                items = new { type = "string" }
                            }
                        },
                        required = new[]
                        {
                            "title", "severity", "pairIndex", "pageAfs1", "pageAfs2",
                            "focusDiffOrdinals", "auditRationale", "recommendedProcedure", "changedNumbers"
                        }
                    }
                },
                ["pageFocus"] = new
                {
                    type = "array",
                    items = new
                    {
                        type = "object",
                        properties = new Dictionary<string, object>
                        {
                            ["pairIndex"] = new { type = "integer" },
                            ["pageAfs1"] = new { type = new[] { "integer", "null" } },
                            ["pageAfs2"] = new { type = new[] { "integer", "null" } },
                            ["priority"] = new { type = "string" },
                            ["why"] = new { type = "string" },
                            ["focusDiffOrdinals"] = new
                            {
                                type = "array",
                                items = new { type = "integer" }
                            }
                        },
                        required = new[] { "pairIndex", "pageAfs1", "pageAfs2", "priority", "why", "focusDiffOrdinals" }
                    }
                },
                ["highlightGuidance"] = new
                {
                    type = "array",
                    items = new { type = "string" }
                },
                ["limitations"] = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            required = new[] { "executiveSummary", "overallRisk", "materialFindings", "pageFocus", "highlightGuidance", "limitations" }
        };
    }

    private static void Normalize(AiAssistResult result)
    {
        result.ExecutiveSummary = Trim(result.ExecutiveSummary, 1200);
        result.Message = Trim(result.Message, 240);
        result.OverallRisk = Trim(result.OverallRisk, 60);
        result.HighlightGuidance = result.HighlightGuidance.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => Trim(s, 220)).Distinct().Take(8).ToList();
        result.Limitations = result.Limitations.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => Trim(s, 220)).Distinct().Take(8).ToList();

        foreach (var finding in result.MaterialFindings)
        {
            finding.Title = Trim(finding.Title, 180);
            finding.Severity = Trim(finding.Severity, 20);
            finding.AuditRationale = Trim(finding.AuditRationale, 360);
            finding.RecommendedProcedure = Trim(finding.RecommendedProcedure, 280);
            finding.FocusDiffOrdinals = finding.FocusDiffOrdinals.Where(i => i > 0).Distinct().Take(8).ToList();
            finding.ChangedNumbers = finding.ChangedNumbers.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => Trim(s, 80)).Distinct().Take(12).ToList();
        }

        foreach (var focus in result.PageFocus)
        {
            focus.Priority = Trim(focus.Priority, 20);
            focus.Why = Trim(focus.Why, 220);
            focus.FocusDiffOrdinals = focus.FocusDiffOrdinals.Where(i => i > 0).Distinct().Take(8).ToList();
        }

        result.MaterialFindings = result.MaterialFindings.Take(12).ToList();
        result.PageFocus = result.PageFocus
            .Where(p => p.PairIndex >= 0)
            .OrderBy(p => PriorityRank(p.Priority))
            .ThenBy(p => p.PairIndex)
            .Take(24)
            .ToList();
    }

    private static int PriorityRank(string? priority)
    {
        return (priority ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "high" => 0,
            "medium" => 1,
            "low" => 2,
            _ => 3,
        };
    }

    private bool GetBool(string key, bool defaultValue)
    {
        string? raw = _configuration[key];
        return bool.TryParse(raw, out bool parsed) ? parsed : defaultValue;
    }

    private static string ExtractModelText(string rawResponse)
    {
        using var doc = JsonDocument.Parse(rawResponse);
        if (!doc.RootElement.TryGetProperty("candidates", out JsonElement candidates) ||
            candidates.ValueKind != JsonValueKind.Array ||
            candidates.GetArrayLength() == 0)
            return string.Empty;

        JsonElement first = candidates[0];
        if (!first.TryGetProperty("content", out JsonElement content) ||
            !content.TryGetProperty("parts", out JsonElement parts) ||
            parts.ValueKind != JsonValueKind.Array)
            return string.Empty;

        foreach (JsonElement part in parts.EnumerateArray())
        {
            if (part.TryGetProperty("text", out JsonElement textNode))
                return textNode.GetString() ?? string.Empty;
        }

        return string.Empty;
    }

    private static string StripCodeFences(string text)
    {
        string trimmed = text.Trim();
        if (!trimmed.StartsWith("```", StringComparison.Ordinal))
            return trimmed;

        int firstLine = trimmed.IndexOf('\n');
        int lastFence = trimmed.LastIndexOf("```", StringComparison.Ordinal);
        if (firstLine < 0 || lastFence <= firstLine)
            return trimmed;

        return trimmed[(firstLine + 1)..lastFence].Trim();
    }

    private static string Trim(string? value, int maxLength)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        string trimmed = value.Trim();
        return trimmed.Length <= maxLength ? trimmed : trimmed[..maxLength] + "...";
    }

    private static AiAssistResult Unavailable(string provider, string model, string status, bool enabled = true) => new()
    {
        Enabled = enabled,
        Provider = provider,
        Model = model,
        Status = status,
        Message = AiUserFacingMessages.ReviewUnavailable,
        Limitations = AiUserFacingMessages.ReviewUnavailableLimitations(),
        GeneratedAt = DateTime.UtcNow,
    };
}
