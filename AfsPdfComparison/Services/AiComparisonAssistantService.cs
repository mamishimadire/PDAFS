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

    public async Task<AiAssistResult> AnalyzeAsync(ComparisonResult comparison, CancellationToken cancellationToken = default, bool lean = false)
    {
        bool enabled = AiConfiguration.GetBool(_configuration, "AiAssist:Enabled", true);
        string provider = _configuration["AiAssist:Provider"] ?? "Gemini";
        string model = AiConfiguration.GetString(
            _configuration,
            "gemini-2.5-flash",
            "AiAssist:GeminiModel",
            "GEMINI_MODEL");

        if (!enabled)
        {
            return Unavailable(provider, model, "disabled", enabled: false);
        }

        string apiKey = AiConfiguration.GetOptional(
            _configuration,
            "AiAssist:GeminiApiKey",
            "GEMINI_API_KEY") ?? string.Empty;

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
            string prompt = BuildPrompt(comparison, lean: lean);
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

            // Post-processing: remove false-positive findings where the AI
            // incorrectly flagged layout artefacts as content removals.
            aiResult.MaterialFindings = SuppressFalsePositiveFindings(
                aiResult.MaterialFindings, comparison);

            Normalize(aiResult);
            return aiResult;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            _logger.LogDebug("Gemini AI review cancelled because another provider completed first.");
            return Unavailable(provider, model, "cancelled");
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

    private string BuildPrompt(ComparisonResult comparison, bool lean = false)
    {
        // Full document text. Gemini 2.5 Flash supports ~1M tokens but we cap for reliability.
        int textCap = lean ? 60_000 : 300_000;
        string fullText1 = Trim(comparison.Report1.FullText ?? "", textCap);
        string fullText2 = Trim(comparison.Report2.FullText ?? "", textCap);

        // Page content for the most-changed page pairs (capped to avoid oversized prompts)
        int pairCap = lean ? 8 : 20;
        var topChangedPairs = comparison.PageDiffs
            .Where(p => p.IssueCount > 0)
            .OrderByDescending(p => p.IssueCount)
            .Take(pairCap)
            .ToList();

        var pageContentSamples = topChangedPairs.Select(p => new
        {
            pairIndex = p.PairIndex,
            pageAfs1 = p.PageAfs1,
            pageAfs2 = p.PageAfs2,
            contentAfs1 = p.PageAfs1.HasValue && p.PageAfs1.Value - 1 < comparison.Report1.Pages.Count
                ? Trim(comparison.Report1.Pages[p.PageAfs1.Value - 1], 3000)
                : null,
            contentAfs2 = p.PageAfs2.HasValue && p.PageAfs2.Value - 1 < comparison.Report2.Pages.Count
                ? Trim(comparison.Report2.Pages[p.PageAfs2.Value - 1], 3000)
                : null,
        }).ToList();

        // Near-identical page pairs (capped)
        var nearIdenticalPairs = comparison.PageDiffs
            .Where(p => p.IssueCount > 0 && p.PctSame >= 95)
            .Take(lean ? 5 : 20)
            .Select(p => new
            {
                pairIndex = p.PairIndex,
                p.PageAfs1,
                p.PageAfs2,
                p.PctSame,
                p.IssueCount,
                issueDiffs = p.Diffs
                    .Where(d => d.Status != "same" && d.Status != "same-cross-page" && d.Status != "same-different-page")
                    .Select((d, i) => new { ordinal = i + 1, status = d.Status, line1 = d.Line1, line2 = d.Line2 })
                    .ToList()
            })
            .ToList();

        // ── Compact diff report — covers ALL changed/added/removed lines ──────────
        // Compact text format (~100 chars/line) is far more token-efficient than JSON
        // objects with nested wordDiff arrays (~400 chars/line), allowing the AI to
        // see every single difference the system detected, not just a capped subset.
        int lineCap = lean ? 50 : int.MaxValue;   // lean: top 50; normal: ALL
        string compactDiffReport = BuildCompactDiffReport(comparison, lineCap);

        // Page Alignment Summary
        var pageAlignmentSummary = comparison.Alignment.Select((a, i) => new
        {
            afs1Page  = a.I1 + 1,
            afs2Page  = a.I2 >= 0 ? (int?)(a.I2 + 1) : null,
            sim       = Math.Round(a.Sim, 3),
            matched   = a.I2 >= 0,
        }).ToList();

        var context = new
        {
            task = "You are an expert audit-review AI at SNG Grant Thornton. " +
                   "You are given: (1) the full text of BOTH documents, (2) a compact line-diff (compactDiffReport) covering EVERY change, " +
                   "and (3) systemReviewFindings — deterministic findings already categorised by the system engine. " +
                   "Your PRIMARY job is to start from systemReviewFindings and produce a complete, professionally written audit review: " +
                   "enrich each system finding with an auditRationale (quoting exact text from both documents) and a recommendedProcedure. " +
                   "Then scan compactDiffReport for any additional genuine material changes the system findings may have missed. " +
                   "CRITICAL: the diff engine is mechanical — verify every finding against the actual document text. " +
                   "A 'removed' line is only a real removal if the content does NOT exist anywhere in AFS 2. " +
                   "Use the pageAfs1 and pageAfs2 fields in systemReviewFindings for correct snapshot references.",

            accuracy_rules = new[]
            {
                "RULE 1 — USE SYSTEM DIFF: The systemChangedLines list shows exact word-level changes (wordDiff: tag='removed' = red in UI, tag='added' = blue in UI). Reference these by ordinal number in your findings (e.g. 'Changed Line #3 shows ...').",
                "RULE 2 — VERIFY BEFORE REPORTING: Before flagging ANY finding, search fullDocumentText2 to confirm the 'removed' content is genuinely absent from AFS 2. If the content exists anywhere in AFS 2, do NOT report it as removed — it is a page layout artefact.",
                "RULE 3 — NO FABRICATION: Only report facts you can directly quote from the system diff data, contentAfs1, contentAfs2, fullDocumentText1, or fullDocumentText2. Never invent clauses, numbers, dates, or parties.",
                "RULE 4 — QUOTE EXACTLY: For every finding, your auditRationale MUST include the exact text or value from each document. Format: 'AFS 1 states: \"...exact quote...\" — AFS 2 states: \"...exact quote...\"'. If you cannot find the exact quote, do not make the finding.",
                "RULE 5 — LAYOUT VS CONTENT: The same clause on different pages is NOT a change. Different line wrapping is NOT a change. Only flag changes where the actual TEXT or NUMBERS differ between AFS 1 and AFS 2.",
                "RULE 6 — NUMERIC ACCURACY: For number changes, cite both the AFS 1 value and AFS 2 value exactly as they appear in the documents (e.g. 'Revenue: R45 234 000 in AFS 1 → R48 612 000 in AFS 2'). The numDiff field in systemChangedLines shows which numbers changed per line.",
                "RULE 7 — ADDED LINES: systemAddedLines are lines present in AFS 2 but not AFS 1. These may represent new disclosures, additional notes, or genuinely new content. Verify against fullDocumentText1 that the content truly did not exist before.",
                "RULE 8 — REMOVED LINES: systemRemovedLines are lines present in AFS 1 but not AFS 2. Verify against fullDocumentText2 that the content was truly removed (not just moved to a different page).",
                "RULE 9 — PAGE ALIGNMENT: Use pageAlignmentSummary to understand which AFS 1 pages map to which AFS 2 pages. Unmatched pages (afs2Page = null) are pages that exist only in one version.",
                "RULE 10 — NEAR-IDENTICAL PAGES: For pages with pctSame ≥ 95%, the diff is almost certainly a formatting artefact. Only flag if you can see a clear numeric or textual difference in the actual page content.",
                "RULE 11 — SPECIFIC REFERENCES: Always cite the specific clause number, section heading, page number, or table name where the change appears (e.g. 'Clause 1.2.19', 'Statement of Financial Position', 'Note 12 — Provisions').",
                "RULE 12 — SEVERITY: High = confirmed numeric change >5% or policy/disclosure genuinely added/removed. Medium = confirmed 1–5% numeric or minor narrative change. Low = immaterial or formatting-only.",
                "RULE 13 — COMPLETENESS: Start from systemReviewFindings — convert EVERY system finding into a materialFinding with proper auditRationale and recommendedProcedure. Then work through ALL remaining items in compactDiffReport for any additional genuine changes the system may have missed. Do NOT stop early. There is no cap — a document with 40 system findings and 10 additional AI-detected changes should produce 50 findings.",
            },

            fullDocumentText1 = fullText1,
            fullDocumentText2 = fullText2,
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
                onlyInAfs1    = comparison.NumCmp.OnlyInAfs1,
                onlyInAfs2    = comparison.NumCmp.OnlyInAfs2,
            },
            // Complete diff report — every changed/added/removed line the system detected.
            // Compact text format: CHG-001 [p3→p3]: "AFS1 text" → "AFS2 text" | nums/words
            // ADD-001 [AFS2-p8]: "new line"    REM-001 [AFS1-p5]: "removed line"
            compactDiffReport,
            // Deterministic system review — categorised findings the engine already produced.
            // Use these as a confirmed starting point. Enrich each with professional English,
            // verify the pageAfs1/pageAfs2 references, add an auditRationale and recommendedProcedure.
            systemReviewFindings = comparison.AutoReview.Select(f => new
            {
                f.Id,
                f.Category,
                f.Title,
                f.Description,
                f.Severity,
                f.PageAfs1,
                f.PageAfs2,
                f.PairIndex,
                f.Before,
                f.After,
                pctChange = f.PctChange.HasValue ? Math.Round(f.PctChange.Value, 1) : (double?)null,
                absChange = f.AbsChange,
            }).ToList(),
            pageAlignmentSummary,
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
                        .Select((d, idx) => new
                        {
                            ordinal        = idx + 1,
                            status         = d.Status,
                            line1          = d.Line1,
                            line2          = d.Line2,
                            changedNumbers = d.NumDiff,
                            crossPageNote  = d.CrossPageNote,
                        })
                        .ToList()
                })
                .ToList()
        };

        return
            "You are an expert financial audit AI assistant inside a PDF comparison tool for SNG Grant Thornton. " +
            "Read BOTH the actual page content samples AND the deterministic diff output provided below. " +
            "The systemChangedLines, systemAddedLines, and systemRemovedLines fields contain the EXACT same data " +
            "shown to the auditor in the Changed Lines, Added Lines, and Removed Lines tabs of the system. " +
            "Use these as your primary evidence source — cross-reference them with the full document text to verify findings. " +
            "Return only JSON matching the requested schema.\n\n" +
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

    // Remove findings that claim content was "removed" when it still exists in AFS2,
    // or that claim content was "added" when it already existed in AFS1.
    // These are caused by the diff engine seeing cross-page layout moves as removals.
    private List<AiMaterialFinding> SuppressFalsePositiveFindings(
        List<AiMaterialFinding> findings,
        ComparisonResult comparison)
    {
        string full1 = (comparison.Report1.FullText ?? "").ToLowerInvariant();
        string full2 = (comparison.Report2.FullText ?? "").ToLowerInvariant();

        var kept = new List<AiMaterialFinding>();
        foreach (var f in findings)
        {
            // Drop findings the AI itself labelled as false positives
            if ((f.Title ?? "").Contains("false positive", StringComparison.OrdinalIgnoreCase) ||
                (f.Title ?? "").Contains("no material difference", StringComparison.OrdinalIgnoreCase))
                continue;

            // Extract meaningful key terms from the title for corpus search
            var titleWords = (f.Title ?? "")
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Where(w => w.Length >= 5)
                .Select(w => System.Text.RegularExpressions.Regex.Replace(w, @"[^\w]", "").ToLowerInvariant())
                .Where(w => w.Length >= 4 && w != "removal" && w != "removed" && w != "added" && w != "change")
                .ToList();

            if (titleWords.Count == 0) { kept.Add(f); continue; }

            bool claimsRemoval = (f.Title ?? "").Contains("remov", StringComparison.OrdinalIgnoreCase) ||
                                  (f.AuditRationale ?? "").StartsWith("removed from", StringComparison.OrdinalIgnoreCase);

            bool claimsAddition = (f.Title ?? "").Contains("new disclosure", StringComparison.OrdinalIgnoreCase) &&
                                   !(f.Title ?? "").Contains("updated", StringComparison.OrdinalIgnoreCase);

            if (claimsRemoval && full2.Length > 100)
            {
                // If all key subject words are found in AFS2, the content wasn't actually removed
                int foundInAfs2 = titleWords.Count(w => full2.Contains(w, StringComparison.Ordinal));
                if (titleWords.Count >= 2 && foundInAfs2 >= titleWords.Count)
                {
                    _logger.LogInformation(
                        "Suppressed false-positive removal finding '{Title}' — all key terms present in AFS2.", f.Title);
                    continue;
                }
            }

            if (claimsAddition && full1.Length > 100)
            {
                int foundInAfs1 = titleWords.Count(w => full1.Contains(w, StringComparison.Ordinal));
                if (titleWords.Count >= 2 && foundInAfs1 >= titleWords.Count)
                {
                    _logger.LogInformation(
                        "Suppressed false-positive addition finding '{Title}' — all key terms present in AFS1.", f.Title);
                    continue;
                }
            }

            kept.Add(f);
        }
        return kept;
    }

    private static void Normalize(AiAssistResult result)
    {
        result.ExecutiveSummary = Trim(result.ExecutiveSummary, 1200);
        result.Message = Trim(result.Message, 240);
        result.OverallRisk = Trim(result.OverallRisk, 60);
        result.HighlightGuidance = result.HighlightGuidance.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        result.Limitations = result.Limitations.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();

        foreach (var finding in result.MaterialFindings)
        {
            finding.Title = Trim(finding.Title, 300);
            finding.Severity = Trim(finding.Severity, 20);
            finding.AuditRationale = Trim(finding.AuditRationale, 2000);
            finding.RecommendedProcedure = Trim(finding.RecommendedProcedure, 800);
            finding.FocusDiffOrdinals = finding.FocusDiffOrdinals.Where(i => i > 0).Distinct().ToList();
            finding.ChangedNumbers = finding.ChangedNumbers.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
        }

        foreach (var focus in result.PageFocus)
        {
            focus.Priority = Trim(focus.Priority, 20);
            focus.Why = Trim(focus.Why, 600);
            focus.FocusDiffOrdinals = focus.FocusDiffOrdinals.Where(i => i > 0).Distinct().ToList();
        }

        // No caps on findings or page focus — return everything the AI identified
        result.PageFocus = result.PageFocus
            .Where(p => p.PairIndex >= 0)
            .OrderBy(p => PriorityRank(p.Priority))
            .ThenBy(p => p.PairIndex)
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

    /// <summary>
    /// Builds a compact plain-text diff report covering ALL (or top N) changed/added/removed lines.
    /// Format: ~100 chars/entry vs ~400 chars/entry for JSON — 4x more token-efficient.
    /// The AI receives every single difference detected by the engine.
    /// </summary>
    private static string BuildCompactDiffReport(ComparisonResult cmp, int lineCap)
    {
        var sb = new System.Text.StringBuilder(8192);

        // Build page context lookup: LineDiffResult → PageDiffResult
        var lineToPage = new Dictionary<LineDiffResult, (int? p1, int? p2)>(
            ReferenceEqualityComparer.Instance);
        foreach (var pd in cmp.PageDiffs)
            foreach (var d in pd.Diffs)
                lineToPage.TryAdd(d, (pd.PageAfs1, pd.PageAfs2));

        // ── Changed lines ─────────────────────────────────────────────────────
        var changed = cmp.ChangedLines.Take(lineCap).ToList();
        sb.AppendLine($"=== CHANGED LINES ({cmp.ChangedLines.Count()} total" +
                      (changed.Count < cmp.ChangedLines.Count() ? $", showing {changed.Count}" : "") + ") ===");
        int ord = 0;
        foreach (var d in changed)
        {
            ord++;
            lineToPage.TryGetValue(d, out var pg);
            string pageRef = pg.p1.HasValue || pg.p2.HasValue
                ? $"p{pg.p1}→p{pg.p2}" : "?";
            string l1 = Trim(d.Line1, 120);
            string l2 = Trim(d.Line2, 120);
            string suffix = "";
            if (d.NumDiff.Count > 0)
                suffix += $" | nums:[{string.Join(",", d.NumDiff.Take(6))}]";
            if (d.WordDiff.Count > 0)
            {
                var rem = d.WordDiff.Where(w => w.Tag == "removed").Select(w => w.Word).Take(5).ToList();
                var add = d.WordDiff.Where(w => w.Tag == "added").Select(w => w.Word).Take(5).ToList();
                if (rem.Count > 0) suffix += $" | -[{string.Join(",", rem)}]";
                if (add.Count > 0) suffix += $" | +[{string.Join(",", add)}]";
            }
            sb.AppendLine($"CHG-{ord:D3} [{pageRef}]: \"{l1}\" → \"{l2}\"{suffix}");
        }

        // ── Added lines ───────────────────────────────────────────────────────
        var added = cmp.AddedLines.Take(lineCap).ToList();
        sb.AppendLine();
        sb.AppendLine($"=== LINES ADDED IN AFS 2 ({cmp.AddedLines.Count()} total" +
                      (added.Count < cmp.AddedLines.Count() ? $", showing {added.Count}" : "") + ") ===");
        ord = 0;
        foreach (var d in added)
        {
            ord++;
            lineToPage.TryGetValue(d, out var pg);
            string pageRef = pg.p2.HasValue ? $"AFS2-p{pg.p2}" : "AFS2-?";
            string line = Trim(d.Line2, 150);
            sb.AppendLine($"ADD-{ord:D3} [{pageRef}]: \"{line}\"");
        }

        // ── Removed lines ─────────────────────────────────────────────────────
        var removed = cmp.RemovedLines.Take(lineCap).ToList();
        sb.AppendLine();
        sb.AppendLine($"=== LINES REMOVED FROM AFS 1 ({cmp.RemovedLines.Count()} total" +
                      (removed.Count < cmp.RemovedLines.Count() ? $", showing {removed.Count}" : "") + ") ===");
        ord = 0;
        foreach (var d in removed)
        {
            ord++;
            lineToPage.TryGetValue(d, out var pg);
            string pageRef = pg.p1.HasValue ? $"AFS1-p{pg.p1}" : "AFS1-?";
            string line = Trim(d.Line1, 150);
            sb.AppendLine($"REM-{ord:D3} [{pageRef}]: \"{line}\"");
        }

        return sb.ToString();
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
