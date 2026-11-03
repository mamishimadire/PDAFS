using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using AfsPdfComparison.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace AfsPdfComparison.Services;

/// <summary>
/// Sends the raw extracted text of both AFS documents directly to Gemini and
/// asks it to read, understand, and compare them as a financial auditor would.
/// This runs IN ADDITION to the deterministic comparison — the AI adds semantic
/// understanding, section-level analysis, and narrative interpretation.
/// </summary>
public class AiContentReaderService
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy        = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition      = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented               = false,
    };

    private readonly IHttpClientFactory _http;
    private readonly OpenAiCompatibleAiService _openAiAdapter;
    private readonly IConfiguration     _config;
    private readonly ILogger<AiContentReaderService> _logger;

    public AiContentReaderService(
        IHttpClientFactory http,
        OpenAiCompatibleAiService openAiAdapter,
        IConfiguration     config,
        ILogger<AiContentReaderService> logger)
    {
        _http          = http;
        _openAiAdapter = openAiAdapter;
        _config        = config;
        _logger        = logger;
    }

    public async Task<AiContentReadResult> ReadAndCompareAsync(
        PdfReport report1,
        PdfReport report2,
        CancellationToken cancellationToken = default)
    {
        bool enabled = AiConfiguration.GetBool(_config, "AiAssist:Enabled", true);

        if (!enabled)
            return Unavailable("disabled", enabled: false);

        string prompt = BuildContentReadPrompt(report1, report2);

        var attempts = new[]
        {
            new AiProviderAttempt<AiContentReadResult>("Gemini", ct => TryGeminiAsync(prompt, ct)),
            new AiProviderAttempt<AiContentReadResult>("OpenRouter", ct => TryOpenAiCompatibleAsync(
                provider: "OpenRouter",
                endpoint: AiConfiguration.GetString(
                    _config,
                    "https://openrouter.ai/api/v1/chat/completions",
                    "AiAssist:OpenRouterUrl",
                    "OPENROUTER_URL"),
                apiKey: AiConfiguration.GetOptional(_config, "AiAssist:OpenRouterApiKey", "OPENROUTER_API_KEY"),
                model: AiConfiguration.GetString(
                    _config,
                    "deepseek/deepseek-chat",
                    "AiAssist:OpenRouterModel",
                    "OPENROUTER_MODEL"),
                prompt: prompt,
                timeoutSeconds: 120,
                cancellationToken: ct)),
            new AiProviderAttempt<AiContentReadResult>("Groq", ct => TryOpenAiCompatibleAsync(
                provider: "Groq",
                endpoint: AiConfiguration.GetString(
                    _config,
                    "https://api.groq.com/openai/v1/chat/completions",
                    "AiAssist:GroqUrl",
                    "GROQ_URL"),
                apiKey: AiConfiguration.GetOptional(_config, "AiAssist:GroqApiKey", "GROQ_API_KEY"),
                model: AiConfiguration.GetString(
                    _config,
                    "llama-3.3-70b-versatile",
                    "AiAssist:GroqModel",
                    "GROQ_MODEL"),
                prompt: prompt,
                timeoutSeconds: 90,
                cancellationToken: ct)),
        };

        AiFailoverOutcome<AiContentReadResult> outcome = await AiFailoverExecutor.ExecuteAsync(
            attempts,
            result => result.Status,
            readyResultTarget: 1,
            waitAfterFirstReady: TimeSpan.Zero,
            cancellationToken);

        if (outcome.ReadyResults.Count > 0)
            return outcome.ReadyResults[0];

        bool configurationOnly = outcome.CompletedResults
            .All(r => string.Equals(r.Status, "disabled", StringComparison.OrdinalIgnoreCase) ||
                      string.Equals(r.Status, "missing_key", StringComparison.OrdinalIgnoreCase));

        return configurationOnly
            ? Unavailable("missing_key", enabled: false)
            : Unavailable("error");
    }

    // ─────────────────────────────────────────────────────────────────────────
    private async Task<AiContentReadResult> TryGeminiAsync(
        string prompt,
        CancellationToken cancellationToken)
    {
        string model = AiConfiguration.GetString(
            _config,
            "gemini-2.5-flash",
            "AiAssist:GeminiModel",
            "GEMINI_MODEL");
        string apiKey = AiConfiguration.GetOptional(
            _config,
            "AiAssist:GeminiApiKey",
            "GEMINI_API_KEY") ?? string.Empty;

        if (string.IsNullOrWhiteSpace(apiKey))
            return Unavailable("missing_key", enabled: false);

        try
        {
            string modelPath = model.StartsWith("models/", StringComparison.OrdinalIgnoreCase)
                ? model : "models/" + model;

            string endpoint = $"https://generativelanguage.googleapis.com/v1beta/{modelPath}:generateContent";
            object requestBody = new
            {
                contents = new[]
                {
                    new { parts = new[] { new { text = prompt } } }
                },
                generationConfig = new
                {
                    temperature       = 0.15,
                    responseMimeType  = "application/json",
                    responseJsonSchema = BuildSchema(),
                }
            };

            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(TimeSpan.FromSeconds(120));

            using var req = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(requestBody, JsonOpts),
                    Encoding.UTF8, "application/json")
            };
            req.Headers.Add("x-goog-api-key", apiKey);

            var client = _http.CreateClient();
            using var resp = await client.SendAsync(req, cts.Token);
            string raw = await resp.Content.ReadAsStringAsync(cts.Token);

            if (!resp.IsSuccessStatusCode)
            {
                _logger.LogWarning("AiContentReader: Gemini returned HTTP {Code}: {Body}", (int)resp.StatusCode, Trim(raw, 400));
                return Unavailable("error");
            }

            string jsonText = ExtractText(raw);
            if (string.IsNullOrWhiteSpace(jsonText))
            {
                _logger.LogWarning("AiContentReader: Gemini returned an empty content-read response.");
                return Unavailable("error");
            }

            return DeserializeContentReadResult(StripFences(jsonText), provider: "Gemini", model);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            _logger.LogDebug("AiContentReader: Gemini request cancelled because another provider completed first.");
            return Unavailable("cancelled");
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("AiContentReader: Gemini request timed out after {TimeoutSeconds}s.", 120);
            return Unavailable("error");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "AiContentReader: Gemini unexpected failure.");
            return Unavailable("error");
        }
    }

    private async Task<AiContentReadResult> TryOpenAiCompatibleAsync(
        string provider,
        string endpoint,
        string? apiKey,
        string model,
        string prompt,
        int timeoutSeconds,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
            return Unavailable("missing_key", enabled: false);

        const string systemPrompt =
            "You are an expert financial auditor at SNG Grant Thornton. " +
            "Read and compare both AFS documents carefully. " +
            "Return only valid JSON that matches the requested structure.";

        var response = await _openAiAdapter.RequestJsonObjectAsync(
            endpoint,
            apiKey,
            model,
            provider,
            systemPrompt,
            prompt,
            timeoutSeconds,
            cancellationToken);

        if (!response.Success)
            return string.Equals(response.Status, "disabled", StringComparison.OrdinalIgnoreCase)
                ? Unavailable("missing_key", enabled: false)
                : Unavailable("error");

        return DeserializeContentReadResult(response.JsonText, provider, model);
    }

    private AiContentReadResult DeserializeContentReadResult(
        string jsonText,
        string provider,
        string model)
    {
        var result = JsonSerializer.Deserialize<AiContentReadResult>(
            jsonText,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new AiContentReadResult();

        result.Enabled     = true;
        result.Status      = "ready";
        result.Message     = $"{provider} content reading completed.";
        result.GeneratedAt = DateTime.UtcNow;
        Normalize(result);
        _logger.LogInformation("AiContentReader: content reading completed via {Provider} ({Model}).", provider, model);
        return result;
    }

    // ─────────────────────────────────────────────────────────────────────────
    private string BuildContentReadPrompt(PdfReport report1, PdfReport report2)
    {
        string text1 = BuildFullDocumentText(report1);
        string text2 = BuildFullDocumentText(report2);

        return
            "You are an expert financial auditor at SNG Grant Thornton. " +
            "You have extracted the text of two versions of an Annual Financial Statement (AFS). " +
            "The document text is prefixed with PAGE MARKERS in the format '=== PAGE N ===' so you know which page each section appears on. " +
            "Read both documents carefully and perform a comprehensive comparison.\n\n" +
            "Your job:\n" +
            "1. Identify every section (Balance Sheet, Income Statement, Cash Flow, Notes, Policies, Directors' Report, Auditors' Report, etc.)\n" +
            "2. For each section, summarise what changed between AFS 1 and AFS 2\n" +
            "3. For each section diff, provide pageAfs1 (the page number in AFS 1 where this section appears) and pageAfs2 (the page number in AFS 2)\n" +
            "4. List all numeric line item changes you can find (e.g. 'Revenue: R45.2M → R48.6M') with the page numbers they appear on\n" +
            "5. Flag accounting policy changes, new disclosures, and removed disclosures\n" +
            "6. Provide an overall materiality assessment and audit conclusion\n\n" +
            "IMPORTANT: Always set pageAfs1 and pageAfs2 on every sectionDiff and lineItemChange. Use the === PAGE N === markers to determine the correct page numbers.\n\n" +
            "Return ONLY valid JSON matching the provided schema. Do not include markdown or code fences.\n\n" +
            "=== AFS 1 — " + report1.Label + " (" + report1.Filename + ") ===\n" +
            text1 +
            "\n\n=== AFS 2 — " + report2.Label + " (" + report2.Filename + ") ===\n" +
            text2;
    }

    // Build full document text with page markers — no character limit.
    // Gemini 2.5 Flash supports ~1M tokens; a typical AFS is well within that.
    private static string BuildFullDocumentText(PdfReport report)
    {
        if (report.Pages == null || report.Pages.Count == 0)
            return string.IsNullOrWhiteSpace(report.FullText)
                ? "(No text extracted — document may be scanned or empty)"
                : report.FullText;

        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < report.Pages.Count; i++)
        {
            string pageText = report.Pages[i]?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(pageText)) continue;
            sb.Append($"\n=== PAGE {i + 1} ===\n");
            sb.Append(pageText);
        }
        return sb.ToString();
    }

    private static object BuildSchema() => new
    {
        type = "object",
        properties = new Dictionary<string, object>
        {
            ["overallAssessment"]  = new { type = "string" },
            ["riskLevel"]          = new { type = "string" },
            ["auditConclusion"]    = new { type = "string" },
            ["sectionDiffs"] = new
            {
                type  = "array",
                items = new
                {
                    type = "object",
                    properties = new Dictionary<string, object>
                    {
                        ["sectionName"]       = new { type = "string" },
                        ["changeType"]        = new { type = "string" },
                        ["severity"]          = new { type = "string" },
                        ["afs1Summary"]       = new { type = "string" },
                        ["afs2Summary"]       = new { type = "string" },
                        ["changeDescription"] = new { type = "string" },
                        ["changedValues"]     = new { type = "array", items = new { type = "string" } },
                        ["pageAfs1"]          = new { type = new[] { "integer", "null" } },
                        ["pageAfs2"]          = new { type = new[] { "integer", "null" } },
                    },
                    required = new[] { "sectionName", "changeType", "severity", "afs1Summary", "afs2Summary", "changeDescription", "changedValues" }
                }
            },
            ["lineItemChanges"] = new
            {
                type  = "array",
                items = new
                {
                    type = "object",
                    properties = new Dictionary<string, object>
                    {
                        ["lineItem"]        = new { type = "string" },
                        ["valueInAfs1"]     = new { type = "string" },
                        ["valueInAfs2"]     = new { type = "string" },
                        ["pctChange"]       = new { type = "string" },
                        ["materialityNote"] = new { type = "string" },
                        ["pageAfs1"]        = new { type = new[] { "integer", "null" } },
                        ["pageAfs2"]        = new { type = new[] { "integer", "null" } },
                    },
                    required = new[] { "lineItem", "valueInAfs1", "valueInAfs2", "pctChange", "materialityNote" }
                }
            },
            ["newDisclosures"]      = new { type = "array", items = new { type = "string" } },
            ["removedDisclosures"]  = new { type = "array", items = new { type = "string" } },
            ["policyChanges"]       = new { type = "array", items = new { type = "string" } },
            ["limitations"]         = new { type = "array", items = new { type = "string" } },
        },
        required = new[] { "overallAssessment", "riskLevel", "auditConclusion",
                           "sectionDiffs", "lineItemChanges",
                           "newDisclosures", "removedDisclosures", "policyChanges", "limitations" }
    };

    // ─────────────────────────────────────────────────────────────────────────
    private static void Normalize(AiContentReadResult r)
    {
        r.OverallAssessment = Trim(r.OverallAssessment, 1400);
        r.RiskLevel         = Trim(r.RiskLevel, 30);
        r.AuditConclusion   = Trim(r.AuditConclusion, 900);

        r.SectionDiffs = (r.SectionDiffs ?? new())
            .Where(s => !string.IsNullOrWhiteSpace(s.SectionName))
            .Select(s =>
            {
                s.SectionName       = Trim(s.SectionName, 200);
                s.ChangeType        = Trim(s.ChangeType, 50);
                s.Severity          = Trim(s.Severity, 20);
                s.Afs1Summary       = Trim(s.Afs1Summary, 2000);
                s.Afs2Summary       = Trim(s.Afs2Summary, 2000);
                s.ChangeDescription = Trim(s.ChangeDescription, 2000);
                s.ChangedValues     = (s.ChangedValues ?? new()).Select(v => Trim(v, 200)).ToList();
                return s;
            })
            .ToList(); // no Take — return all sections

        r.LineItemChanges = (r.LineItemChanges ?? new())
            .Where(l => !string.IsNullOrWhiteSpace(l.LineItem))
            .Select(l =>
            {
                l.LineItem        = Trim(l.LineItem, 200);
                l.ValueInAfs1     = Trim(l.ValueInAfs1, 100);
                l.ValueInAfs2     = Trim(l.ValueInAfs2, 100);
                l.PctChange       = Trim(l.PctChange, 40);
                l.MaterialityNote = Trim(l.MaterialityNote, 400);
                return l;
            })
            .ToList(); // no Take — return all line items

        r.NewDisclosures     = (r.NewDisclosures ?? new()).Select(x => Trim(x, 500)).Where(x => x.Length > 0).ToList();
        r.RemovedDisclosures = (r.RemovedDisclosures ?? new()).Select(x => Trim(x, 500)).Where(x => x.Length > 0).ToList();
        r.PolicyChanges      = (r.PolicyChanges ?? new()).Select(x => Trim(x, 600)).Where(x => x.Length > 0).ToList();
        r.Limitations        = (r.Limitations ?? new()).Select(x => Trim(x, 300)).Where(x => x.Length > 0).ToList();
    }

    // ─────────────────────────────────────────────────────────────────────────
    private static string ExtractText(string raw)
    {
        using var doc = JsonDocument.Parse(raw);
        if (!doc.RootElement.TryGetProperty("candidates", out var cands) ||
            cands.ValueKind != JsonValueKind.Array || cands.GetArrayLength() == 0)
            return string.Empty;

        var first = cands[0];
        if (!first.TryGetProperty("content", out var content) ||
            !content.TryGetProperty("parts", out var parts) ||
            parts.ValueKind != JsonValueKind.Array)
            return string.Empty;

        foreach (var part in parts.EnumerateArray())
            if (part.TryGetProperty("text", out var textNode))
                return textNode.GetString() ?? string.Empty;

        return string.Empty;
    }

    private static string StripFences(string text)
    {
        string t = text.Trim();
        if (!t.StartsWith("```", StringComparison.Ordinal)) return t;
        int nl    = t.IndexOf('\n');
        int fence = t.LastIndexOf("```", StringComparison.Ordinal);
        if (nl < 0 || fence <= nl) return t;
        return t[(nl + 1)..fence].Trim();
    }

    private static string Trim(string? value, int max)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        string t = value.Trim();
        return t.Length <= max ? t : t[..max] + "...";
    }

    private static AiContentReadResult Unavailable(string status, bool enabled = true) => new()
    {
        Enabled = enabled,
        Status = status,
        Message = AiUserFacingMessages.ContentReadUnavailable,
        Limitations = AiUserFacingMessages.ContentReadUnavailableLimitations(),
        GeneratedAt = DateTime.UtcNow,
    };
}
