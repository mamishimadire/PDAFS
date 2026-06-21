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
    private readonly IConfiguration     _config;
    private readonly ILogger<AiContentReaderService> _logger;

    public AiContentReaderService(
        IHttpClientFactory http,
        IConfiguration     config,
        ILogger<AiContentReaderService> logger)
    {
        _http   = http;
        _config = config;
        _logger = logger;
    }

    public async Task<AiContentReadResult> ReadAndCompareAsync(
        PdfReport report1,
        PdfReport report2,
        CancellationToken cancellationToken = default)
    {
        bool enabled = GetBool("AiAssist:Enabled", defaultValue: true);
        string model = _config["AiAssist:GeminiModel"] ?? "gemini-2.5-flash";

        if (!enabled)
            return Unavailable("disabled", enabled: false);

        string apiKey =
            _config["AiAssist:GeminiApiKey"] ??
            _config["GEMINI_API_KEY"] ??
            Environment.GetEnvironmentVariable("GEMINI_API_KEY") ??
            string.Empty;

        if (string.IsNullOrWhiteSpace(apiKey))
            return Unavailable("missing_key", enabled: false);

        try
        {
            string modelPath = model.StartsWith("models/", StringComparison.OrdinalIgnoreCase)
                ? model : "models/" + model;

            string endpoint = $"https://generativelanguage.googleapis.com/v1beta/{modelPath}:generateContent";
            string prompt   = BuildContentReadPrompt(report1, report2);

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

            var result = JsonSerializer.Deserialize<AiContentReadResult>(
                StripFences(jsonText),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new AiContentReadResult();

            result.Enabled     = true;
            result.Status      = "ready";
            result.Message     = "AI content reading completed.";
            result.GeneratedAt = DateTime.UtcNow;
            Normalize(result);
            return result;
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("AiContentReader: request timed out after {TimeoutSeconds}s.", 120);
            return Unavailable("error");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "AiContentReader: unexpected failure.");
            return Unavailable("error");
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    private string BuildContentReadPrompt(PdfReport report1, PdfReport report2)
    {
        // Strategy: send the full text of each document up to a combined budget
        // of ~60 000 characters (well within Gemini's context window), taking
        // the beginning of each document (where cover, policies, and key
        // financials usually appear) plus evenly spaced pages across the rest.
        const int maxCharsPerDoc = 28_000;

        string text1 = SampleDocumentText(report1, maxCharsPerDoc);
        string text2 = SampleDocumentText(report2, maxCharsPerDoc);

        return
            "You are an expert financial auditor at SNG Grant Thornton. " +
            "You have extracted the text of two versions of an Annual Financial Statement (AFS). " +
            "Read both documents carefully and perform a comprehensive comparison.\n\n" +
            "Your job:\n" +
            "1. Identify every section (Balance Sheet, Income Statement, Cash Flow, Notes, Policies, Directors' Report, Auditors' Report, etc.)\n" +
            "2. For each section, summarise what changed between AFS 1 and AFS 2\n" +
            "3. List all numeric line item changes you can find (e.g. 'Revenue: R45.2M → R48.6M')\n" +
            "4. Flag accounting policy changes, new disclosures, and removed disclosures\n" +
            "5. Provide an overall materiality assessment and audit conclusion\n\n" +
            "Return ONLY valid JSON matching the provided schema. Do not include markdown or code fences.\n\n" +
            "=== AFS 1 — " + report1.Label + " (" + report1.Filename + ") ===\n" +
            text1 +
            "\n\n=== AFS 2 — " + report2.Label + " (" + report2.Filename + ") ===\n" +
            text2;
    }

    private static string SampleDocumentText(PdfReport report, int maxChars)
    {
        if (string.IsNullOrWhiteSpace(report.FullText))
            return "(No text extracted — document may be scanned or empty)";

        if (report.FullText.Length <= maxChars)
            return report.FullText;

        // Take first 40% (cover + opening statements) + last 60% (notes)
        int headLen = (int)(maxChars * 0.4);
        int tailLen = maxChars - headLen;
        string head = report.FullText[..headLen];
        string tail = report.FullText[^tailLen..];
        return head + "\n\n[... middle pages omitted for brevity ...]\n\n" + tail;
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
                s.SectionName       = Trim(s.SectionName, 120);
                s.ChangeType        = Trim(s.ChangeType, 30);
                s.Severity          = Trim(s.Severity, 20);
                s.Afs1Summary       = Trim(s.Afs1Summary, 400);
                s.Afs2Summary       = Trim(s.Afs2Summary, 400);
                s.ChangeDescription = Trim(s.ChangeDescription, 500);
                s.ChangedValues     = (s.ChangedValues ?? new()).Select(v => Trim(v, 120)).Take(10).ToList();
                return s;
            })
            .Take(20)
            .ToList();

        r.LineItemChanges = (r.LineItemChanges ?? new())
            .Where(l => !string.IsNullOrWhiteSpace(l.LineItem))
            .Select(l =>
            {
                l.LineItem        = Trim(l.LineItem, 120);
                l.ValueInAfs1     = Trim(l.ValueInAfs1, 60);
                l.ValueInAfs2     = Trim(l.ValueInAfs2, 60);
                l.PctChange       = Trim(l.PctChange, 30);
                l.MaterialityNote = Trim(l.MaterialityNote, 200);
                return l;
            })
            .Take(30)
            .ToList();

        r.NewDisclosures     = (r.NewDisclosures ?? new()).Select(x => Trim(x, 220)).Where(x => x.Length > 0).Take(12).ToList();
        r.RemovedDisclosures = (r.RemovedDisclosures ?? new()).Select(x => Trim(x, 220)).Where(x => x.Length > 0).Take(12).ToList();
        r.PolicyChanges      = (r.PolicyChanges ?? new()).Select(x => Trim(x, 280)).Where(x => x.Length > 0).Take(10).ToList();
        r.Limitations        = (r.Limitations ?? new()).Select(x => Trim(x, 220)).Where(x => x.Length > 0).Take(6).ToList();
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

    private bool GetBool(string key, bool defaultValue)
    {
        string? raw = _config[key];
        return bool.TryParse(raw, out bool parsed) ? parsed : defaultValue;
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
