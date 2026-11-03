using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using AfsPdfComparison.Models;
using Microsoft.Extensions.Logging;

namespace AfsPdfComparison.Services;

/// <summary>
/// Calls any OpenAI-compatible chat-completion endpoint (Groq, OpenRouter, etc.)
/// and returns an AiAssistResult using the same prompt as the Gemini service.
/// </summary>
public class OpenAiCompatibleJsonResponse
{
    public bool Success { get; set; }
    public string Provider { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public string Status { get; set; } = "error";
    public string JsonText { get; set; } = string.Empty;
}

public class OpenAiCompatibleAiService
{
    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy        = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition      = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented               = false,
    };

    private readonly IHttpClientFactory _http;
    private readonly ILogger<OpenAiCompatibleAiService> _logger;

    public OpenAiCompatibleAiService(IHttpClientFactory http, ILogger<OpenAiCompatibleAiService> logger)
    {
        _http   = http;
        _logger = logger;
    }

    public async Task<AiAssistResult> AnalyzeAsync(
        string endpoint,
        string apiKey,
        string model,
        string providerName,
        string prompt,
        int timeoutSeconds = 90,
        CancellationToken cancellationToken = default)
    {
        const string systemPrompt =
            "You are an expert financial auditor AI at SNG Grant Thornton. " +
            "You read Annual Financial Statements (AFS) and identify ALL material differences " +
            "between two versions, including directorate changes, policy changes, numeric changes, " +
            "and removed/added disclosures. Be thorough — do not miss single-line changes.";

        var response = await RequestJsonObjectAsync(
            endpoint,
            apiKey,
            model,
            providerName,
            systemPrompt,
            prompt,
            timeoutSeconds,
            cancellationToken);

        if (!response.Success)
        {
            return string.Equals(response.Status, "disabled", StringComparison.OrdinalIgnoreCase)
                ? Disabled(providerName, model)
                : Error(providerName, model);
        }

        var result = JsonSerializer.Deserialize<AiAssistResult>(
            response.JsonText,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new AiAssistResult();

        result.Enabled     = true;
        result.Provider    = providerName;
        result.Model       = model;
        result.Status      = "ready";
        result.Message     = $"{providerName} analysis complete.";
        result.GeneratedAt = DateTime.UtcNow;
        return result;
    }

    public async Task<OpenAiCompatibleJsonResponse> RequestJsonObjectAsync(
        string endpoint,
        string apiKey,
        string model,
        string providerName,
        string systemPrompt,
        string userPrompt,
        int timeoutSeconds = 90,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(apiKey))
        {
            return new OpenAiCompatibleJsonResponse
            {
                Provider = providerName,
                Model = model,
                Status = "disabled",
            };
        }

        try
        {
            var messages = new[]
            {
                new { role = "system", content = systemPrompt },
                new { role = "user", content = userPrompt }
            };

            var requestBody = new
            {
                model,
                messages,
                temperature     = 0.15,
                max_tokens      = 4096,
                response_format = new { type = "json_object" }
            };

            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

            using var req = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(requestBody, JsonOpts),
                    Encoding.UTF8, "application/json")
            };
            req.Headers.Add("Authorization", $"Bearer {apiKey}");
            if (endpoint.Contains("openrouter", StringComparison.OrdinalIgnoreCase))
                req.Headers.Add("HTTP-Referer", "https://sng-afs-comparison.local");

            var client = _http.CreateClient();
            using var resp = await client.SendAsync(req, cts.Token);
            string raw = await resp.Content.ReadAsStringAsync(cts.Token);

            if (!resp.IsSuccessStatusCode)
            {
                _logger.LogWarning("[{Provider}] HTTP {Code}: {Body}", providerName, (int)resp.StatusCode, Trim(raw, 300));
                return new OpenAiCompatibleJsonResponse
                {
                    Provider = providerName,
                    Model = model,
                    Status = "error",
                };
            }

            string text = ExtractContent(raw);
            if (string.IsNullOrWhiteSpace(text))
            {
                _logger.LogWarning("[{Provider}] returned an empty response body.", providerName);
                return new OpenAiCompatibleJsonResponse
                {
                    Provider = providerName,
                    Model = model,
                    Status = "error",
                };
            }

            return new OpenAiCompatibleJsonResponse
            {
                Success = true,
                Provider = providerName,
                Model = model,
                Status = "ready",
                JsonText = StripFences(text),
            };
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            _logger.LogDebug("[{Provider}] request cancelled because another provider completed first.", providerName);
            return new OpenAiCompatibleJsonResponse
            {
                Provider = providerName,
                Model = model,
                Status = "cancelled",
            };
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("[{Provider}] timed out after {TimeoutSeconds}s.", providerName, timeoutSeconds);
            return new OpenAiCompatibleJsonResponse
            {
                Provider = providerName,
                Model = model,
                Status = "error",
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "[{Provider}] unexpected failure.", providerName);
            return new OpenAiCompatibleJsonResponse
            {
                Provider = providerName,
                Model = model,
                Status = "error",
            };
        }
    }

    // ── helpers ──────────────────────────────────────────────────────────────
    private static string ExtractContent(string raw)
    {
        using var doc = JsonDocument.Parse(raw);
        if (doc.RootElement.TryGetProperty("choices", out var choices) &&
            choices.ValueKind == JsonValueKind.Array &&
            choices.GetArrayLength() > 0)
        {
            var first = choices[0];
            if (first.TryGetProperty("message", out var msg) &&
                msg.TryGetProperty("content", out var content))
                return content.GetString() ?? string.Empty;
        }
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

    private static string Trim(string? s, int max)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        string t = s.Trim();
        return t.Length <= max ? t : t[..max] + "...";
    }

    private static AiAssistResult Disabled(string provider, string model) => new()
    {
        Enabled = false,
        Provider = provider,
        Model = model,
        Status = "disabled",
        Message = AiUserFacingMessages.ReviewUnavailable,
        Limitations = AiUserFacingMessages.ReviewUnavailableLimitations(),
        GeneratedAt = DateTime.UtcNow,
    };

    private static AiAssistResult Error(string provider, string model) => new()
    {
        Enabled = true,
        Provider = provider,
        Model = model,
        Status = "error",
        Message = AiUserFacingMessages.ReviewUnavailable,
        Limitations = AiUserFacingMessages.ReviewUnavailableLimitations(),
        GeneratedAt = DateTime.UtcNow,
    };
}
