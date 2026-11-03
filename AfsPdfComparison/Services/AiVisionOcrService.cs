using System.Diagnostics;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace AfsPdfComparison.Services;

public class AiVisionOcrAttempt
{
    public string Status { get; set; } = "disabled";
    public string Provider { get; set; } = "Gemini";
    public string Model { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public List<string> Pages { get; set; } = new();
}

public class AiVisionOcrService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private static readonly Regex SpaceRegex = new(@"\s+", RegexOptions.Compiled);

    private static readonly string[] PopplerBinCandidates =
    {
        @"C:\Program Files\poppler\Library\bin",
        @"C:\Program Files\poppler-24.02.0\Library\bin",
        @"C:\Program Files (x86)\poppler\bin",
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            @"Microsoft\WinGet\Packages\oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe\poppler-25.07.0\Library\bin"),
    };

    private readonly IHttpClientFactory _httpClientFactory;
    private readonly IConfiguration _configuration;
    private readonly ILogger<AiVisionOcrService> _logger;

    public AiVisionOcrService(
        IHttpClientFactory httpClientFactory,
        IConfiguration configuration,
        ILogger<AiVisionOcrService> logger)
    {
        _httpClientFactory = httpClientFactory;
        _configuration = configuration;
        _logger = logger;
    }

    public async Task<AiVisionOcrAttempt> TryExtractPagesAsync(
        string pdfPath,
        int expectedPageCount,
        CancellationToken cancellationToken = default)
    {
        string provider = _configuration["AiOcr:Provider"] ?? "Gemini";
        string model = AiConfiguration.GetString(
            _configuration,
            "gemini-2.5-flash",
            "AiOcr:GeminiModel",
            "AiAssist:GeminiModel",
            "GEMINI_MODEL");

        if (!AiConfiguration.GetBool(_configuration, "AiOcr:Enabled", true))
        {
            return new AiVisionOcrAttempt
            {
                Status = "disabled",
                Provider = provider,
                Model = model,
                Message = "AI OCR disabled in configuration.",
            };
        }

        if (expectedPageCount <= 0 || !File.Exists(pdfPath))
        {
            return new AiVisionOcrAttempt
            {
                Status = "error",
                Provider = provider,
                Model = model,
                Message = "PDF file missing or has no pages to process.",
            };
        }

        string apiKey = AiConfiguration.GetOptional(
            _configuration,
            "AiOcr:GeminiApiKey",
            "AiAssist:GeminiApiKey",
            "GEMINI_API_KEY") ?? string.Empty;

        if (string.IsNullOrWhiteSpace(apiKey))
        {
            return new AiVisionOcrAttempt
            {
                Status = "missing_key",
                Provider = provider,
                Model = model,
                Message = "Set GEMINI_API_KEY to enable AI OCR fallback.",
            };
        }

        int maxPages = GetInt("AiOcr:MaxPagesPerDocument", 80);
        if (expectedPageCount > maxPages)
        {
            return new AiVisionOcrAttempt
            {
                Status = "error",
                Provider = provider,
                Model = model,
                Message = $"Document has {expectedPageCount} pages which exceeds AI OCR limit of {maxPages}.",
            };
        }

        string? pdftoppm = FindExecutable("pdftoppm.exe", PopplerBinCandidates);
        if (pdftoppm == null)
        {
            return new AiVisionOcrAttempt
            {
                Status = "error",
                Provider = provider,
                Model = model,
                Message = "Poppler renderer not found for AI OCR page rendering.",
            };
        }

        var pages = new List<string>(expectedPageCount);
        try
        {
            for (int pageNum = 1; pageNum <= expectedPageCount; pageNum++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                byte[] imageBytes = RenderPageImage(pdftoppm, pdfPath, pageNum);
                if (imageBytes.Length == 0)
                {
                    pages.Add(string.Empty);
                    continue;
                }

                string pageText = await ExtractPageTextAsync(
                    imageBytes,
                    pageNum,
                    provider,
                    model,
                    apiKey,
                    cancellationToken);
                pages.Add(NormalizePageText(pageText));
            }

            return new AiVisionOcrAttempt
            {
                Status = "ready",
                Provider = provider,
                Model = model,
                Message = "AI OCR completed.",
                Pages = pages,
            };
        }
        catch (OperationCanceledException)
        {
            return new AiVisionOcrAttempt
            {
                Status = "error",
                Provider = provider,
                Model = model,
                Message = "AI OCR timed out.",
                Pages = pages,
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "AI OCR fallback failed for {PdfPath}", pdfPath);
            return new AiVisionOcrAttempt
            {
                Status = "error",
                Provider = provider,
                Model = model,
                Message = Trim(ex.Message, 240),
                Pages = pages,
            };
        }
    }

    private async Task<string> ExtractPageTextAsync(
        byte[] imageBytes,
        int pageNum,
        string provider,
        string model,
        string apiKey,
        CancellationToken cancellationToken)
    {
        string modelPath = model.StartsWith("models/", StringComparison.OrdinalIgnoreCase)
            ? model
            : "models/" + model;
        string endpoint = $"https://generativelanguage.googleapis.com/v1beta/{modelPath}:generateContent";
        string prompt =
            $"You are performing OCR on page {pageNum} of a financial statement PDF. " +
            "Extract the visible text exactly in reading order. Preserve meaningful line breaks. " +
            "Do not summarize, explain, label sections, or correct numbers. " +
            "If the page is blank or unreadable, return an empty string.";

        object requestBody = new
        {
            contents = new[]
            {
                new
                {
                    parts = new object[]
                    {
                        new
                        {
                            inline_data = new
                            {
                                mime_type = "image/png",
                                data = Convert.ToBase64String(imageBytes),
                            }
                        },
                        new { text = prompt }
                    }
                }
            },
            generationConfig = new
            {
                temperature = 0.0,
                responseMimeType = "application/json",
                responseJsonSchema = new
                {
                    type = "object",
                    properties = new Dictionary<string, object>
                    {
                        ["pageText"] = new
                        {
                            type = "string",
                            description = "The extracted page text in reading order with preserved line breaks."
                        }
                    },
                    required = new[] { "pageText" }
                }
            }
        };

        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        cts.CancelAfter(TimeSpan.FromSeconds(45));

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
            _logger.LogWarning(
                "AI OCR request failed on page {PageNum}: HTTP {StatusCode} {Body}",
                pageNum,
                (int)response.StatusCode,
                Trim(raw, 500));
            return string.Empty;
        }

        string jsonText = ExtractModelText(raw);
        if (string.IsNullOrWhiteSpace(jsonText))
            return string.Empty;

        string normalizedJson = StripCodeFences(jsonText);
        using var doc = JsonDocument.Parse(normalizedJson);
        if (doc.RootElement.TryGetProperty("pageText", out JsonElement pageText))
            return pageText.GetString() ?? string.Empty;

        return string.Empty;
    }

    private static byte[] RenderPageImage(string pdftoppm, string pdfPath, int pageNum)
    {
        string tempDir = Path.Combine(Path.GetTempPath(), "afs_aiocr_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        try
        {
            string prefix = Path.Combine(tempDir, "page");
            string imagePath = Path.Combine(tempDir, $"page-{pageNum:D2}.png");
            var psi = new ProcessStartInfo(pdftoppm)
            {
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            };
            psi.ArgumentList.Add("-png");
            psi.ArgumentList.Add("-r");
            psi.ArgumentList.Add("180");
            psi.ArgumentList.Add("-f");
            psi.ArgumentList.Add(pageNum.ToString());
            psi.ArgumentList.Add("-l");
            psi.ArgumentList.Add(pageNum.ToString());
            psi.ArgumentList.Add(pdfPath);
            psi.ArgumentList.Add(prefix);

            using var proc = Process.Start(psi);
            if (proc == null)
                return Array.Empty<byte>();

            if (!proc.WaitForExit(30_000))
            {
                try { proc.Kill(true); } catch { }
                return Array.Empty<byte>();
            }

            return File.Exists(imagePath) ? File.ReadAllBytes(imagePath) : Array.Empty<byte>();
        }
        catch
        {
            return Array.Empty<byte>();
        }
        finally
        {
            try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); } catch { }
        }
    }

    private static string NormalizePageText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        var sb = new StringBuilder();
        foreach (string rawLine in text.Replace("\r", "").Split('\n'))
        {
            string line = SpaceRegex.Replace(rawLine.Trim(), " ");
            if (line.Length > 0)
                sb.Append(line).Append('\n');
        }

        return sb.ToString();
    }

    private int GetInt(string key, int defaultValue)
    {
        string? raw = _configuration[key];
        return int.TryParse(raw, out int parsed) ? parsed : defaultValue;
    }

    private static string? FindExecutable(string fileName, IEnumerable<string> directCandidates)
    {
        string? fromPath = FindExecutableOnPath(fileName);
        if (!string.IsNullOrWhiteSpace(fromPath))
            return fromPath;

        foreach (string candidate in directCandidates)
        {
            if (string.IsNullOrWhiteSpace(candidate))
                continue;

            if (candidate.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
            {
                if (File.Exists(candidate))
                    return candidate;
                continue;
            }

            string fullPath = Path.Combine(candidate, fileName);
            if (File.Exists(fullPath))
                return fullPath;
        }

        return null;
    }

    private static string? FindExecutableOnPath(string fileName)
    {
        string? pathValue = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(pathValue))
            return null;

        foreach (string dir in pathValue.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries))
        {
            try
            {
                string fullPath = Path.Combine(dir.Trim(), fileName);
                if (File.Exists(fullPath))
                    return fullPath;
            }
            catch
            {
                // Ignore invalid PATH segments.
            }
        }

        return null;
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
}
