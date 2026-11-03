using System.Net;
using System.Net.Http;
using System.Diagnostics;
using System.Text;
using AfsPdfComparison.Models;
using AfsPdfComparison.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;

namespace AfsPdfComparison.Tests;

public class AiFallbackBehaviorTests
{
    [Fact]
    public async Task AnalyzeAsync_UsesFallbackModelWhenGeminiIsUnavailable()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            string uri = request.RequestUri!.ToString();

            if (uri.Contains("generativelanguage.googleapis.com", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":{"message":"Gemini overloaded"}}""");

            if (uri.Contains("groq.test", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.OK, """
{"choices":[{"message":{"content":"{\"executiveSummary\":\"Groq found the material change.\",\"overallRisk\":\"High\",\"materialFindings\":[],\"pageFocus\":[],\"highlightGuidance\":[\"Review the revenue note.\"],\"limitations\":[]}"}}]}
""");

            throw new InvalidOperationException("Unexpected request URI: " + uri);
        });

        var orchestrator = CreateOrchestrator(handler, new Dictionary<string, string?>
        {
            ["AiAssist:GeminiApiKey"] = "gemini-test-key",
            ["GROQ_API_KEY"] = "groq-test-key",
            ["AiAssist:GroqUrl"] = "https://groq.test/openai/v1/chat/completions",
        });

        AiAssistResult result = await orchestrator.AnalyzeAsync(CreateComparisonResult());

        Assert.Equal("ready", result.Status);
        Assert.Contains("Groq", result.Provider, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Gemini request failed", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("503", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task AnalyzeAsync_ReturnsWithoutWaitingForSlowGeminiWhenGroqIsReady()
    {
        var handler = new StubHttpMessageHandler(async (request, cancellationToken) =>
        {
            string uri = request.RequestUri!.ToString();

            if (uri.Contains("generativelanguage.googleapis.com", StringComparison.OrdinalIgnoreCase))
            {
                await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":{"message":"Gemini overloaded"}}""");
            }

            if (uri.Contains("groq.test", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.OK, """
{"choices":[{"message":{"content":"{\"executiveSummary\":\"Groq found the material change.\",\"overallRisk\":\"High\",\"materialFindings\":[],\"pageFocus\":[],\"highlightGuidance\":[],\"limitations\":[]}"}}]}
""");

            throw new InvalidOperationException("Unexpected request URI: " + uri);
        });

        var orchestrator = CreateOrchestrator(handler, new Dictionary<string, string?>
        {
            ["AiAssist:GeminiApiKey"] = "gemini-test-key",
            ["GROQ_API_KEY"] = "groq-test-key",
            ["AiAssist:GroqUrl"] = "https://groq.test/openai/v1/chat/completions",
            ["AiAssist:ConsensusWindowMs"] = "0",
        });

        var sw = Stopwatch.StartNew();
        AiAssistResult result = await orchestrator.AnalyzeAsync(CreateComparisonResult());
        sw.Stop();

        Assert.Equal("ready", result.Status);
        Assert.Contains("Groq", result.Provider, StringComparison.OrdinalIgnoreCase);
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(2), $"Expected fast failover, but took {sw.Elapsed}.");
    }

    [Fact]
    public async Task AnalyzeAsync_HidesInternalFailureDetailsWhenAllModelsFail()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            string uri = request.RequestUri!.ToString();

            if (uri.Contains("generativelanguage.googleapis.com", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":{"message":"Gemini overloaded"}}""");

            if (uri.Contains("groq.test", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":"Groq unavailable"}""");

            throw new InvalidOperationException("Unexpected request URI: " + uri);
        });

        var orchestrator = CreateOrchestrator(handler, new Dictionary<string, string?>
        {
            ["AiAssist:GeminiApiKey"] = "gemini-test-key",
            ["GROQ_API_KEY"] = "groq-test-key",
            ["AiAssist:GroqUrl"] = "https://groq.test/openai/v1/chat/completions",
        });

        AiAssistResult result = await orchestrator.AnalyzeAsync(CreateComparisonResult());

        Assert.NotEqual("ready", result.Status);
        Assert.Equal(
            "Intelligent review is currently unavailable. The document comparison is still available.",
            result.Message);
        Assert.DoesNotContain("Gemini", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("503", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.All(result.Limitations, item =>
        {
            Assert.DoesNotContain("Gemini", item, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("503", item, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("API key", item, StringComparison.OrdinalIgnoreCase);
        });
    }

    [Fact]
    public async Task ReadAndCompareAsync_UsesOpenRouterWhenGeminiIsUnavailable()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            string uri = request.RequestUri!.ToString();

            if (uri.Contains("generativelanguage.googleapis.com", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":{"message":"Gemini overloaded"}}""");

            if (uri.Contains("openrouter.test", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.OK, """
{"choices":[{"message":{"content":"{\"overallAssessment\":\"OpenRouter reviewed the statements.\",\"riskLevel\":\"Medium\",\"auditConclusion\":\"Review the revenue and cash flow note.\",\"sectionDiffs\":[],\"lineItemChanges\":[],\"newDisclosures\":[],\"removedDisclosures\":[],\"policyChanges\":[],\"limitations\":[]}"}}]}
""");

            throw new InvalidOperationException("Unexpected request URI: " + uri);
        });

        var reader = CreateContentReader(handler, new Dictionary<string, string?>
        {
            ["AiAssist:GeminiApiKey"] = "gemini-test-key",
            ["OPENROUTER_API_KEY"] = "openrouter-test-key",
            ["AiAssist:OpenRouterUrl"] = "https://openrouter.test/api/v1/chat/completions",
        });

        AiContentReadResult result = await reader.ReadAndCompareAsync(
            CreatePdfReport("AFS 1", "afs-1.pdf", "Revenue 100"),
            CreatePdfReport("AFS 2", "afs-2.pdf", "Revenue 120"));

        Assert.Equal("ready", result.Status);
        Assert.Contains("OpenRouter", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("OpenRouter reviewed the statements.", result.OverallAssessment);
        Assert.DoesNotContain("Gemini", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("503", result.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ReadAndCompareAsync_ReturnsWithoutWaitingForSlowGeminiWhenOpenRouterIsReady()
    {
        var handler = new StubHttpMessageHandler(async (request, cancellationToken) =>
        {
            string uri = request.RequestUri!.ToString();

            if (uri.Contains("generativelanguage.googleapis.com", StringComparison.OrdinalIgnoreCase))
            {
                await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":{"message":"Gemini overloaded"}}""");
            }

            if (uri.Contains("openrouter.test", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.OK, """
{"choices":[{"message":{"content":"{\"overallAssessment\":\"OpenRouter reviewed the statements.\",\"riskLevel\":\"Medium\",\"auditConclusion\":\"Review the revenue and cash flow note.\",\"sectionDiffs\":[],\"lineItemChanges\":[],\"newDisclosures\":[],\"removedDisclosures\":[],\"policyChanges\":[],\"limitations\":[]}"}}]}
""");

            throw new InvalidOperationException("Unexpected request URI: " + uri);
        });

        var reader = CreateContentReader(handler, new Dictionary<string, string?>
        {
            ["AiAssist:GeminiApiKey"] = "gemini-test-key",
            ["OPENROUTER_API_KEY"] = "openrouter-test-key",
            ["AiAssist:OpenRouterUrl"] = "https://openrouter.test/api/v1/chat/completions",
        });

        var sw = Stopwatch.StartNew();
        AiContentReadResult result = await reader.ReadAndCompareAsync(
            CreatePdfReport("AFS 1", "afs-1.pdf", "Revenue 100"),
            CreatePdfReport("AFS 2", "afs-2.pdf", "Revenue 120"));
        sw.Stop();

        Assert.Equal("ready", result.Status);
        Assert.Contains("OpenRouter", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.True(sw.Elapsed < TimeSpan.FromSeconds(2), $"Expected fast failover, but took {sw.Elapsed}.");
    }

    [Fact]
    public async Task ReadAndCompareAsync_HidesInternalFailureDetailsWhenAllModelsFail()
    {
        var handler = new StubHttpMessageHandler(request =>
        {
            string uri = request.RequestUri!.ToString();

            if (uri.Contains("generativelanguage.googleapis.com", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":{"message":"Gemini overloaded"}}""");

            if (uri.Contains("openrouter.test", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":"OpenRouter unavailable"}""");

            if (uri.Contains("groq.test", StringComparison.OrdinalIgnoreCase))
                return Json(HttpStatusCode.ServiceUnavailable, """{"error":"Groq unavailable"}""");

            throw new InvalidOperationException("Unexpected request URI: " + uri);
        });

        var reader = CreateContentReader(handler, new Dictionary<string, string?>
        {
            ["AiAssist:GeminiApiKey"] = "gemini-test-key",
            ["OPENROUTER_API_KEY"] = "openrouter-test-key",
            ["AiAssist:OpenRouterUrl"] = "https://openrouter.test/api/v1/chat/completions",
            ["GROQ_API_KEY"] = "groq-test-key",
            ["AiAssist:GroqUrl"] = "https://groq.test/openai/v1/chat/completions",
        });

        AiContentReadResult result = await reader.ReadAndCompareAsync(
            CreatePdfReport("AFS 1", "afs-1.pdf", "Revenue 100"),
            CreatePdfReport("AFS 2", "afs-2.pdf", "Revenue 120"));

        Assert.NotEqual("ready", result.Status);
        Assert.Equal(
            "AI content reading is currently unavailable. The document comparison is still available.",
            result.Message);
        Assert.DoesNotContain("Gemini", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("503", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.All(result.Limitations, item =>
        {
            Assert.DoesNotContain("Gemini", item, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("503", item, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("API key", item, StringComparison.OrdinalIgnoreCase);
        });
    }

    private static MultiAiOrchestrator CreateOrchestrator(
        HttpMessageHandler handler,
        IDictionary<string, string?> settings)
    {
        var config = new ConfigurationBuilder()
            .AddInMemoryCollection(settings)
            .Build();

        var httpFactory = new StubHttpClientFactory(handler);
        var gemini = new AiComparisonAssistantService(
            httpFactory,
            config,
            NullLogger<AiComparisonAssistantService>.Instance);
        var openAi = new OpenAiCompatibleAiService(
            httpFactory,
            NullLogger<OpenAiCompatibleAiService>.Instance);

        return new MultiAiOrchestrator(
            gemini,
            openAi,
            config,
            NullLogger<MultiAiOrchestrator>.Instance);
    }

    private static AiContentReaderService CreateContentReader(
        HttpMessageHandler handler,
        IDictionary<string, string?> settings)
    {
        var config = new ConfigurationBuilder()
            .AddInMemoryCollection(settings)
            .Build();

        var httpFactory = new StubHttpClientFactory(handler);
        var openAi = new OpenAiCompatibleAiService(
            httpFactory,
            NullLogger<OpenAiCompatibleAiService>.Instance);

        return new AiContentReaderService(
            httpFactory,
            openAi,
            config,
            NullLogger<AiContentReaderService>.Instance);
    }

    private static ComparisonResult CreateComparisonResult() => new()
    {
        Report1 = CreatePdfReport("AFS 1", "afs-1.pdf", "Revenue 100", 2024),
        Report2 = CreatePdfReport("AFS 2", "afs-2.pdf", "Revenue 120", 2025),
        Counts = new Dictionary<string, int>
        {
            ["same"] = 0,
            ["changed"] = 1,
            ["added"] = 0,
            ["removed"] = 0,
        },
        PageDiffs = new List<PageDiffResult>(),
    };

    private static PdfReport CreatePdfReport(string label, string filename, string text, int primaryYear = 2025) => new()
    {
        Label = label,
        Filename = filename,
        Pages = new List<string> { text },
        FullText = text,
        PageCount = 1,
        PrimaryYear = primaryYear,
    };

    private static HttpResponseMessage Json(HttpStatusCode statusCode, string body) => new(statusCode)
    {
        Content = new StringContent(body, Encoding.UTF8, "application/json"),
    };

    private sealed class StubHttpClientFactory : IHttpClientFactory
    {
        private readonly HttpClient _client;

        public StubHttpClientFactory(HttpMessageHandler handler)
        {
            _client = new HttpClient(handler, disposeHandler: false);
        }

        public HttpClient CreateClient(string name) => _client;
    }

    private sealed class StubHttpMessageHandler : HttpMessageHandler
    {
        private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _handler;

        public StubHttpMessageHandler(Func<HttpRequestMessage, HttpResponseMessage> handler)
            : this((request, _) => Task.FromResult(handler(request)))
        {
        }

        public StubHttpMessageHandler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> handler)
        {
            _handler = handler;
        }

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            return _handler(request, cancellationToken);
        }
    }
}
