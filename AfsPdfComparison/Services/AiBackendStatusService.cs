using AfsPdfComparison.Models;
using Microsoft.Extensions.Configuration;

namespace AfsPdfComparison.Services;

public class AiBackendStatusService
{
    private readonly IConfiguration _configuration;

    public AiBackendStatusService(IConfiguration configuration)
    {
        _configuration = configuration;
    }

    public AiBackendStatusSnapshot GetSnapshot()
    {
        bool aiEnabled = AiConfiguration.GetBool(_configuration, "AiAssist:Enabled", true);
        var providers = new List<AiBackendProviderStatus>
        {
            BuildProviderStatus(
                "gemini",
                "Gemini",
                aiEnabled,
                AiConfiguration.GetString(_configuration, "gemini-2.5-flash", "AiAssist:GeminiModel", "GEMINI_MODEL"),
                AiConfiguration.GetOptional(_configuration, "AiAssist:GeminiApiKey", "GEMINI_API_KEY")),
            BuildProviderStatus(
                "groq",
                "Groq",
                aiEnabled,
                AiConfiguration.GetString(_configuration, "llama-3.3-70b-versatile", "AiAssist:GroqModel", "GROQ_MODEL"),
                AiConfiguration.GetOptional(_configuration, "AiAssist:GroqApiKey", "GROQ_API_KEY")),
            BuildProviderStatus(
                "openrouter",
                "OpenRouter",
                aiEnabled,
                AiConfiguration.GetString(_configuration, "deepseek/deepseek-chat", "AiAssist:OpenRouterModel", "OPENROUTER_MODEL"),
                AiConfiguration.GetOptional(_configuration, "AiAssist:OpenRouterApiKey", "OPENROUTER_API_KEY")),
        };

        int activeCount = providers.Count(p => p.Active);
        return new AiBackendStatusSnapshot
        {
            BoostActive = activeCount > 0,
            Summary = $"{activeCount}/{providers.Count} model(s) active",
            Providers = providers,
        };
    }

    private static AiBackendProviderStatus BuildProviderStatus(
        string key,
        string name,
        bool enabled,
        string model,
        string? apiKey)
    {
        bool active = enabled && !string.IsNullOrWhiteSpace(apiKey);
        string reason = !enabled
            ? "AI review disabled"
            : active
                ? "Configured"
                : "Missing API key";

        return new AiBackendProviderStatus
        {
            Key = key,
            Name = name,
            Model = model,
            Active = active,
            Reason = reason,
        };
    }
}
