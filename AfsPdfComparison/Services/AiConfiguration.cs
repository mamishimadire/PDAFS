using Microsoft.Extensions.Configuration;

namespace AfsPdfComparison.Services;

internal static class AiConfiguration
{
    public static string? GetOptional(IConfiguration configuration, params string[] keys)
    {
        foreach (string key in keys)
        {
            string? value = configuration[key];
            if (!string.IsNullOrWhiteSpace(value))
                return value.Trim();
        }

        return null;
    }

    public static string GetString(
        IConfiguration configuration,
        string fallback,
        params string[] keys) =>
        GetOptional(configuration, keys) ?? fallback;

    public static bool GetBool(IConfiguration configuration, string key, bool fallback)
    {
        string? raw = configuration[key];
        return bool.TryParse(raw, out bool parsed) ? parsed : fallback;
    }

    public static int GetInt(IConfiguration configuration, int fallback, params string[] keys)
    {
        foreach (string key in keys)
        {
            string? raw = configuration[key];
            if (int.TryParse(raw, out int parsed))
                return parsed;
        }

        return fallback;
    }
}
