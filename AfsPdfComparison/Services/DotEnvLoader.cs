using System.Threading;

namespace AfsPdfComparison.Services;

internal static class DotEnvLoader
{
    private static int _loaded;

    public static void LoadIfPresent()
    {
        if (Interlocked.Exchange(ref _loaded, 1) == 1)
            return;

        foreach (string candidate in GetCandidates())
        {
            if (!File.Exists(candidate))
                continue;

            LoadFile(candidate);
            break;
        }
    }

    private static IEnumerable<string> GetCandidates()
    {
        string? explicitPath =
            Environment.GetEnvironmentVariable("AFS_DOTENV_PATH") ??
            Environment.GetEnvironmentVariable("DOTENV_PATH");

        if (!string.IsNullOrWhiteSpace(explicitPath))
            yield return explicitPath;

        string current = Directory.GetCurrentDirectory();
        yield return Path.Combine(current, ".env");

        string? parent = Directory.GetParent(current)?.FullName;
        if (!string.IsNullOrWhiteSpace(parent))
            yield return Path.Combine(parent, ".env");

        yield return Path.Combine(AppContext.BaseDirectory, ".env");
    }

    private static void LoadFile(string path)
    {
        foreach (string rawLine in File.ReadAllLines(path))
        {
            string line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
                continue;

            int separator = line.IndexOf('=');
            if (separator <= 0)
                continue;

            string key = line[..separator].Trim();
            if (key.Length == 0)
                continue;

            string value = line[(separator + 1)..].Trim();
            if (value.Length >= 2 &&
                ((value[0] == '"' && value[^1] == '"') ||
                 (value[0] == '\'' && value[^1] == '\'')))
            {
                value = value[1..^1];
            }

            if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(key)))
                Environment.SetEnvironmentVariable(key, value);
        }
    }
}
