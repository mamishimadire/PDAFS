using AfsPdfComparison.Models;

namespace AfsPdfComparison.Services;

internal static class AiUserFacingMessages
{
    public const string ReviewUnavailable =
        "Intelligent review is currently unavailable. The document comparison is still available.";

    public const string ContentReadUnavailable =
        "AI content reading is currently unavailable. The document comparison is still available.";

    public static List<string> ReviewUnavailableLimitations() =>
        new() { "The intelligent review could not be completed for this comparison run." };

    public static List<string> ContentReadUnavailableLimitations() =>
        new() { "The AI content reading step could not be completed for this comparison run." };

    public static AiAssistResult CreateUnavailableReviewResult(IEnumerable<AiAssistResult> attempts)
    {
        var snapshot = attempts.ToList();
        bool configurationOnly = snapshot.Count > 0 && snapshot.All(r =>
            string.Equals(r.Status, "disabled", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(r.Status, "missing_key", StringComparison.OrdinalIgnoreCase));

        var providers = snapshot
            .Select(r => r.Provider)
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Distinct(StringComparer.OrdinalIgnoreCase);

        var models = snapshot
            .Select(r => r.Model)
            .Where(m => !string.IsNullOrWhiteSpace(m))
            .Distinct(StringComparer.OrdinalIgnoreCase);

        return new AiAssistResult
        {
            Enabled = !configurationOnly,
            Provider = string.Join(" + ", providers),
            Model = string.Join(", ", models),
            Status = configurationOnly ? "disabled" : "error",
            Message = ReviewUnavailable,
            Limitations = ReviewUnavailableLimitations(),
            GeneratedAt = DateTime.UtcNow,
        };
    }
}
