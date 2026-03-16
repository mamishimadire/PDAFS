// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3                                                               ║
// ║                                                                              ║
// ║  SERVICE — SessionStateService                                               ║
// ║  Serialises and deserialises session state to/from ASP.NET Core session.    ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • ASP.NET Core Session middleware — Microsoft Docs                        ║
// ║   • System.Text.Json — used for in-memory serialisation                    ║
// ║   • Python equivalent: STATE dict in the notebook                           ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using System.Text.Json;
using AfsPdfComparison.Models;
using Microsoft.AspNetCore.Http;

namespace AfsPdfComparison.Services;

// ─────────────────────────────────────────────────────────────────────────────
// SECTION 6 · SESSION STATE SERVICE
//
// Responsibility: provide strongly-typed get/set helpers on top of the raw
// ASP.NET Core ISession byte array store.
//
// Keys:
//   "Engagement"  → EngagementDetails JSON
//   "Reports"     → List<PdfReport> JSON  (pdf paths are server-side only)
//   "Comparison"  → ComparisonResult JSON
//   "Comments"    → Dictionary<string,string> JSON
//
// Note: PdfReport.PdfPath contains a server-side temp path.  If the server
// restarts the temp files are lost; the user must re-upload.  This mirrors
// the Python notebook behaviour where files live in /tmp during the session.
// ─────────────────────────────────────────────────────────────────────────────

/// <summary>
/// Thin typed wrapper over <see cref="ISession"/> for the application's
/// key data objects.  Injected into controllers as a scoped service.
/// </summary>
public class SessionStateService
{
    private readonly IHttpContextAccessor _http;

    private static readonly JsonSerializerOptions _opts = new()
    {
        // Allow tuples (Value Tuple) to be serialised as arrays
        IncludeFields = true,
    };

    private ISession Session => _http.HttpContext!.Session;

    public SessionStateService(IHttpContextAccessor accessor)
    {
        _http = accessor;
    }

    // ── Engagement Details ────────────────────────────────────────────────

    /// <summary>Saves engagement details to session.</summary>
    public void SetEngagement(EngagementDetails eng) =>
        Session.SetString("Engagement", JsonSerializer.Serialize(eng, _opts));

    /// <summary>Retrieves engagement details from session (never null).</summary>
    public EngagementDetails GetEngagement() =>
        Deserialise<EngagementDetails>(Session.GetString("Engagement")) ?? new();

    // ── PDF Reports (extracted documents) ─────────────────────────────────

    /// <summary>Saves the list of extracted PdfReports to session.</summary>
    public void SetReports(List<PdfReport> reports) =>
        Session.SetString("Reports", JsonSerializer.Serialize(reports, _opts));

    /// <summary>Returns extracted PdfReports from session (empty list if none).</summary>
    public List<PdfReport> GetReports() =>
        Deserialise<List<PdfReport>>(Session.GetString("Reports")) ?? new();

    // ── Comparison Result ─────────────────────────────────────────────────

    /// <summary>Saves the full comparison result to session.</summary>
    public void SetComparison(ComparisonResult cmp) =>
        Session.SetString("Comparison", JsonSerializer.Serialize(cmp, _opts));

    /// <summary>Returns the comparison result from session (null if not run yet).</summary>
    public ComparisonResult? GetComparison() =>
        Deserialise<ComparisonResult>(Session.GetString("Comparison"));

    // ── Auditor Comments ──────────────────────────────────────────────────

    /// <summary>Saves the auditor comments dictionary to session.</summary>
    public void SetComments(Dictionary<string, string> comments) =>
        Session.SetString("Comments", JsonSerializer.Serialize(comments, _opts));

    /// <summary>Returns the auditor comments dictionary (empty if none saved).</summary>
    public Dictionary<string, string> GetComments() =>
        Deserialise<Dictionary<string, string>>(Session.GetString("Comments")) ?? new();

    /// <summary>
    /// Saves or updates a single comment by key.
    /// Reads the existing dictionary, upserts the entry, then re-serialises.
    /// </summary>
    public void UpsertComment(string key, string comment)
    {
        var all = GetComments();
        all[key] = comment;
        SetComments(all);
    }

    // ── Helper ────────────────────────────────────────────────────────────
    private static T? Deserialise<T>(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return default;
        try  { return JsonSerializer.Deserialize<T>(json, _opts); }
        catch { return default; }
    }
}
