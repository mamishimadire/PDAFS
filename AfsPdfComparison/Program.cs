// ╔══════════════════════════════════════════════════════════════════════════════╗
// ║  AFS PDF COMPARISON ANALYZER  — C# ASP.NET Core Web Application             ║
// ║  SNG Grant Thornton | CAATs Platform                                         ║
// ║                                                                              ║
// ║  Author  : Mamishi Tonny Madire                                              ║
// ║  Date    : 2026-03-15                                                        ║
// ║  Version : 4.3                                                               ║
// ║                                                                              ║
// ║  PROGRAM.CS — Dependency Injection, Middleware, Application Startup          ║
// ║                                                                              ║
// ║  References:                                                                 ║
// ║   • ASP.NET Core startup docs — https://learn.microsoft.com/aspnet/core    ║
// ║   • QuestPDF licence — https://www.questpdf.com/license/                   ║
// ║   • Session middleware — https://learn.microsoft.com/aspnet/core/session   ║
// ╚══════════════════════════════════════════════════════════════════════════════╝

using AfsPdfComparison.Services;
using QuestPDF.Infrastructure;

var builder = WebApplication.CreateBuilder(args);

// SECTION 1 · MVC + RAZOR VIEWS
builder.Services.AddControllersWithViews();

// SECTION 2 · SESSION STATE
// IDistributedMemoryCache backs the in-memory session store.
// IdleTimeout = 2 hours to accommodate large PDF comparison workflows.
// Reference: Python STATE dict — ISession provides the same cross-request state.
builder.Services.AddDistributedMemoryCache();
builder.Services.AddSession(options =>
{
    options.IdleTimeout         = TimeSpan.FromMinutes(30); // security: expire after 30 min idle
    options.Cookie.HttpOnly     = true;
    options.Cookie.IsEssential  = true;
    options.Cookie.SecurePolicy = CookieSecurePolicy.SameAsRequest;
    options.Cookie.SameSite     = SameSiteMode.Strict;     // security: block cross-site access
});

// SECTION 3 · HTTP CONTEXT ACCESSOR
// Required by SessionStateService to access ISession outside of controller scope.
builder.Services.AddHttpContextAccessor();

// SECTION 4 · APPLICATION SERVICES
// Singleton — stateless, thread-safe analysis & extraction services
// Scoped    — SessionStateService requires per-request ISession
//
// Service ↔ Python notebook mapping:
//   PageExtractorService  → extract_pdf(), extract_page_text()
//   PageAlignmentService  → build_alignment(), page_sim()
//   LineComparatorService → compare_lines(), word_diff()
//   ComparisonService     → run_comparison() orchestrator
//   ExportService         → build_pdf_report(), build_word_report(), etc.
//   SessionStateService   → STATE dict wrapper
builder.Services.AddSingleton<PageExtractorService>();
builder.Services.AddSingleton<PageAlignmentService>();
builder.Services.AddSingleton<LineComparatorService>();
builder.Services.AddSingleton<ComparisonService>();
builder.Services.AddSingleton<ExportService>();
builder.Services.AddScoped<SessionStateService>();

// SECTION 4b · NEW SERVICES (v4.3 patch — bugs #1/#2/#3 fixes)
builder.Services.AddScoped<PdfExtractionService>();
builder.Services.AddScoped<LineComparisonService>();
builder.Services.AddScoped<PageSnapshotService>();
builder.Services.AddScoped<WorkingPaperExportService>();

// SECTION 5 · QUESTPDF LICENSE
// Community licence is free for open/internal tools. Must be set before any API call.
// Reference: https://www.questpdf.com/license/
QuestPDF.Settings.License = LicenseType.Community;

var app = builder.Build();

// SECTION 6 · MIDDLEWARE PIPELINE
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();
app.UseSession();       // must come after UseRouting, before MapControllerRoute
app.UseAuthorization();

// SECTION 7 · UPLOADS DIRECTORY — create wwwroot/uploads and WIPE IT at every startup
// Security: ensures no uploaded PDF files from a previous session remain on disk.
// Every run starts completely fresh with no stored user data.
var uploadsDir = Path.Combine(app.Environment.WebRootPath, "uploads");
Directory.CreateDirectory(uploadsDir);
foreach (var f in Directory.GetFiles(uploadsDir))
    try { File.Delete(f); } catch { /* ignore locked files */ }

// SECTION 8 · ROUTING
app.MapControllerRoute(
    name:    "default",
    pattern: "{controller=Afs}/{action=Index}/{id?}");

app.Run();
