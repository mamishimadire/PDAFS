
# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  AFS PDF COMPARISON ANALYSER — Launch Script                                ║
# ║  Author  : Mamishi Tonny Madire                                              ║
# ║  Date    : 2026-03-15                                                        ║
# ║  Version : 4.3                                                               ║
# ║                                                                              ║
# ║  Usage: Right-click this file → "Run with PowerShell"                       ║
# ║         OR open PowerShell and run:  .\run web app.ps1                      ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

$projectDir = "$PSScriptRoot\AfsPdfComparison"
$url        = "http://localhost:5075"

Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════╗" -ForegroundColor Magenta
Write-Host "║   SNG Grant Thornton — AFS PDF Comparison Analyser   ║" -ForegroundColor Magenta
Write-Host "║   Author : Mamishi Tonny Madire  |  v4.3  |  2026    ║" -ForegroundColor Magenta
Write-Host "╚═══════════════════════════════════════════════════════╝" -ForegroundColor Magenta
Write-Host ""

# ── Check .NET SDK is available ───────────────────────────────────────────────
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: .NET SDK not found. Download from https://aka.ms/dotnet/download" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# ── Check the project folder exiPowerShell -ExecutionPolicy Bypass -File "C:\Users\Mamishi.Madire\Desktop\pdf comparison\run web app.ps1"sts ──────────────────────────────────────────
if (-not (Test-Path $projectDir)) {
    Write-Host "ERROR: Project folder not found at:" -ForegroundColor Red
    Write-Host "  $projectDir" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# ── Kill any existing instance (releases the port + DLL file lock) ────────────
Write-Host "Stopping any existing instances ..." -ForegroundColor DarkGray
Get-Process -Name "AfsPdfComparison" -ErrorAction SilentlyContinue | Stop-Process -Force
Get-Process -Name "dotnet"           -ErrorAction SilentlyContinue |
    Where-Object { $_.CommandLine -like "*AfsPdfComparison*" } |
    Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 1

# ── Build the project (picks up any code changes) ─────────────────────────────
Write-Host "Building project ..." -ForegroundColor Cyan
Set-Location $projectDir
$buildOutput = dotnet build --no-restore 2>&1
$buildLine   = $buildOutput | Select-String -Pattern "Build succeeded|FAILED"
if ("$buildLine" -match "FAILED") {
    Write-Host ""
    Write-Host "BUILD FAILED - see errors above." -ForegroundColor Red
    $buildOutput | Select-String -Pattern "error CS" | ForEach-Object { Write-Host $_ -ForegroundColor Red }
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "Build OK ✓" -ForegroundColor Green
Write-Host ""

# ── Open browser after a short delay (app needs ~3 s to start) ───────────────
Write-Host "Starting web server ..." -ForegroundColor Cyan
Write-Host "URL : $url" -ForegroundColor Green
Write-Host ""
Write-Host "Press Ctrl+C in this window to stop the server." -ForegroundColor Yellow
Write-Host ""

# Launch browser in background after 4 seconds
Start-Job -ScriptBlock {
    param($u)
    Start-Sleep -Seconds 4
    Start-Process $u
} -ArgumentList $url | Out-Null

# ── Start the ASP.NET Core app ────────────────────────────────────────────────
# Uses the compiled DLL directly — avoids Windows blocking the .exe on Desktop.
$env:ASPNETCORE_ENVIRONMENT = "Development"
$env:ASPNETCORE_URLS        = $url
dotnet "bin\Debug\net8.0\AfsPdfComparison.dll"
