# ==============================================================================
#  AFS PDF COMPARISON ANALYSER -- Launch Script
#  Author  : Mamishi Tonny Madire
#  Version : 4.4
#  Usage   : Right-click -> "Run with PowerShell"  OR  run from any terminal
# ==============================================================================

# Capture script folder
if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
} else {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

$projectDir = Join-Path $scriptDir "AfsPdfComparison"
$url        = "http://localhost:5075"

Write-Host ""
Write-Host "+-------------------------------------------------------+" -ForegroundColor Magenta
Write-Host "|   SNG Grant Thornton -- AFS PDF Comparison Analyser  |" -ForegroundColor Magenta
Write-Host "|   Author : Mamishi Tonny Madire  |  v4.4  |  2026    |" -ForegroundColor Magenta
Write-Host "+-------------------------------------------------------+" -ForegroundColor Magenta
Write-Host ""

# Check .NET SDK
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    Write-Host "ERROR: .NET SDK not found. Download from https://aka.ms/dotnet/download" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check project folder
if (-not (Test-Path $projectDir)) {
    Write-Host "ERROR: Project folder not found at $projectDir" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Load .env config
$dotenvPath = $null
foreach ($c in @($env:AFS_DOTENV_PATH, $env:DOTENV_PATH, (Join-Path $projectDir ".env"), (Join-Path $scriptDir ".env"))) {
    if ($c -and (Test-Path $c)) { $dotenvPath = $c; break }
}
if ($dotenvPath) {
    $env:DOTENV_PATH = (Resolve-Path $dotenvPath).Path
    Write-Host "Config  : $($env:DOTENV_PATH)" -ForegroundColor DarkGray
} else {
    Write-Host "Config  : no .env file found -- using system environment variables" -ForegroundColor Yellow
}

# Try Defender exclusion (silently skipped if not admin -- not required)
try {
    Add-MpPreference -ExclusionPath (Join-Path $projectDir "bin") -ErrorAction Stop
    Add-MpPreference -ExclusionPath (Join-Path $projectDir "obj") -ErrorAction Stop
    Write-Host "Defender: exclusions added for bin\ and obj\" -ForegroundColor DarkGray
} catch {
    # Not admin or Defender not active -- harmless, build will retry on lock
}

# Kill any running instance
Write-Host "Stopping any existing instances ..." -ForegroundColor DarkGray
Get-Process -Name "AfsPdfComparison" -ErrorAction SilentlyContinue | Stop-Process -Force
$dotnetProcs = Get-CimInstance Win32_Process -Filter "Name = 'dotnet.exe'" -ErrorAction SilentlyContinue
foreach ($p in $dotnetProcs) {
    if ($p.CommandLine -like "*AfsPdfComparison*") {
        try { Stop-Process -Id $p.ProcessId -Force } catch {}
    }
}
Start-Sleep -Seconds 2

# Build
Write-Host "Building project ..." -ForegroundColor Cyan
Set-Location $projectDir

$buildOutput = dotnet build --no-restore 2>&1
$buildOk     = $buildOutput | Select-String "Build succeeded"

if (-not $buildOk) {
    $lockHit = $buildOutput | Select-String "CS2012|MSB3027|MSB3021"
    if ($lockHit) {
        Write-Host "File lock detected -- waiting 4 s and retrying ..." -ForegroundColor Yellow
        Start-Sleep -Seconds 4
        $buildOutput = dotnet build --no-restore 2>&1
        $buildOk     = $buildOutput | Select-String "Build succeeded"
    }
}

if (-not $buildOk) {
    Write-Host ""
    Write-Host "BUILD FAILED" -ForegroundColor Red
    $buildOutput | Select-String "error CS|error MSB" | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Build OK" -ForegroundColor Green
Write-Host ""
Write-Host "Starting web server at $url ..." -ForegroundColor Cyan
Write-Host "Press Ctrl+C in this window to stop the server." -ForegroundColor Yellow
Write-Host ""

# Open browser after 6 s (background process -- no admin needed)
$browserCmd = "Start-Sleep -Seconds 6; Start-Process '" + $url + "'"
Start-Process powershell.exe -ArgumentList @("-NoProfile", "-WindowStyle", "Hidden", "-Command", $browserCmd) -WindowStyle Hidden

# Launch app
$env:ASPNETCORE_ENVIRONMENT = "Development"
$env:ASPNETCORE_URLS        = $url
dotnet "bin\Debug\net8.0\AfsPdfComparison.dll"