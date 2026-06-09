<#
.SYNOPSIS
  Starts the GOV4Agents app locally (server + web in parallel).

.DESCRIPTION
  This script:
  1. Verifies Node.js is installed
  2. Verifies the required env files exist (server/.env, web/.env.local)
  3. Installs npm dependencies if node_modules is missing
  4. Starts both the backend (port 8787) and the web UI (port 5173) via `npm run dev`

.EXAMPLE
  .\scripts\Start-App.ps1
#>

param(
    [switch]$SkipInstall
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Resolve repo root (parent of this script's folder)
$repoRoot = Split-Path -Parent $PSScriptRoot

Write-Host "`n=== GOV4Agents Local Startup ===" -ForegroundColor Cyan
Write-Host "Repo root: $repoRoot`n" -ForegroundColor DarkGray

# --- 1. Verify Node.js ---
try {
    $nodeVersion = node --version
    Write-Host "Node.js: $nodeVersion" -ForegroundColor Green
} catch {
    Write-Host "Node.js is not installed or not on PATH. Install Node 20+ first." -ForegroundColor Red
    exit 1
}

# --- 2. Verify env files ---
$serverEnv = Join-Path $repoRoot "server\.env"
$webEnv = Join-Path $repoRoot "web\.env.local"
$missing = $false

if (-not (Test-Path $serverEnv)) {
    Write-Host "Missing: server\.env" -ForegroundColor Red
    Write-Host "  Create it with TENANT_ID, CLIENT_ID, CLIENT_SECRET, SPA_ORIGIN." -ForegroundColor Yellow
    $missing = $true
} else {
    $serverContent = Get-Content $serverEnv -Raw
    foreach ($key in @("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET")) {
        if ($serverContent -notmatch "(?m)^\s*$key\s*=\s*\S") {
            Write-Host "server\.env is missing a value for $key" -ForegroundColor Red
            $missing = $true
        }
    }
    if (-not $missing) { Write-Host "server\.env: OK" -ForegroundColor Green }
}

if (-not (Test-Path $webEnv)) {
    Write-Host "Missing: web\.env.local" -ForegroundColor Red
    Write-Host "  Create it with VITE_CLIENT_ID and VITE_AUTHORITY." -ForegroundColor Yellow
    $missing = $true
} else {
    $webContent = Get-Content $webEnv -Raw
    if ($webContent -notmatch "(?m)^\s*VITE_CLIENT_ID\s*=\s*\S") {
        Write-Host "web\.env.local is missing VITE_CLIENT_ID" -ForegroundColor Red
        $missing = $true
    } else {
        Write-Host "web\.env.local: OK" -ForegroundColor Green
    }
}

if ($missing) {
    Write-Host "`nFix the env files above, then re-run this script." -ForegroundColor Red
    Write-Host "Tip: run .\scripts\Register-App.ps1 to provision the app registration." -ForegroundColor Yellow
    exit 1
}

# --- 3. Install dependencies if needed ---
if (-not $SkipInstall) {
    if (-not (Test-Path (Join-Path $repoRoot "node_modules"))) {
        Write-Host "`nInstalling npm dependencies (this may take a minute)..." -ForegroundColor Yellow
        Push-Location $repoRoot
        try {
            npm install
        } finally {
            Pop-Location
        }
    } else {
        Write-Host "Dependencies already installed (use without -SkipInstall to force check)." -ForegroundColor DarkGray
    }
}

# --- 4. Start the app (server first, then web once the API port is ready) ---
Write-Host "`nStarting GOV4Agents..." -ForegroundColor Cyan

$serverPort = 8787

# Start the backend in a separate process so we can wait for it to bind.
Write-Host "  Starting server (port $serverPort)..." -ForegroundColor Yellow
$serverProc = Start-Process -FilePath "cmd.exe" `
    -ArgumentList "/c", "npm run dev:server" `
    -WorkingDirectory $repoRoot `
    -NoNewWindow -PassThru

# Poll until the API port accepts TCP connections (max ~30s).
$ready = $false
for ($i = 0; $i -lt 60; $i++) {
    if ($serverProc.HasExited) {
        Write-Host "Server process exited unexpectedly (code $($serverProc.ExitCode))." -ForegroundColor Red
        exit 1
    }
    try {
        $client = $null
        $client = New-Object System.Net.Sockets.TcpClient
        $client.Connect("localhost", $serverPort)
        if ($client.Connected) { $ready = $true; $client.Close(); break }
    } catch {
        # not up yet
    } finally {
        if ($client) { $client.Dispose() }
    }
    Start-Sleep -Milliseconds 500
}

if (-not $ready) {
    Write-Host "Server did not become ready on port $serverPort within 30s." -ForegroundColor Red
    if (-not $serverProc.HasExited) { Stop-Process -Id $serverProc.Id -Force }
    exit 1
}

Write-Host "  Server ready: http://localhost:$serverPort" -ForegroundColor Green
Write-Host "  Starting web UI: http://localhost:5173" -ForegroundColor Green
Write-Host "`nPress Ctrl+C to stop (both server and web will be stopped).`n" -ForegroundColor DarkGray

Push-Location $repoRoot
try {
    npm run dev:web
} finally {
    Pop-Location
    if ($serverProc -and -not $serverProc.HasExited) {
        Write-Host "`nStopping server (PID $($serverProc.Id))..." -ForegroundColor Yellow
        # Stop child processes (cmd.exe spawns npm -> node) by PID, then the parent.
        $children = Get-CimInstance Win32_Process -Filter "ParentProcessId = $($serverProc.Id)" -ErrorAction SilentlyContinue
        foreach ($child in $children) {
            Stop-Process -Id $child.ProcessId -Force -ErrorAction SilentlyContinue
        }
        Stop-Process -Id $serverProc.Id -Force -ErrorAction SilentlyContinue
    }
}

