# MHT Agentic - RDP Setup Script (Run Elevated)
# 1. Syncs code to Public Desktop MHTAgentic
# 2. Creates scheduled task for auto-start on logon
# 3. Updates the Public Desktop shortcut

$ErrorActionPreference = "Continue"

$src = Join-Path $env:USERPROFILE "Desktop\MHTAgentic"
$dst = "C:\Users\Public\Desktop\MHTAgentic"

Write-Host "=== Step 1: Syncing code to Public Desktop ===" -ForegroundColor Cyan

# Key files to sync
$files = @(
    "launcher.pyw",
    "start_demo.py",
    "start_dashboard.py",
    "mhtagentic\outbound\outbound_worker.py",
    "mhtagentic\desktop\control_overlay.py",
    "mhtagentic\desktop\automation.py",
    "mhtagentic\desktop\analytics.py",
    "mhtagentic\db\database.py",
    "mhtagentic\db\mht_simulator.py",
    "mhtagentic\db\__init__.py",
    "mhtagentic\__init__.py",
    "mhtagentic\outbound\__init__.py",
    "mhtagentic\desktop\__init__.py",
    "dashboard\server.py",
    "dashboard\session_monitor.py",
    "dashboard\screenshot_capture.py",
    "dashboard\__init__.py",
    "dashboard\templates\index.html",
    "dashboard\static\style.css",
    "dashboard\static\dashboard.js",
    "config\.env"
)

$copied = 0
foreach ($f in $files) {
    $srcFile = Join-Path $src $f
    $dstFile = Join-Path $dst $f
    if (Test-Path $srcFile) {
        $dstDir = Split-Path $dstFile -Parent
        if (!(Test-Path $dstDir)) {
            New-Item -ItemType Directory -Path $dstDir -Force | Out-Null
        }
        Copy-Item $srcFile $dstFile -Force
        $copied++
    }
}
Write-Host "  Copied $copied files" -ForegroundColor Green

# Also sync requirements.txt if it exists
$reqSrc = Join-Path $src "requirements.txt"
if (Test-Path $reqSrc) {
    Copy-Item $reqSrc (Join-Path $dst "requirements.txt") -Force
}

# Ensure shared output directories exist in ProgramData (writable by all users)
$outputBase = "C:\ProgramData\MHTAgentic"
$outputDirs = @("", "debug", "analytics", "mht_api", "screenshots")
foreach ($d in $outputDirs) {
    $dirPath = Join-Path $outputBase $d
    if (!(Test-Path $dirPath)) {
        New-Item -ItemType Directory -Path $dirPath -Force | Out-Null
    }
}

Write-Host ""
Write-Host "=== Step 2: Creating scheduled task ===" -ForegroundColor Cyan

$taskName = "MHT_Bot_AutoStart"
$pythonw = (Get-Command pythonw.exe -ErrorAction SilentlyContinue).Source
if (-not $pythonw) { $pythonw = (Get-Command python.exe).Source -replace 'python\.exe$','pythonw.exe' }
$launcher = Join-Path $dst "launcher.pyw"

# Delete existing task if any
schtasks /delete /tn $taskName /f 2>$null

# Create task: runs on logon of ANY user, 5 second delay
schtasks /create /tn $taskName `
    /tr "`"$pythonw`" `"$launcher`"" `
    /sc onlogon `
    /rl HIGHEST `
    /delay 0000:05 `
    /f

if ($LASTEXITCODE -eq 0) {
    Write-Host "  Task '$taskName' created successfully" -ForegroundColor Green
} else {
    Write-Host "  Failed to create task" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== Step 3: Updating Public Desktop shortcut ===" -ForegroundColor Cyan

$lnkPath = "C:\Users\Public\Desktop\MHT Agentic.lnk"
$ws = New-Object -ComObject WScript.Shell
$sc = $ws.CreateShortcut($lnkPath)
$sc.TargetPath = $pythonw
$sc.Arguments = "`"$launcher`""
$sc.WorkingDirectory = $dst
$sc.Save()
Write-Host "  Shortcut updated to point to $dst" -ForegroundColor Green

Write-Host ""
Write-Host "=== Done! ===" -ForegroundColor Green
Write-Host "  - Code synced to $dst"
Write-Host "  - Scheduled task '$taskName' will run launcher.pyw on any user logon"
Write-Host "  - Shortcut updated"
Write-Host ""
Write-Host "Press any key to close..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
