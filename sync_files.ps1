$src = Join-Path $env:USERPROFILE "Desktop\MHTAgentic"
$dst = "C:\Users\Public\Desktop\MHTAgentic"

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

foreach ($f in $files) {
    $s = Join-Path $src $f
    $d = Join-Path $dst $f
    if (Test-Path $s) {
        $dir = Split-Path $d -Parent
        if (!(Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Copy-Item $s $d -Force
        Write-Host "OK: $f"
    } else {
        Write-Host "SKIP: $f (not found)"
    }
}
Write-Host "Done"
