# Multi-RDP Clean Startup

## What It Does

`start_all_clean.py` opens RDP sessions and automatically logs into Experity in each one. No manual interaction needed — just one UAC prompt to approve the elevated helper.

**Tested:** Single RDP login completed in ~50 seconds (full Experity + Okta + TOTP flow).

## Usage

```bash
# All 3 RDPs (ExperityB, ExperityC, ExperityD)
python start_all_clean.py

# Single RDP
python start_all_clean.py ExperityB
```

## Architecture

### The Problem

All 3 RDP sessions connect as the **same Windows user** (ExperityB) to localhost via RDP Wrapper. This creates two challenges:

1. **Session detection**: RDP Wrapper (stascorp/rdpwrap) on Windows 11 Home does NOT properly register sessions in the WTS (Windows Terminal Services) API. `WTSEnumerateSessionsW` returns ghost sessions (state=8, no username) even for active connections.

2. **Process injection**: The bot must run as the **interactive session user** (not SYSTEM) to get UIA access to windows. PsExec `-s -i <sid>` runs as SYSTEM which has limited UIA access.

### The Solution

Uses the same `CreateProcessAsUser` approach as `start_monitoring` (proven to work):

1. **Open RDPs** — `mstsc.exe` opens each `.rdp` file
2. **Elevated helper** — A Python script runs as SYSTEM (via PsExec `-s`) that:
   - **Brute-force session detection**: Tries `WTSQueryUserToken` on session IDs 2-60 (not just WTS-enumerated ones). RDP Wrapper creates sessions with IDs that don't appear in `WTSEnumerateSessionsW`, but `WTSQueryUserToken` still works on them.
   - **CreateProcessAsUser**: Duplicates the session user's token and spawns `clean_bot.py` as the interactive user with full UIA access.
3. **Monitor** — Polls per-RDP status files (`clean_status_RDP1.txt`, etc.) and logs to SQLite.

### Flow Diagram

```
start_all_clean.py (console, jaden)
  |
  ├─ mstsc.exe ExperityB_MHT.rdp  (opens RDP window)
  ├─ mstsc.exe ExperityC_MHT.rdp
  ├─ mstsc.exe ExperityD_MHT.rdp
  |
  ├─ ShellExecuteW("runas") ──► psexec_bat (1 UAC prompt)
  |     |
  |     └─ PsExec -s -d python clean_launch_helper.py  (runs as SYSTEM)
  |           |
  |           ├─ WTSQueryUserToken(sid=2..60) ──► finds live sessions
  |           ├─ CreateProcessAsUserW(session 34) ──► clean_bot_RDP1.bat
  |           ├─ CreateProcessAsUserW(session 35) ──► clean_bot_RDP2.bat
  |           └─ CreateProcessAsUserW(session 36) ──► clean_bot_RDP3.bat
  |
  └─ monitor_status() ──► polls clean_status_RDP1.txt, RDP2, RDP3
```

### Key Files

| File | Location | Purpose |
|------|----------|---------|
| `start_all_clean.py` | Repo root | Main orchestrator (run from console) |
| `clean_bot.py` | Repo root + `C:\MHTAgentic\` | Login bot (runs inside RDP session) |
| `clean_launch_helper.py` | `C:\ProgramData\MHTAgentic\` | Generated elevated helper (CreateProcessAsUser) |
| `clean_bot_RDP1.bat` | `C:\ProgramData\MHTAgentic\` | Generated per-RDP wrapper with env vars |
| `clean_status_RDP1.txt` | `C:\ProgramData\MHTAgentic\` | Status file (polled by monitor) |
| `clean_launch.log` | `C:\ProgramData\MHTAgentic\` | Helper debug log |
| `clean_bot_RDP1_stderr.log` | `C:\ProgramData\MHTAgentic\` | Per-RDP Python stderr |
| `mht_data.db` | `C:\ProgramData\MHTAgentic\` | SQLite DB with `clean_startup_log` table |

## Status Codes

Written to `clean_status_<label>.txt` by clean_bot.py:

| Code | Meaning |
|------|---------|
| 10 | Opening Experity |
| 20 | Entering username (Screen 1) |
| 30 | Clicking Next (Screen 1) |
| 40 | Waiting for Okta sign-in (Screen 2) |
| 50 | Entering username (Screen 2, if not pre-filled) |
| 60 | Entering password (Screen 2) |
| 70 | Clicking Sign In (Screen 2) |
| 80 | Generating OTP code (Screen 3) |
| 90 | Entering OTP code (Screen 3) |
| 100 | Clicking Verify (Screen 3) |
| 110 | Login complete |
| -1 | Error |
| -2 | Skipped (already logged in) |

## SQLite Tracking

Table `clean_startup_log` in `C:\ProgramData\MHTAgentic\mht_data.db`:

```sql
SELECT run_seed, rdp_label, session_id, status_code, status_msg, timestamp, duration_ms
FROM clean_startup_log
ORDER BY id DESC LIMIT 20;
```

## Critical Notes

### clean_bot.py must exist at `C:\MHTAgentic\clean_bot.py`

The bot runs as ExperityB who **cannot read** `C:\Users\jaden\Desktop\MHTAgentic\`. Copy after changes:
```bash
cp clean_bot.py C:\MHTAgentic\clean_bot.py
```

### RDP Wrapper Session Detection

Standard WTS enumeration is broken on this system (Windows 11 Home, termsrv.dll 10.0.26100.6725, rdpwrap.ini from sebaxakerhtc). Sessions show as state=8 with no username. The brute-force `WTSQueryUserToken` probe (IDs 2-60) is the only reliable method.

### Ghost Sessions

Dead sessions (state=8) accumulate and only clear on reboot. They don't interfere with new sessions but make WTS output noisy. If session IDs climb above 60, increase the probe range in `build_helper_script`.

### Old Overlay (`launcher.pyw`)

When reconnecting to a disconnected RDP session, any previously running `pythonw.exe` (old overlay) will still be visible. `clean_bot.py` kills `pythonw.exe` in its session on startup. If overlay persists, manually kill:
```bash
PsExec64.exe -s taskkill /F /IM pythonw.exe
```

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| "No live sessions" in helper log | RDP didn't connect. Check port 3389 is listening. Check rdpwrap.dll is loaded. |
| Permission denied on clean_bot.py | Copy to `C:\MHTAgentic\clean_bot.py` |
| Bot timeout | Check `clean_bot_RDP1_stderr.log` for Python errors |
| Old overlay visible | Kill pythonw: `PsExec64.exe -s taskkill /F /IM pythonw.exe` |
| Session IDs > 60 | Reboot to clear ghosts, or increase probe range |
| UAC prompt not appearing | PsExec64.exe missing from `C:\ProgramData\MHTAgentic\` |
