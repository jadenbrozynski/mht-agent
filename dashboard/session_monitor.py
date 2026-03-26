"""
RDP session detection and bot health checks.

Finds mstsc.exe (Remote Desktop) windows, tracks their state,
and infers bot health from SQLite database freshness.
"""

import ctypes
import ctypes.wintypes
import logging
import os
import sqlite3
import subprocess
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional

from mhtagentic.db import release_slot

logger = logging.getLogger("mht_dashboard.monitor")


def _find_system_python(exe_name="pythonw.exe"):
    """Find Python exe dynamically — prefers system-wide install."""
    # Check Program Files first — accessible by all users
    for base in [r"C:\Program Files", r"C:\Program Files (x86)"]:
        for d in sorted(Path(base).glob("Python*"), reverse=True):
            c = d / exe_name
            if c.exists():
                return c
    import shutil, sys
    candidate = Path(sys.executable).parent / exe_name
    if candidate.exists():
        return candidate
    found = shutil.which(exe_name)
    if found:
        return Path(found)
    return Path(exe_name)


# OTP signal directory — shared across sessions via ProgramData
_OTP_SIGNAL_DIR = Path(r"C:\ProgramData\MHTAgentic\session_status")

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
advapi32 = ctypes.windll.advapi32

# Process access rights
PROCESS_TERMINATE = 0x0001
PROCESS_QUERY_INFORMATION = 0x0400

# Privilege constants
TOKEN_ADJUST_PRIVILEGES = 0x0020
TOKEN_QUERY = 0x0008
SE_PRIVILEGE_ENABLED = 0x00000002


class _LUID(ctypes.Structure):
    _fields_ = [("LowPart", ctypes.wintypes.DWORD), ("HighPart", ctypes.wintypes.LONG)]


class _LUID_AND_ATTRIBUTES(ctypes.Structure):
    _fields_ = [("Luid", _LUID), ("Attributes", ctypes.wintypes.DWORD)]


class _TOKEN_PRIVILEGES(ctypes.Structure):
    _fields_ = [
        ("PrivilegeCount", ctypes.wintypes.DWORD),
        ("Privileges", _LUID_AND_ATTRIBUTES * 1),
    ]


def _enable_debug_privilege():
    """Enable SeDebugPrivilege so we can kill processes owned by other users."""
    token = ctypes.wintypes.HANDLE()
    if not advapi32.OpenProcessToken(
        kernel32.GetCurrentProcess(),
        TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY,
        ctypes.byref(token),
    ):
        return False

    luid = _LUID()
    if not advapi32.LookupPrivilegeValueW(None, "SeDebugPrivilege", ctypes.byref(luid)):
        kernel32.CloseHandle(token)
        return False

    tp = _TOKEN_PRIVILEGES()
    tp.PrivilegeCount = 1
    tp.Privileges[0].Luid = luid
    tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED

    result = advapi32.AdjustTokenPrivileges(token, False, ctypes.byref(tp), 0, None, None)
    kernel32.CloseHandle(token)
    return bool(result)


def _kill_process(pid: int) -> bool:
    """Terminate a process by PID using the Win32 API."""
    handle = kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
    if not handle:
        return False
    result = kernel32.TerminateProcess(handle, 1)
    kernel32.CloseHandle(handle)
    return bool(result)


def _run_elevated_cleanup(pids: List[int], session_ids: List[int]) -> bool:
    """
    Kill PIDs and log off sessions by running an elevated Python script.
    Triggers one UAC prompt. The script:
    1. Enables SeDebugPrivilege
    2. Terminates all target PIDs
    3. Calls WTSLogoffSession for each RDP session
    """
    import sys
    import tempfile

    # Write a self-contained cleanup script
    script = f"""
import ctypes, ctypes.wintypes, sys, time

kernel32 = ctypes.windll.kernel32
advapi32 = ctypes.windll.advapi32
wts = ctypes.windll.Wtsapi32

# Enable SeDebugPrivilege
class LUID(ctypes.Structure):
    _fields_ = [("Lo", ctypes.wintypes.DWORD), ("Hi", ctypes.wintypes.LONG)]
class LA(ctypes.Structure):
    _fields_ = [("Luid", LUID), ("Attr", ctypes.wintypes.DWORD)]
class TP(ctypes.Structure):
    _fields_ = [("Count", ctypes.wintypes.DWORD), ("Privs", LA * 1)]

tok = ctypes.wintypes.HANDLE()
advapi32.OpenProcessToken(kernel32.GetCurrentProcess(), 0x0028, ctypes.byref(tok))
luid = LUID()
advapi32.LookupPrivilegeValueW(None, "SeDebugPrivilege", ctypes.byref(luid))
tp = TP()
tp.Count = 1
tp.Privs[0].Luid = luid
tp.Privs[0].Attr = 2
advapi32.AdjustTokenPrivileges(tok, False, ctypes.byref(tp), 0, None, None)
kernel32.CloseHandle(tok)

# Kill processes
pids = {pids!r}
killed = 0
for pid in pids:
    h = kernel32.OpenProcess(0x0001, False, pid)
    if h:
        if kernel32.TerminateProcess(h, 1):
            killed += 1
        kernel32.CloseHandle(h)

# Wait briefly for processes to die
time.sleep(1)

# Log off sessions
session_ids = {session_ids!r}
for sid in session_ids:
    wts.WTSLogoffSession(0, sid, True)

sys.exit(0)
"""
    # Write to temp file
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".py", prefix="mht_cleanup_", delete=False
    )
    tmp.write(script)
    tmp.close()

    shell32 = ctypes.windll.shell32
    python_exe = sys.executable

    # Run elevated: pythonw.exe temp_script.py (hidden window)
    result = shell32.ShellExecuteW(
        None, "runas", python_exe, f'"{tmp.name}"', None, 0
    )
    ok = result > 32
    if ok:
        logger.info(
            f"Elevated cleanup launched: {len(pids)} PIDs, "
            f"{len(session_ids)} sessions to logoff"
        )
    else:
        logger.error(f"ShellExecuteW(runas) failed with code {result}")

    # Clean up temp file after a delay (in background)
    import threading
    def _cleanup():
        import time as _t, os
        _t.sleep(10)
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
    threading.Thread(target=_cleanup, daemon=True).start()

    return ok


# Enable DPI awareness (must match screenshot_capture.py)
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    user32.SetProcessDPIAware()

# Try to enable debug privilege (works if running as admin)
_enable_debug_privilege()

# Callback type for EnumWindows
WNDENUMPROC = ctypes.WINFUNCTYPE(
    ctypes.wintypes.BOOL,
    ctypes.wintypes.HWND,
    ctypes.wintypes.LPARAM,
)


def _get_window_text(hwnd: int) -> str:
    """Get window title text."""
    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def _get_class_name(hwnd: int) -> str:
    """Get window class name."""
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def wait_for_otp_complete(username: str, timeout: int = 180) -> bool:
    """
    Poll for bot-ready signals from a given RDP username.

    Checks TWO sources (whichever fires first = success):
    1. OTP signal file: C:\\ProgramData\\MHTAgentic\\session_status\\{username}_otp_complete
    2. Bot slot in DB: bot_slot row for the user shows status='active'

    Args:
        username: The RDP session username (e.g. "ExperityB")
        timeout: Max seconds to wait

    Returns:
        True if either signal appeared before timeout, False otherwise
    """
    import time as _time
    from mhtagentic import OUTPUT_DIR
    from mhtagentic.db import get_slots

    signal_path = _OTP_SIGNAL_DIR / f"{username}_otp_complete"
    slot_name = username.lower()
    db_path = str(OUTPUT_DIR / "mht_data.db")
    deadline = _time.time() + timeout
    logger.info(f"Waiting for OTP signal or slot active: {username} (timeout={timeout}s)")

    while _time.time() < deadline:
        # Check 1: OTP signal file
        if signal_path.exists():
            logger.info(f"OTP signal FILE found for {username}")
            return True

        # Check 2: Bot slot claimed in DB (bot is alive and running)
        try:
            slots = get_slots(db_path)
            for s in slots:
                if s["slot_name"] == slot_name and s["status"] == "active":
                    logger.info(
                        f"Bot slot '{slot_name}' is ACTIVE in DB "
                        f"(PID {s.get('session_id')}) — treating as success"
                    )
                    return True
        except Exception:
            pass

        _time.sleep(3)

    logger.warning(f"OTP signal timeout for {username} after {timeout}s (no file, no active slot)")
    return False


def clear_all_otp_signals():
    """Remove all OTP signal files so the next start-all cycle works cleanly."""
    if not _OTP_SIGNAL_DIR.exists():
        return
    for f in _OTP_SIGNAL_DIR.glob("*_otp_complete"):
        try:
            f.unlink()
            logger.info(f"Cleared OTP signal: {f}")
        except Exception:
            pass


def find_rdp_windows() -> List[Dict]:
    """
    Find all visible mstsc (Remote Desktop) windows.

    Returns:
        List of dicts with keys: hwnd, title, left, top, width, height
    """
    windows = []

    def enum_callback(hwnd, _lparam):
        if not user32.IsWindowVisible(hwnd):
            return True

        class_name = _get_class_name(hwnd)
        title = _get_window_text(hwnd)

        # mstsc windows have class "TscShellContainerClass" or title contains RDP indicators
        is_rdp = (
            class_name == "TscShellContainerClass"
            or "remote desktop" in title.lower()
            or "mstsc" in title.lower()
        )

        if is_rdp and title:
            rect = ctypes.wintypes.RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            windows.append({
                "hwnd": hwnd,
                "title": title,
                "left": rect.left,
                "top": rect.top,
                "width": rect.right - rect.left,
                "height": rect.bottom - rect.top,
            })

        return True

    user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    return windows


def find_rdp_files(search_dirs: List[Path]) -> List[Dict]:
    """
    Scan directories for .rdp files.

    Args:
        search_dirs: List of directories to scan

    Returns:
        List of dicts with keys: name, path
    """
    rdp_files = []
    seen = set()
    for d in search_dirs:
        if not d.exists():
            continue
        for f in d.glob("*.rdp"):
            if f.name not in seen:
                seen.add(f.name)
                rdp_files.append({
                    "name": f.stem,
                    "path": str(f),
                })
    return rdp_files


TASK_NAME = "MHT_Bot_AutoStart"
_BOT_USERS = ["ExperityB", "ExperityC", "ExperityD"]


def _parse_rdp_username(rdp_path: str) -> Optional[str]:
    """Extract 'username' from an .rdp file."""
    try:
        with open(rdp_path, "r") as f:
            for line in f:
                if line.lower().startswith("username:s:"):
                    return line.split(":", 2)[2].strip()
    except Exception:
        pass
    return None


def ensure_autostart_task(project_root: Path) -> Dict:
    """
    Ensure per-user MHT_Bot_AutoStart scheduled tasks exist for each bot user.

    Checks for existing tasks first, then creates ALL missing tasks via a
    single elevated PowerShell script (one UAC prompt).
    """
    import sys
    import tempfile

    # Use C:\MHTAgentic as the shared path accessible by all users
    shared_root = Path(r"C:\MHTAgentic")
    launcher_path = shared_root / "launcher.pyw"
    if not launcher_path.exists():
        launcher_path = project_root / "launcher.pyw"
    if not launcher_path.exists():
        return {"success": False, "error": "launcher.pyw not found"}

    # Find pythonw.exe with full path
    python_dir = Path(sys.executable).parent
    pythonw_path = python_dir / "pythonw.exe"
    if not pythonw_path.exists():
        pythonw_path = _find_system_python("pythonw.exe")

    results = {}
    missing_users = []

    for username in _BOT_USERS:
        task_name = f"{TASK_NAME}_{username}"

        # Check if this user's task already exists
        check = subprocess.run(
            ["schtasks", "/query", "/tn", task_name],
            capture_output=True, text=True, timeout=5,
        )
        if check.returncode == 0:
            logger.info(f"Scheduled task '{task_name}' already exists")
            results[username] = {"exists": True}
        else:
            logger.warning(f"Scheduled task '{task_name}' NOT found — needs creation")
            results[username] = {"exists": False}
            missing_users.append(username)

    if not missing_users:
        logger.info("All scheduled tasks already exist — no action needed")
        return {"success": True, "tasks": results}

    # Build a single PowerShell script that creates all missing tasks (one UAC prompt)
    ps_lines = [
        "# MHT Auto-Start Task Creator (auto-generated)",
        "$ErrorActionPreference = 'Continue'",
        f'$pythonw = "{pythonw_path}"',
        f'$launcher = "{launcher_path}"',
        "",
    ]

    for username in missing_users:
        task_name = f"{TASK_NAME}_{username}"
        computer_name = "$env:COMPUTERNAME"
        ps_lines.extend([
            f'# --- {username} ---',
            f'Unregister-ScheduledTask -TaskName "{task_name}" -Confirm:$false -ErrorAction SilentlyContinue',
            f'$action = New-ScheduledTaskAction -Execute $pythonw -Argument "`"$launcher`"" -WorkingDirectory (Split-Path $launcher)',
            f'$trigger = New-ScheduledTaskTrigger -AtLogOn',
            f'$trigger.UserId = "{computer_name}\\{username}"',
            f'$trigger.Delay = "PT5S"',
            f'$principal = New-ScheduledTaskPrincipal -UserId "{computer_name}\\{username}" -LogonType Interactive -RunLevel Highest',
            f'$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 12)',
            f'Register-ScheduledTask -TaskName "{task_name}" -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force',
            f'Write-Host "Created task: {task_name}"',
            "",
        ])

    # Also remove old global task
    ps_lines.extend([
        f'Unregister-ScheduledTask -TaskName "{TASK_NAME}" -Confirm:$false -ErrorAction SilentlyContinue',
    ])

    # Write the PowerShell script to a temp file
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".ps1", prefix="mht_create_tasks_", delete=False
    )
    tmp.write("\r\n".join(ps_lines))
    tmp.close()

    logger.info(f"Creating {len(missing_users)} scheduled tasks via elevated PowerShell: {tmp.name}")

    # Run elevated (single UAC prompt)
    shell32 = ctypes.windll.shell32
    result = shell32.ShellExecuteW(
        None, "runas", "powershell.exe",
        f'-ExecutionPolicy Bypass -File "{tmp.name}"',
        None, 0  # SW_HIDE
    )
    ok = result > 32

    if ok:
        logger.info(f"Elevated PowerShell launched to create tasks for: {missing_users}")
    else:
        logger.error(f"Failed to launch elevated PowerShell (code {result})")

    for username in missing_users:
        results[username]["created"] = ok

    # Clean up temp file after a delay
    import threading as _thr
    def _cleanup():
        import time as _t, os
        _t.sleep(15)
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
    _thr.Thread(target=_cleanup, daemon=True).start()

    return {"success": ok, "tasks": results, "created": missing_users}


def _kill_bot_processes_in_session(session_id: int) -> int:
    """Kill all python.exe processes in a specific RDP session via wmic."""
    killed = 0
    try:
        r = subprocess.run(
            ['tasklist', '/V', '/FI', 'IMAGENAME eq python.exe', '/FO', 'CSV'],
            capture_output=True, text=True, timeout=10,
        )
        for line in r.stdout.strip().split('\n')[1:]:
            parts = line.replace('"', '').split(',')
            if len(parts) >= 4:
                try:
                    pid = int(parts[1])
                    sess = parts[2].strip()
                    sess_num = int(parts[3])
                except (ValueError, IndexError):
                    continue
                if sess_num == session_id:
                    subprocess.run(
                        ['wmic', 'process', 'where', f'ProcessId={pid}', 'call', 'terminate'],
                        capture_output=True, text=True, timeout=10,
                    )
                    logger.info(f"Killed python.exe PID {pid} in session {session_id}")
                    killed += 1
    except Exception as e:
        logger.error(f"Failed to kill processes in session {session_id}: {e}")
    return killed


def _kill_all_rdp_bot_processes() -> int:
    """Kill all python.exe processes running in RDP sessions using PsExec as SYSTEM."""
    killed = 0

    # Use PsExec to run taskkill as SYSTEM — can kill processes in other user sessions
    try:
        psexec = str(_PSEXEC_PATH)
        if not _PSEXEC_PATH.exists():
            logger.error(f"PsExec not found at {_PSEXEC_PATH}")
            return 0

        # Get our own PID so we don't kill the dashboard
        my_pid = os.getpid()
        # List python processes, then kill ones that aren't us
        r = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq python.exe', '/FO', 'CSV'],
            capture_output=True, text=True, timeout=10,
        )
        for line in r.stdout.strip().split('\n')[1:]:
            parts = line.replace('"', '').split(',')
            if len(parts) >= 2:
                try:
                    pid = int(parts[1])
                except (ValueError, IndexError):
                    continue
                if pid == my_pid:
                    continue
                subprocess.run(
                    [psexec, '-accepteula', '-nobanner', '-s',
                     'taskkill', '/F', '/PID', str(pid)],
                    capture_output=True, text=True, timeout=10,
                )
                logger.info(f"PsExec killed python.exe PID {pid}")
                killed += 1
    except Exception as e:
        logger.error(f"PsExec taskkill failed: {e}")

    # NOTE: mstsc windows are closed AFTER agent sessions are logged off (in stop_all)

    return killed


def _logoff_sessions_elevated(session_ids: List[int]) -> bool:
    """Log off sessions via elevated PowerShell script (proven approach)."""
    if not session_ids:
        return True

    script_path = Path(r"C:\ProgramData\MHTAgentic\_logoff_sessions.ps1")
    ps_code = (
        "Add-Type -TypeDefinition @'\n"
        "using System;\n"
        "using System.Runtime.InteropServices;\n"
        "public class WTS {\n"
        '    [DllImport("wtsapi32.dll", SetLastError=true)]\n'
        "    public static extern bool WTSLogoffSession(IntPtr hServer, int sessionId, bool bWait);\n"
        "}\n"
        "'@\n"
    )
    for sid in session_ids:
        ps_code += f"[WTS]::WTSLogoffSession([IntPtr]::Zero, {sid}, $true) | Out-Null\n"
    # Kill mstsc after logoff to prevent "session ended" error dialogs
    ps_code += "Start-Sleep -Seconds 2\n"
    ps_code += "Get-Process mstsc -ErrorAction SilentlyContinue | Stop-Process -Force\n"

    script_path.write_text(ps_code)

    try:
        r = subprocess.run(
            ['powershell.exe', '-Command',
             f"Start-Process -Verb RunAs -Wait -FilePath 'powershell.exe' "
             f"-ArgumentList '-ExecutionPolicy','Bypass','-File','{script_path}'"],
            capture_output=True, text=True, timeout=30,
        )
        logger.info(f"Elevated logoff complete for sessions {session_ids}")
        return True
    except Exception as e:
        logger.error(f"Elevated logoff failed: {e}")
        return False


def _logoff_rdp_user(username: str) -> bool:
    """Logoff any existing session for a user so next connect is a fresh logon.
    Kills bot processes first, then logs off via elevated WTSLogoffSession."""
    rdp_sessions = _find_rdp_session_ids()
    target_sessions = [
        s for s in rdp_sessions
        if s["username"].lower() == username.lower()
    ]
    if not target_sessions:
        logger.info(f"No existing session found for {username}")
        return False

    session_ids = [s["session_id"] for s in target_sessions]
    logger.info(f"Found sessions for {username}: {session_ids}")

    # Kill bot processes in target sessions first
    for sid in session_ids:
        _kill_bot_processes_in_session(sid)

    import time as _time
    _time.sleep(1)

    # Log off sessions and kill mstsc
    return _logoff_sessions_elevated(session_ids)


def start_rdp_session(rdp_path: str, project_root: Optional[Path] = None) -> Dict:
    """
    Launch mstsc.exe with the given .rdp file.
    The bot auto-starts inside the RDP session via the MHT_Bot_AutoStart
    scheduled task (triggers on logon).
    """
    rdp = Path(rdp_path)
    if not rdp.exists() or rdp.suffix.lower() != ".rdp":
        return {"success": False, "error": "Invalid .rdp file path"}

    try:
        proc = subprocess.Popen(
            ["mstsc.exe", str(rdp)],
            creationflags=subprocess.DETACHED_PROCESS,
        )
        logger.info(f"Started mstsc.exe with {rdp.name} (pid {proc.pid})")

        return {
            "success": True,
            "pid": proc.pid,
            "file": rdp.name,
        }
    except Exception as e:
        logger.error(f"Failed to start RDP: {e}")
        return {"success": False, "error": str(e)}


def stop_rdp_session(hwnd: int) -> Dict:
    """
    Close an mstsc window by its HWND (sends WM_CLOSE).

    Args:
        hwnd: Window handle

    Returns:
        Dict with success status
    """
    WM_CLOSE = 0x0010
    try:
        result = user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
        if result:
            logger.info(f"Sent WM_CLOSE to hwnd {hwnd}")
            return {"success": True}
        else:
            return {"success": False, "error": "PostMessage failed"}
    except Exception as e:
        logger.error(f"Failed to close hwnd {hwnd}: {e}")
        return {"success": False, "error": str(e)}


def _get_pid_from_hwnd(hwnd: int) -> Optional[int]:
    """Get the process ID that owns a window handle."""
    pid = ctypes.wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value if pid.value else None


# --- WTS (Windows Terminal Services) API for session management ---

wts = ctypes.windll.Wtsapi32
WTS_CURRENT_SERVER_HANDLE = 0
WTSUserName = 5


class _WTS_SESSION_INFO(ctypes.Structure):
    _fields_ = [
        ("SessionId", ctypes.wintypes.DWORD),
        ("pWinStationName", ctypes.wintypes.LPWSTR),
        ("State", ctypes.wintypes.DWORD),
    ]


def _find_rdp_session_ids() -> List[Dict]:
    """
    Enumerate Windows sessions and return non-console, non-services sessions
    (i.e. RDP sessions) with their usernames.
    """
    sessions = []
    pInfo = ctypes.POINTER(_WTS_SESSION_INFO)()
    count = ctypes.wintypes.DWORD()

    if not wts.WTSEnumerateSessionsW(
        WTS_CURRENT_SERVER_HANDLE, 0, 1,
        ctypes.byref(pInfo), ctypes.byref(count),
    ):
        return sessions

    for i in range(count.value):
        s = pInfo[i]
        station = s.pWinStationName or ""

        # Query username for this session
        buf = ctypes.wintypes.LPWSTR()
        size = ctypes.wintypes.DWORD()
        wts.WTSQuerySessionInformationW(
            WTS_CURRENT_SERVER_HANDLE, s.SessionId, WTSUserName,
            ctypes.byref(buf), ctypes.byref(size),
        )
        username = buf.value if buf.value else ""

        # Skip sessions with no user (services, listeners)
        if not username:
            continue

        # Skip the console session (that's the local desktop, not RDP)
        if station.lower() == "console":
            continue

        # This is an RDP session (active, disconnected, or idle)
        sessions.append({
            "session_id": s.SessionId,
            "station": station,
            "username": username,
            "state": s.State,
        })

    wts.WTSFreeMemory(pInfo)
    return sessions


def _logoff_session(session_id: int) -> bool:
    """Logoff a Windows session by ID. Kills all processes in that session."""
    result = wts.WTSLogoffSession(WTS_CURRENT_SERVER_HANDLE, session_id, True)
    return bool(result)


class _WTS_PROCESS_INFO(ctypes.Structure):
    _fields_ = [
        ("SessionId", ctypes.wintypes.DWORD),
        ("ProcessId", ctypes.wintypes.DWORD),
        ("pProcessName", ctypes.wintypes.LPWSTR),
        ("pUserSid", ctypes.c_void_p),
    ]


# System processes that should NOT be killed (they'll crash the session manager)
_SYSTEM_PROCS = frozenset({
    "csrss.exe", "winlogon.exe", "fontdrvhost.exe", "dwm.exe",
    "logonui.exe", "smss.exe", "wininit.exe", "services.exe", "lsass.exe",
})


def stop_all_rdp_sessions(rdp_search_dirs: Optional[List[Path]] = None) -> Dict:
    """
    Stop bot processes and close agent RDP windows.
    NEVER logoff sessions or kill the orchestrator RDP.
    """
    import time as _time

    # Clean up OTP signals
    clear_all_otp_signals()

    # Step 1: Kill bot processes (skips dashboard, uses PsExec for agent sessions)
    logger.info("[Stop All] Step 1: Killing bot processes...")
    killed = _kill_all_rdp_bot_processes()
    _time.sleep(1)

    # Step 2: Logoff agent sessions ONLY — use WTS API to get exact usernames
    logger.info("[Stop All] Step 2: Logging off agent sessions...")
    logged_off = 0
    try:
        import ctypes, ctypes.wintypes
        wtsapi32 = ctypes.windll.Wtsapi32

        class WTS_SESSION_INFO(ctypes.Structure):
            _fields_ = [
                ("SessionId", ctypes.wintypes.DWORD),
                ("pWinStationName", ctypes.wintypes.LPWSTR),
                ("State", ctypes.wintypes.DWORD),
            ]

        pInfo = ctypes.POINTER(WTS_SESSION_INFO)()
        count = ctypes.wintypes.DWORD()
        wtsapi32.WTSEnumerateSessionsW(0, 0, 1, ctypes.byref(pInfo), ctypes.byref(count))

        for i in range(count.value):
            sid = pInfo[i].SessionId
            if sid in (0, 1, 65536):
                continue
            # Get username for this session
            buf = ctypes.wintypes.LPWSTR()
            size = ctypes.wintypes.DWORD()
            ok = wtsapi32.WTSQuerySessionInformationW(0, sid, 5, ctypes.byref(buf), ctypes.byref(size))
            if ok and buf.value:
                username = buf.value.lower()
                wtsapi32.WTSFreeMemory(buf)
                if username == "agent":
                    logger.info(f"[Stop All] Logging off agent (session {sid})")
                    psexec = str(_PSEXEC_PATH)
                    # Try multiple methods to fully logoff the session
                    if _PSEXEC_PATH.exists():
                        # Method 1: reset session (force logoff)
                        r = subprocess.run(
                            [psexec, '-accepteula', '-nobanner', '-s', 'reset', 'session', str(sid)],
                            capture_output=True, text=True, timeout=10,
                        )
                        logger.info(f"[Stop All] Reset session {sid}: {r.stdout.strip()} {r.stderr.strip()}")
                    else:
                        subprocess.run(['logoff', str(sid)], capture_output=True, text=True, timeout=10)
                    logged_off += 1
                else:
                    logger.info(f"[Stop All] Skipping {username} (session {sid})")
        wtsapi32.WTSFreeMemory(pInfo)
    except Exception as e:
        logger.error(f"[Stop All] Session logoff failed: {e}")

    # Step 3: Close agent mstsc windows (localhost only) after logoff
    _time.sleep(2)
    try:
        r = subprocess.run(
            ['tasklist', '/V', '/FI', 'IMAGENAME eq mstsc.exe', '/FO', 'CSV'],
            capture_output=True, text=True, timeout=10,
        )
        for line in r.stdout.strip().split('\n')[1:]:
            parts = line.replace('"', '').split(',')
            if len(parts) >= 2:
                title = parts[-1] if parts else ''
                if 'localhost' in title.lower():
                    try:
                        pid = int(parts[1])
                        subprocess.run(['taskkill', '/F', '/PID', str(pid)], capture_output=True, text=True, timeout=10)
                        logger.info(f"Closed agent mstsc PID {pid}")
                    except (ValueError, IndexError):
                        continue
    except Exception:
        pass

    # Step 4: Force-reset bot slots
    slots_reset = 0
    output_dir = Path(r"C:\ProgramData\MHTAgentic")
    slots_reset = _force_reset_all_slots(output_dir)

    # Verify clean state
    _time.sleep(1)
    verified_clean = _verify_clean()

    logger.info(f"[Stop All] Done: killed={killed}, logged_off={logged_off}, "
                f"slots_reset={slots_reset}, clean={verified_clean}")

    return {
        "success": True,
        "killed": killed,
        "elevated": True,
        "sessions_logged_off": logged_off,
        "slots_reset": slots_reset,
        "verified_clean": verified_clean,
        "sessions": [],
    }


def _force_reset_all_slots(output_dir: Path) -> int:
    """Force-reset both inbound and outbound bot slots to 'open' immediately."""
    db_path = output_dir / "mht_data.db"
    if not db_path.exists():
        return 0
    reset_count = 0
    for slot_name in ("experityb", "experityc", "experityd"):
        try:
            release_slot(str(db_path), slot_name)
            reset_count += 1
            logger.info(f"Force-released bot slot: {slot_name}")
        except Exception as e:
            logger.error(f"Failed to release slot {slot_name}: {e}")
    return reset_count


def _sweep_orphan_processes():
    """
    Kill orphan pythonw.exe, python.exe (running launcher.pyw), and
    lingering mstsc.exe on the local machine via taskkill.
    Catches bots launched via PsExec that may not show up in WTS process lists.
    """
    targets = [
        ("pythonw.exe", "orphan pythonw"),
        ("mstsc.exe", "lingering mstsc"),
    ]
    for image_name, label in targets:
        try:
            result = subprocess.run(
                ["taskkill", "/F", "/IM", image_name],
                capture_output=True, text=True, timeout=10,
            )
            if result.returncode == 0:
                logger.info(f"taskkill swept {label}: {result.stdout.strip()}")
            else:
                # returncode 128 = "no matching processes" — not an error
                logger.debug(f"taskkill {label}: {result.stderr.strip()}")
        except Exception as e:
            logger.error(f"taskkill {label} failed: {e}")


def _verify_clean() -> bool:
    """
    Verify that no RDP sessions or mstsc windows remain.
    Returns True if everything is clean.
    """
    remaining_sessions = _find_rdp_session_ids()
    remaining_windows = find_rdp_windows()
    if remaining_sessions or remaining_windows:
        logger.warning(
            f"Verify: {len(remaining_sessions)} RDP sessions, "
            f"{len(remaining_windows)} mstsc windows still present"
        )
        return False
    logger.info("Verify: clean — no RDP sessions or mstsc windows remain")
    return True


def check_bot_health(db_path: Path, stale_minutes: int = 5) -> Dict:
    """
    Check bot health by querying the SQLite database.

    Args:
        db_path: Path to mht_data.db
        stale_minutes: Minutes before considering the bot stale

    Returns:
        Dict with health info: active, last_event_at, events_today, status
    """
    result = {
        "active": False,
        "last_event_at": None,
        "events_today": 0,
        "inbound_today": 0,
        "outbound_today": 0,
        "errors_today": 0,
        "status": "unknown",
    }

    if not db_path.exists():
        result["status"] = "no_database"
        return result

    try:
        conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        today = datetime.now().strftime("%Y-%m-%d")

        # Last event timestamp
        cursor.execute(
            "SELECT MAX(updated_at) as last_update FROM common_event"
        )
        row = cursor.fetchone()
        if row and row["last_update"]:
            result["last_event_at"] = row["last_update"]
            last_dt = datetime.fromisoformat(row["last_update"])
            age = datetime.now() - last_dt
            result["active"] = age < timedelta(minutes=stale_minutes)
            if result["active"]:
                result["status"] = "active"
            else:
                result["status"] = "stale"
        else:
            result["status"] = "empty"

        # Today's event counts
        cursor.execute(
            "SELECT COUNT(*) as cnt FROM common_event WHERE received_at >= ?",
            (today,),
        )
        result["events_today"] = cursor.fetchone()["cnt"]

        cursor.execute(
            "SELECT COUNT(*) as cnt FROM common_event WHERE received_at >= ? AND direction = 'I'",
            (today,),
        )
        result["inbound_today"] = cursor.fetchone()["cnt"]

        cursor.execute(
            "SELECT COUNT(*) as cnt FROM common_event WHERE received_at >= ? AND direction = 'O'",
            (today,),
        )
        result["outbound_today"] = cursor.fetchone()["cnt"]

        cursor.execute(
            "SELECT COUNT(*) as cnt FROM common_event WHERE received_at >= ? AND status < 0",
            (today,),
        )
        result["errors_today"] = cursor.fetchone()["cnt"]

        conn.close()
    except Exception as e:
        logger.error(f"Health check failed: {e}")
        result["status"] = "error"

    return result


def get_recent_events(db_path: Path, limit: int = 50) -> List[Dict]:
    """
    Get recent patient events from the database.

    Args:
        db_path: Path to mht_data.db
        limit: Max events to return

    Returns:
        List of event dicts
    """
    if not db_path.exists():
        return []

    try:
        conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        cursor.execute("""
            SELECT id, received_at, direction, status, kind, updated_at, error_count,
                   raw_data, converted_data, sent_at, response_data
            FROM common_event
            ORDER BY id DESC
            LIMIT ?
        """, (limit,))

        events = []
        json = __import__("json")
        for row in cursor.fetchall():
            event = dict(row)
            # Parse raw_data (scraped demographics)
            raw = {}
            if event.get("raw_data"):
                try:
                    raw = json.loads(event["raw_data"])
                except Exception:
                    pass
            # Parse converted_data (MHT API format)
            converted = {}
            if event.get("converted_data"):
                try:
                    converted = json.loads(event["converted_data"])
                except Exception:
                    pass

            patient = converted.get("patient", {})
            first = patient.get("patient_first_name", "") or raw.get("first_name", "")
            last = patient.get("patient_last_name", "") or raw.get("last_name", "")
            event["patient_name"] = f"{last}, {first}".strip(", ") if (first or last) else ""
            event["first_name"] = first
            event["last_name"] = last
            event["dob"] = raw.get("dob", "")
            event["cell_phone"] = raw.get("cell_phone", "")
            event["home_phone"] = raw.get("home_phone", "")
            event["email"] = raw.get("email", "")
            event["address"] = raw.get("address1", "")
            event["zip"] = raw.get("zip", "")
            event["location"] = raw.get("clinic_location", "") or raw.get("location", "")

            # Don't send raw JSON blobs to the frontend
            del event["raw_data"]
            del event["converted_data"]
            del event["response_data"]
            events.append(event)

        conn.close()
        return events
    except Exception as e:
        logger.error(f"Failed to get recent events: {e}")
        return []


# --- Start Monitoring (remote launch into RDP sessions) ---

# PsExec64 path (downloaded by dashboard setup)
_PSEXEC_PATH = Path(r"C:\Program Files\pstools\PsExec64.exe")
if not _PSEXEC_PATH.exists():
    _PSEXEC_PATH = Path(r"C:\ProgramData\MHTAgentic\PsExec64.exe")


def start_monitoring_in_sessions(project_root: Path) -> Dict:
    """
    Launch 'launcher.pyw --monitor-only' inside each active RDP session
    using PsExec64 (-s -i <session_id>).

    Each launcher instance auto-claims its own bot slot from the DB
    (no --mode flag needed).
    """
    import sys
    import tempfile

    if not _PSEXEC_PATH.exists():
        return {"success": False, "error": f"PsExec64 not found at {_PSEXEC_PATH}"}

    rdp_sessions = _find_rdp_session_ids()
    if not rdp_sessions:
        return {"success": False, "error": "No RDP sessions found"}

    # Use C:\MHTAgentic if available, else project_root
    shared_root = Path(r"C:\MHTAgentic")
    launcher_path = shared_root / "launcher.pyw"
    if not launcher_path.exists():
        launcher_path = project_root / "launcher.pyw"
    if not launcher_path.exists():
        return {"success": False, "error": "launcher.pyw not found"}

    python_dir = Path(sys.executable).parent
    pythonw_path = python_dir / "pythonw.exe"
    if not pythonw_path.exists():
        pythonw_path = _find_system_python("pythonw.exe")

    # Sort sessions by session_id so launch order is stable
    rdp_sessions.sort(key=lambda s: s["session_id"])

    # Assign modes: first session = inbound, second = outbound
    mode_order = ["inbound", "outbound"]
    sessions_to_launch = []
    for i, sess in enumerate(rdp_sessions):
        sessions_to_launch.append({
            "session_id": sess["session_id"],
            "username": sess["username"],
            "mode": mode_order[i] if i < len(mode_order) else "both",
        })

    if not sessions_to_launch:
        return {"success": False, "error": "No sessions to launch"}

    # Build an elevated batch script that runs PsExec for each session.
    log_file = Path(r"C:\ProgramData\MHTAgentic\monitor_launch.log")

    # Collect Python package paths so SYSTEM can import everything
    # IMPORTANT: C:\ProgramData\MHTAgentic\site-packages MUST be first —
    # PsExec runs as SYSTEM which cannot access user-profile paths
    import site as _site
    extra_paths = [r"C:\ProgramData\MHTAgentic\site-packages"]
    for p in __import__("sys").path:
        if "site-packages" in p:
            if "\\Users\\" in p and "\\Users\\Public" not in p:
                continue
            extra_paths.append(p)
    sys_sp = str(pythonw_path).replace("pythonw.exe", "").rstrip("\\") + "\\Lib\\site-packages"
    if sys_sp not in extra_paths:
        extra_paths.append(sys_sp)
    pythonpath_str = ";".join(extra_paths)

    # pywin32 DLLs path
    _pywin32_dll_dir = str(Path(r"C:\MHTAgentic"))
    stderr_log = Path(r"C:\ProgramData\MHTAgentic\bot_stderr.log")

    # Write a per-session wrapper .bat (avoids cmd /c quoting issues)
    # Use python.exe (not pythonw) with per-session logs to capture ALL output
    python_exe = str(pythonw_path).replace("pythonw.exe", "python.exe")
    wrapper_bats = []
    for s in sessions_to_launch:
        sid = s["session_id"]
        mode = s.get("mode", "")
        mode_flag = f" -{mode}" if mode in ("inbound", "outbound") else ""
        force_user = "ExperityB" if mode == "inbound" else "ExperityC"
        session_log = Path(r"C:\ProgramData\MHTAgentic") / f"bot_{mode}_{sid}.log"

        wrapper_bat = Path(r"C:\ProgramData\MHTAgentic") / f"run_monitor_{sid}.bat"
        wrapper_lines = [
            "@echo off",
            f'set PYTHONPATH={pythonpath_str}',
            f'set FORCE_BOT_USER={force_user}',
            f'set PATH={_pywin32_dll_dir};%PATH%',
            'set PYTHONUTF8=1',
            f'"{python_exe}" -u "{launcher_path}" --monitor-only{mode_flag} > "{session_log}" 2>&1',
        ]
        with open(wrapper_bat, "w") as wf:
            wf.write("\r\n".join(wrapper_lines))
        wrapper_bats.append((sid, wrapper_bat))

    lines = [
        "@echo off",
        f'echo [%date% %time%] Starting monitor launch >> "{log_file}"',
    ]
    for sid, wrapper_bat in wrapper_bats:
        cmd = (
            f'"{_PSEXEC_PATH}" -accepteula -nobanner -s -i {sid} -d '
            f'"{wrapper_bat}"'
        )
        lines.append(f'echo Launching session {sid} >> "{log_file}"')
        lines.append(f'{cmd} >> "{log_file}" 2>&1')
        lines.append(f'echo Exit code: %errorlevel% >> "{log_file}"')
        lines.append("")

    lines.append(f'echo [%date% %time%] Done >> "{log_file}"')

    # Write batch file
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".bat", prefix="mht_start_monitor_", delete=False
    )
    tmp.write("\r\n".join(lines))
    tmp.close()

    logger.info(f"Wrote launch batch: {tmp.name}")
    for s in sessions_to_launch:
        logger.info(f"  Session {s['session_id']} ({s['username']}) → auto-claim")

    # Run elevated via UAC in a background thread
    import threading as _threading

    def _launch_and_cleanup():
        import time as _t
        shell32 = ctypes.windll.shell32
        result = shell32.ShellExecuteW(None, "runas", tmp.name, None, None, 0)
        ok = result > 32
        if ok:
            logger.info(f"Elevated PsExec batch launched for {len(sessions_to_launch)} sessions")
        else:
            logger.error(f"Failed to launch elevated batch (code {result})")
        _t.sleep(30)
        try:
            import os as _os
            _os.unlink(tmp.name)
        except Exception:
            pass

    _threading.Thread(target=_launch_and_cleanup, daemon=True).start()

    return {
        "success": True,
        "message": "UAC prompt shown — accept to start monitoring",
        "sessions": [{"username": s["username"], "mode": "auto"} for s in sessions_to_launch],
    }


def wait_for_rdp_session(username: str, timeout: int = 60) -> Optional[int]:
    """Poll WTS until an RDP session for `username` appears. Returns session_id or None."""
    import time
    deadline = time.time() + timeout
    while time.time() < deadline:
        for sess in _find_rdp_session_ids():
            if sess["username"].lower() == username.lower():
                logger.info(f"RDP session found for {username}: session_id={sess['session_id']}")
                return sess["session_id"]
        time.sleep(2)
    logger.error(f"No RDP session appeared for {username} within {timeout}s")
    return None


def launch_bot_in_session(session_id: int, project_root: Path,
                          username: str = "", agent_mode: str = "") -> Dict:
    """
    Launch launcher.pyw inside a specific RDP session.

    Strategy: PsExec with -s -i <session_id> to run as SYSTEM in the
    interactive session.  launcher.pyw's run_silent() checks username,
    so we set the FORCE_BOT_USER env var to bypass that check when
    running as SYSTEM.
    """
    import sys
    import tempfile

    if not _PSEXEC_PATH.exists():
        return {"success": False, "error": f"PsExec64 not found at {_PSEXEC_PATH}"}

    shared_root = Path(r"C:\MHTAgentic")
    launcher_path = shared_root / "launcher.pyw"
    if not launcher_path.exists():
        launcher_path = project_root / "launcher.pyw"
    if not launcher_path.exists():
        return {"success": False, "error": "launcher.pyw not found"}

    python_dir = Path(sys.executable).parent
    pythonw_path = python_dir / "pythonw.exe"
    if not pythonw_path.exists():
        pythonw_path = _find_system_python("pythonw.exe")

    # Collect site-packages paths
    # IMPORTANT: C:\ProgramData\MHTAgentic\site-packages MUST be first —
    # PsExec runs as SYSTEM which cannot access user-profile paths
    extra_paths = [r"C:\ProgramData\MHTAgentic\site-packages"]
    for p in sys.path:
        if "site-packages" in p:
            if "\\Users\\" in p and "\\Users\\Public" not in p:
                continue
            extra_paths.append(p)
    sys_sp = str(pythonw_path).replace("pythonw.exe", "").rstrip("\\") + "\\Lib\\site-packages"
    if sys_sp not in extra_paths:
        extra_paths.append(sys_sp)
    pythonpath_str = ";".join(extra_paths)

    log_file = Path(r"C:\ProgramData\MHTAgentic\bot_launch.log")
    stderr_log = Path(r"C:\ProgramData\MHTAgentic\bot_stderr.log")

    # Write a wrapper batch script that sets up env vars and launches pythonw.
    # PsExec will run THIS script (not cmd /c with inline env vars).
    # pywin32 DLLs (pywintypes39.dll, pythoncom39.dll) are copied to C:\MHTAgentic
    pywin32_dll_dir = str(shared_root)

    wrapper_bat = Path(r"C:\ProgramData\MHTAgentic") / f"run_bot_{session_id}.bat"
    mode_flag = f" -{agent_mode}" if agent_mode in ("inbound", "outbound") else ""
    wrapper_lines = [
        "@echo off",
        f'set PYTHONPATH={pythonpath_str}',
        f'set FORCE_BOT_USER={username}',
        f'set PATH={pywin32_dll_dir};%PATH%',
        'set PYTHONUTF8=1',
        f'"{pythonw_path}" "{launcher_path}"{mode_flag} 2>> "{stderr_log}"',
    ]
    with open(wrapper_bat, "w") as wf:
        wf.write("\r\n".join(wrapper_lines))

    # Outer batch: calls PsExec to run the wrapper inside the RDP session
    lines = [
        "@echo off",
        f'echo [%date% %time%] Launching bot in session {session_id} (user={username}) >> "{log_file}"',
    ]
    cmd = (
        f'"{_PSEXEC_PATH}" -accepteula -nobanner -s -i {session_id} -d '
        f'"{wrapper_bat}"'
    )
    lines.append(f'{cmd} >> "{log_file}" 2>&1')
    lines.append(f'echo Exit code: %errorlevel% >> "{log_file}"')

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".bat", prefix="mht_launch_bot_", delete=False
    )
    tmp.write("\r\n".join(lines))
    tmp.close()

    logger.info(f"Launching bot in session {session_id} (user={username}) via {tmp.name}")

    shell32 = ctypes.windll.shell32
    result = shell32.ShellExecuteW(None, "runas", tmp.name, None, None, 0)
    ok = result > 32

    if ok:
        logger.info(f"Elevated PsExec launched for session {session_id}")
    else:
        logger.error(f"Failed to launch bot in session {session_id} (code {result})")

    # Clean up temp file after a delay
    import threading as _threading
    def _cleanup():
        import time as _t, os
        _t.sleep(60)
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
    _threading.Thread(target=_cleanup, daemon=True).start()

    return {"success": ok, "session_id": session_id}
