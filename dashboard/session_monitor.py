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
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional

from mhtagentic.db import release_slot

logger = logging.getLogger("mht_dashboard.monitor")

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


def wait_for_otp_complete(username: str, timeout: int = 180,
                          abort_flag: Optional[dict] = None) -> bool:
    """
    Poll for bot-ready signal from a given RDP username.

    Checks TWO sources (whichever fires first = success):
    1. OTP signal file: {_OTP_SIGNAL_DIR}/{username}_otp_complete
    2. Bot slot in DB: bot_slot row for the user shows status='active'

    IMPORTANT: The caller MUST force-release the slot to 'open' before
    launching the RDP session.  Otherwise a stale 'active' slot from a
    previous run will cause an instant false-positive.

    Args:
        username: The RDP session username (e.g. "ExperityB")
        timeout: Max seconds to wait
        abort_flag: Optional dict with "abort" key — if set to True,
                    polling stops immediately and returns False.

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
    logger.info(f"Waiting for OTP/slot signal: {username} (timeout={timeout}s)")

    while _time.time() < deadline:
        # Check abort flag (set by stop-all to cancel in-flight start-all)
        if abort_flag and abort_flag.get("abort"):
            logger.info(f"OTP wait aborted for {username} (stop-all requested)")
            return False

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

    logger.warning(f"OTP signal timeout for {username} after {timeout}s")
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


# --- Start-All signal file helpers ---

def _write_signal(name: str, data: str = "") -> Path:
    """Write a signal file to the shared session_status directory.

    Args:
        name: Signal filename (e.g. "ExperityB_logoff_done")
        data: Optional data to include after the timestamp line

    Returns:
        Path to the written signal file
    """
    _OTP_SIGNAL_DIR.mkdir(parents=True, exist_ok=True)
    sig = _OTP_SIGNAL_DIR / name
    import time as _time
    content = f"ts={_time.strftime('%Y-%m-%dT%H:%M:%S')}"
    if data:
        content += f"\n{data}"
    sig.write_text(content)
    logger.debug(f"Signal written: {name}")
    return sig


def _wait_for_signal(name: str, timeout: float,
                     abort_check=None) -> Optional[str]:
    """Poll for a signal file, returning its contents or None on timeout.

    Args:
        name: Signal filename to wait for
        timeout: Max seconds to wait
        abort_check: Optional callable returning True if we should abort

    Returns:
        File contents as string, or None if timed out / aborted
    """
    import time as _time
    sig = _OTP_SIGNAL_DIR / name
    deadline = _time.time() + timeout
    while _time.time() < deadline:
        if abort_check and abort_check():
            return None
        if sig.exists():
            try:
                return sig.read_text().strip()
            except Exception:
                return ""
        _time.sleep(0.5)
    return None


def _clear_start_signals():
    """Remove all start_all_* signal files for a fresh orchestration cycle."""
    if not _OTP_SIGNAL_DIR.exists():
        return
    patterns = ["start_all_*", "*_logoff_done", "*_mstsc_opened",
                "*_session_found", "*_bot_ready", "*_bot_failed"]
    for pat in patterns:
        for f in _OTP_SIGNAL_DIR.glob(pat):
            try:
                f.unlink()
            except Exception:
                pass
    logger.info("Cleared all start-all signal files")


def write_start_all_script(launch_order: list, project_root: Path) -> Path:
    """Generate the elevated mht_start_all.py orchestrator script.

    This script runs with admin privileges (single UAC prompt) and handles
    ALL logoffs and PsExec injections, coordinated with the dashboard via
    signal files.

    Args:
        launch_order: List of dicts with "username" keys,
                      e.g. [{"username": "ExperityB"}, ...]
        project_root: Path to MHTAgentic project directory.

    Returns:
        Path to the generated script.
    """
    import sys

    script_path = Path(r"C:\ProgramData\MHTAgentic\mht_start_all.py")

    # Resolve paths for embedding in the generated script
    shared_root = Path(r"C:\MHTAgentic")
    launcher_path = shared_root / "launcher.pyw"
    if not launcher_path.exists():
        launcher_path = project_root / "launcher.pyw"

    python_dir = Path(sys.executable).parent
    pythonw_path = python_dir / "pythonw.exe"
    if not pythonw_path.exists():
        pythonw_path = Path(r"C:\Program Files\Python39\pythonw.exe")

    extra_paths = [r"C:\ProgramData\MHTAgentic\site-packages"]
    for p in sys.path:
        if "site-packages" in p:
            extra_paths.append(p)
    pythonpath_str = ";".join(extra_paths)
    pywin32_dll_dir = str(shared_root)

    usernames_repr = [e["username"] for e in launch_order]

    script = f'''\
"""
Elevated Start All — runs as admin to orchestrate all bot launches.

Coordinates with the dashboard via signal files in
C:\\ProgramData\\MHTAgentic\\session_status\\

Generated dynamically by write_start_all_script().
"""
import ctypes
import ctypes.wintypes
import subprocess
import sqlite3
import sys
import time
from pathlib import Path

# --- Configuration (embedded at generation time) ---
USERNAMES = {usernames_repr!r}
PSEXEC_PATH = Path(r"C:\\ProgramData\\MHTAgentic\\PsExec64.exe")
DB_PATH = Path(r"C:\\ProgramData\\MHTAgentic\\mht_data.db")
SIGNAL_DIR = Path(r"C:\\ProgramData\\MHTAgentic\\session_status")
LOG_FILE = Path(r"C:\\ProgramData\\MHTAgentic\\start_all.log")
PYTHONW_PATH = Path(r"{pythonw_path}")
LAUNCHER_PATH = Path(r"{launcher_path}")
PYTHONPATH_STR = r"{pythonpath_str}"
PYWIN32_DLL_DIR = r"{pywin32_dll_dir}"
BOT_READY_TIMEOUT = 180  # seconds
PSEXEC_RETRIES = 3
PSEXEC_RETRY_DELAY = 5  # seconds

# --- Logging ---
log_lines = []
def log(msg):
    ts = time.strftime("%Y-%m-%dT%H:%M:%S")
    line = f"[{{ts}}] {{msg}}"
    log_lines.append(line)
    print(line, flush=True)

def flush_log():
    try:
        LOG_FILE.write_text("\\n".join(log_lines))
    except Exception:
        pass

# --- Signal file helpers ---
def write_signal(name, data=""):
    SIGNAL_DIR.mkdir(parents=True, exist_ok=True)
    sig = SIGNAL_DIR / name
    content = f"ts={{time.strftime('%Y-%m-%dT%H:%M:%S')}}"
    if data:
        content += f"\\n{{data}}"
    sig.write_text(content)
    log(f"Signal written: {{name}}")

def wait_for_signal(name, timeout):
    sig = SIGNAL_DIR / name
    deadline = time.time() + timeout
    while time.time() < deadline:
        if is_aborted():
            return None
        if sig.exists():
            try:
                return sig.read_text().strip()
            except Exception:
                return ""
        time.sleep(0.5)
    return None

def is_aborted():
    return (SIGNAL_DIR / "start_all_abort").exists()

# --- WTS API for session enumeration ---
wts = ctypes.windll.Wtsapi32

class WTS_SESSION_INFO(ctypes.Structure):
    _fields_ = [
        ("SessionId", ctypes.wintypes.DWORD),
        ("pWinStationName", ctypes.wintypes.LPWSTR),
        ("State", ctypes.wintypes.DWORD),
    ]

WTSUserName = 5

def find_rdp_sessions():
    """Return list of dicts with session_id, station, username for RDP sessions."""
    sessions = []
    pInfo = ctypes.POINTER(WTS_SESSION_INFO)()
    count = ctypes.wintypes.DWORD()
    if not wts.WTSEnumerateSessionsW(0, 0, 1, ctypes.byref(pInfo), ctypes.byref(count)):
        return sessions
    for i in range(count.value):
        s = pInfo[i]
        station = (s.pWinStationName or "").lower()
        buf = ctypes.wintypes.LPWSTR()
        size = ctypes.wintypes.DWORD()
        wts.WTSQuerySessionInformationW(0, s.SessionId, WTSUserName,
                                         ctypes.byref(buf), ctypes.byref(size))
        username = buf.value if buf.value else ""
        if not username or station == "console":
            continue
        sessions.append({{
            "session_id": s.SessionId,
            "station": station,
            "username": username,
        }})
    wts.WTSFreeMemory(pInfo)
    return sessions

def logoff_user_sessions(username):
    """Logoff all RDP sessions for a given username."""
    sessions = find_rdp_sessions()
    targets = [s for s in sessions if s["username"].lower() == username.lower()]
    for s in targets:
        sid = s["session_id"]
        result = wts.WTSLogoffSession(0, sid, True)
        log(f"WTSLogoffSession({{sid}}) for {{username}} = {{result}}")
    return len(targets)

def force_release_slot(username):
    """Force-release a bot slot in SQLite so it can be reclaimed."""
    slot_name = username.lower()
    if not DB_PATH.exists():
        return
    try:
        conn = sqlite3.connect(str(DB_PATH), timeout=5)
        conn.execute(
            "UPDATE bot_slot SET status='open', session_id=NULL, "
            "claimed_at=NULL, heartbeat_at=NULL WHERE slot_name=?",
            (slot_name,)
        )
        conn.commit()
        conn.close()
        log(f"Force-released slot: {{slot_name}}")
    except Exception as e:
        log(f"Failed to release slot {{slot_name}}: {{e}}")

def write_wrapper_bat(username, session_id):
    """Write the wrapper .bat file for PsExec to run inside the RDP session."""
    stderr_log = Path(r"C:\\ProgramData\\MHTAgentic") / f"bot_{{username}}_stderr.log"
    wrapper_bat = Path(r"C:\\ProgramData\\MHTAgentic") / f"run_bot_{{session_id}}.bat"
    lines = [
        "@echo off",
        f"set PYTHONPATH={{PYTHONPATH_STR}}",
        f"set FORCE_BOT_USER={{username}}",
        f"set PATH={{PYWIN32_DLL_DIR}};%PATH%",
        "set PYTHONUTF8=1",
        f'"{{PYTHONW_PATH}}" "{{LAUNCHER_PATH}}" 2>> "{{stderr_log}}"',
    ]
    with open(wrapper_bat, "w") as wf:
        wf.write("\\r\\n".join(lines))
    log(f"Wrapper .bat written: {{wrapper_bat}}")
    return wrapper_bat

def run_psexec(session_id, wrapper_bat):
    """Run PsExec synchronously (already elevated) with retries."""
    cmd = [
        str(PSEXEC_PATH), "-accepteula", "-nobanner",
        "-s", "-i", str(session_id), "-d", str(wrapper_bat),
    ]
    for attempt in range(1, PSEXEC_RETRIES + 1):
        if is_aborted():
            return False
        log(f"PsExec attempt {{attempt}}/{{PSEXEC_RETRIES}}: {{' '.join(cmd)}}")
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            log(f"PsExec exit={{result.returncode}} stdout={{result.stdout.strip()}} stderr={{result.stderr.strip()}}")
            if result.returncode == 0:
                return True
        except Exception as e:
            log(f"PsExec exception: {{e}}")
        if attempt < PSEXEC_RETRIES:
            log(f"Retrying PsExec in {{PSEXEC_RETRY_DELAY}}s...")
            time.sleep(PSEXEC_RETRY_DELAY)
    return False

def wait_for_bot_ready(username, timeout=BOT_READY_TIMEOUT):
    """Poll for bot ready: check DB slot active OR OTP signal file."""
    slot_name = username.lower()
    signal_file = SIGNAL_DIR / f"{{username}}_otp_complete"
    deadline = time.time() + timeout
    poll = 0
    while time.time() < deadline:
        if is_aborted():
            return False
        poll += 1
        # Check OTP signal file
        if signal_file.exists():
            log(f"{{username}}: OTP signal file found after {{poll * 3}}s")
            return True
        # Check DB slot
        try:
            conn = sqlite3.connect(f"file:{{DB_PATH}}?mode=ro", uri=True)
            conn.row_factory = sqlite3.Row
            cur = conn.cursor()
            cur.execute("SELECT status FROM bot_slot WHERE slot_name=?", (slot_name,))
            row = cur.fetchone()
            conn.close()
            if row and row["status"] == "active":
                log(f"{{username}}: DB slot active after {{poll * 3}}s")
                return True
        except Exception:
            pass
        if poll % 10 == 0:
            log(f"{{username}}: still waiting ({{poll * 3}}s elapsed)")
        time.sleep(3)
    return False


# ============================================================
#  MAIN ORCHESTRATION
# ============================================================
log("=" * 60)
log("MHT Start All — elevated orchestrator starting")
log(f"Usernames: {{USERNAMES}}")
log(f"PsExec: {{PSEXEC_PATH}} (exists={{PSEXEC_PATH.exists()}})")
log("=" * 60)

SIGNAL_DIR.mkdir(parents=True, exist_ok=True)

for i, username in enumerate(USERNAMES):
    log(f"")
    log(f"── Bot {{i+1}}/{{len(USERNAMES)}}: {{username}} ──")

    if is_aborted():
        log(f"{{username}}: ABORTED")
        break

    # 1. Force-release the bot's slot
    force_release_slot(username)

    # 2. Logoff existing session
    log(f"{{username}}: Logging off existing sessions...")
    logoff_user_sessions(username)
    time.sleep(2)

    # 3. Write logoff_done signal
    write_signal(f"{{username}}_logoff_done")

    if is_aborted():
        log(f"{{username}}: ABORTED after logoff")
        break

    # 4. Wait for dashboard to open mstsc and signal back
    log(f"{{username}}: Waiting for mstsc_opened signal (120s timeout)...")
    result = wait_for_signal(f"{{username}}_mstsc_opened", timeout=120)
    if result is None:
        log(f"{{username}}: TIMEOUT or ABORT waiting for mstsc_opened")
        write_signal(f"{{username}}_bot_failed", "error=mstsc_opened_timeout")
        continue

    if is_aborted():
        break

    # 5. Poll WTS for new session ID
    log(f"{{username}}: Polling WTS for new session (120s timeout)...")
    sid = None
    deadline = time.time() + 120
    poll_n = 0
    while time.time() < deadline:
        if is_aborted():
            break
        all_sessions = find_rdp_sessions()
        poll_n += 1
        # Match by username (case-insensitive)
        for sess in all_sessions:
            if sess["username"].lower() == username.lower():
                sid = sess["session_id"]
                break
        if sid is not None:
            break
        if poll_n <= 3 or poll_n % 5 == 0:
            log(f"{{username}}: WTS poll #{{poll_n}} — {{[(s['username'], s['session_id']) for s in all_sessions]}}")
        time.sleep(2)

    if is_aborted():
        break

    if sid is None:
        log(f"{{username}}: NO SESSION FOUND after 120s")
        write_signal(f"{{username}}_bot_failed", "error=no_session_found")
        continue

    # 6. Write session_found signal
    log(f"{{username}}: Session found! session_id={{sid}}")
    write_signal(f"{{username}}_session_found", f"session_id={{sid}}")

    # 7. Write the wrapper .bat
    wrapper_bat = write_wrapper_bat(username, sid)

    # 8. Run PsExec (synchronous, already elevated)
    log(f"{{username}}: Running PsExec for session {{sid}}...")
    psexec_ok = run_psexec(sid, wrapper_bat)
    if not psexec_ok:
        log(f"{{username}}: PsExec FAILED after {{PSEXEC_RETRIES}} attempts")
        write_signal(f"{{username}}_bot_failed", "error=psexec_failed")
        continue

    if is_aborted():
        break

    # 9. Poll for bot ready
    log(f"{{username}}: Waiting for bot ready ({{BOT_READY_TIMEOUT}}s timeout)...")
    ready = wait_for_bot_ready(username)

    # 10. Write result signal
    if ready:
        log(f"{{username}}: BOT READY")
        write_signal(f"{{username}}_bot_ready")
    else:
        log(f"{{username}}: Bot NOT ready after {{BOT_READY_TIMEOUT}}s")
        write_signal(f"{{username}}_bot_failed", "error=bot_not_ready_timeout")

# Final signal
write_signal("start_all_complete")
log("")
log("=" * 60)
log("MHT Start All — orchestrator finished")
log("=" * 60)
flush_log()
sys.exit(0)
'''

    script_path.write_text(script)
    logger.info(f"[write_start_all_script] Generated: {script_path}")
    return script_path


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
        pythonw_path = Path(r"C:\Program Files\Python39\pythonw.exe")

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


def _logoff_rdp_user(username: str) -> bool:
    """Logoff any existing session for a user so next connect is a fresh logon.
    Uses an elevated script because logging off another user requires admin."""
    import sys
    import tempfile

    rdp_sessions = _find_rdp_session_ids()
    target_sessions = [
        s for s in rdp_sessions
        if s["username"].lower() == username.lower()
    ]
    if not target_sessions:
        logger.info(f"No existing session found for {username}")
        return False

    session_ids = [s["session_id"] for s in target_sessions]
    logger.info(f"Found sessions for {username}: {session_ids} — logging off (elevated)")

    # Elevated script to logoff sessions
    script = f"""
import ctypes, ctypes.wintypes, sys, time
wts = ctypes.windll.Wtsapi32
session_ids = {session_ids!r}
for sid in session_ids:
    result = wts.WTSLogoffSession(0, sid, True)
    print(f"WTSLogoffSession({{sid}}) = {{result}}")
time.sleep(1)
sys.exit(0)
"""
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".py", prefix="mht_logoff_", delete=False
    )
    tmp.write(script)
    tmp.close()

    shell32 = ctypes.windll.shell32
    python_exe = sys.executable
    result = shell32.ShellExecuteW(None, "runas", python_exe, f'"{tmp.name}"', None, 0)
    ok = result > 32

    if ok:
        logger.info(f"Elevated logoff launched for sessions {session_ids}")
    else:
        logger.error(f"Failed to elevate logoff (code {result})")

    # Clean up temp file
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
    Stop everything: kill all apps inside RDP sessions, logoff sessions,
    close mstsc windows, reset slots.

    Runs an elevated Python script (UAC prompt) that has full admin power to
    kill processes in other users' sessions and logoff those sessions.
    """
    import time as _time
    import sys

    _STOP_SCRIPT = Path(r"C:\ProgramData\MHTAgentic\mht_stop_all.py")
    _SIGNAL_FILE = Path(r"C:\ProgramData\MHTAgentic\stop_all_done")

    # Clean up OTP signals
    clear_all_otp_signals()

    # Remove previous signal file
    try:
        _SIGNAL_FILE.unlink(missing_ok=True)
    except Exception:
        pass

    # Launch the elevated stop script (triggers one UAC prompt)
    python_exe = Path(r"C:\Program Files\Python39\python.exe")
    if not python_exe.exists():
        python_exe = Path(sys.executable)

    shell32 = ctypes.windll.shell32
    result = shell32.ShellExecuteW(
        None, "runas", str(python_exe),
        f'"{_STOP_SCRIPT}"', None, 0  # SW_HIDE — runs silently
    )
    launched = result > 32

    if not launched:
        logger.error(f"Failed to launch elevated stop script (code {result})")
        # Fallback: at least kill mstsc on the local console
        _sweep_orphan_processes()
        return {
            "success": False,
            "error": "UAC elevation failed or was declined",
            "killed": 0,
            "sessions_logged_off": 0,
            "slots_reset": 0,
            "verified_clean": False,
        }

    logger.info("Elevated stop script launched — waiting for completion...")

    # Wait for the signal file (elevated script writes it when done)
    deadline = _time.time() + 30
    while _time.time() < deadline:
        if _SIGNAL_FILE.exists():
            break
        _time.sleep(0.5)

    # Parse results from signal file
    killed = 0
    logged_off = 0
    slots_reset = 0
    if _SIGNAL_FILE.exists():
        try:
            text = _SIGNAL_FILE.read_text().strip()
            for part in text.split(","):
                k, v = part.split("=")
                if k == "killed":
                    killed = int(v)
                elif k == "logged_off":
                    logged_off = int(v)
                elif k == "slots_reset":
                    slots_reset = int(v)
        except Exception:
            pass
        logger.info(f"Stop All complete: killed={killed}, logged_off={logged_off}")
    else:
        logger.warning("Stop All signal file not found — script may still be running")

    # Verify clean state
    _time.sleep(1)
    verified_clean = _verify_clean()

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
                   converted_data
            FROM common_event
            ORDER BY id DESC
            LIMIT ?
        """, (limit,))

        events = []
        for row in cursor.fetchall():
            event = dict(row)
            # Extract patient name from converted_data if available
            if event.get("converted_data"):
                try:
                    data = __import__("json").loads(event["converted_data"])
                    patient = data.get("patient", {})
                    name = f"{patient.get('patient_last_name', '')}, {patient.get('patient_first_name', '')}"
                    event["patient_name"] = name.strip(", ")
                except Exception:
                    event["patient_name"] = ""
            else:
                event["patient_name"] = ""
            # Don't send raw JSON blobs to the frontend
            del event["converted_data"]
            events.append(event)

        conn.close()
        return events
    except Exception as e:
        logger.error(f"Failed to get recent events: {e}")
        return []


# --- Start Monitoring (remote launch into RDP sessions) ---

# PsExec64 path (downloaded by dashboard setup)
_PSEXEC_PATH = Path(r"C:\ProgramData\MHTAgentic\PsExec64.exe")


def start_monitoring_in_sessions(project_root: Path) -> Dict:
    """
    Launch 'launcher.pyw --monitor-only' inside each active RDP session.

    Uses an elevated helper script that calls WTSQueryUserToken +
    CreateProcessAsUser to spawn each bot as the interactive session user
    (NOT SYSTEM). This gives the bot full UIA access to windows in that
    session — something PsExec and scheduled tasks cannot do.
    """
    import sys

    rdp_sessions = _find_rdp_session_ids()
    if not rdp_sessions:
        return {"success": False, "error": "No RDP sessions found"}

    shared_root = Path(r"C:\MHTAgentic")
    launcher_path = shared_root / "launcher.pyw"
    if not launcher_path.exists():
        launcher_path = project_root / "launcher.pyw"
    if not launcher_path.exists():
        return {"success": False, "error": "launcher.pyw not found"}

    python_dir = Path(sys.executable).parent
    python_exe = python_dir / "python.exe"
    if not python_exe.exists():
        python_exe = Path(r"C:\Program Files\Python39\python.exe")

    # Bot config: bot_name -> (mode, location)
    BOT_CONFIG = {
        "ExperityB": ("inbound",  "ATTALLA"),
        "ExperityC": ("inbound",  "ANNISTON"),
        "ExperityD": ("outbound", "ATTALLA"),
    }

    # Assign bots to sessions in sorted session-ID order.
    # All WTS sessions show the same username so we can't map by name;
    # instead we rely on deterministic ordering: lowest session ID gets
    # the first bot, etc.
    BOT_ORDER = ["ExperityB", "ExperityC", "ExperityD"]
    rdp_sessions.sort(key=lambda s: s["session_id"])

    sessions_to_launch = []
    for i, sess in enumerate(rdp_sessions[:len(BOT_ORDER)]):
        bot_name = BOT_ORDER[i]
        mode, location = BOT_CONFIG[bot_name]
        sessions_to_launch.append({
            "session_id": sess["session_id"],
            "username": bot_name,
            "mode": mode,
            "location": location,
        })

    if not sessions_to_launch:
        return {"success": False, "error": "No sessions to launch"}

    # Collect Python package paths — ProgramData paths first (including
    # win32 subdirs for pywin32), then system site-packages.
    pd_sp = r"C:\ProgramData\MHTAgentic\site-packages"
    extra_paths = [
        pd_sp,
        pd_sp + r"\win32",
        pd_sp + r"\win32\lib",
        pd_sp + r"\Pythonwin",
    ]
    for p in sys.path:
        if "site-packages" in p and p not in extra_paths:
            extra_paths.append(p)
    pythonpath_str = ";".join(extra_paths)
    pywin32_dll_dir = str(shared_root)
    pywin32_sys32 = r"C:\ProgramData\MHTAgentic\site-packages\pywin32_system32"

    # Write per-session wrapper .bat files
    for s in sessions_to_launch:
        sid = s["session_id"]
        mode = s["mode"]
        location = s["location"]
        force_user = s["username"]
        mode_flag = f" -{mode}" if mode in ("inbound", "outbound") else ""
        session_log = Path(r"C:\ProgramData\MHTAgentic") / f"bot_{mode}_{sid}.log"

        wrapper_bat = Path(r"C:\ProgramData\MHTAgentic") / f"run_monitor_{sid}.bat"
        wrapper_lines = [
            "@echo off",
            f'echo [%date% %time%] Starting bot {force_user} mode={mode} location={location} > "{session_log}" 2>&1',
            f'set PYTHONPATH={pythonpath_str}',
            f'set FORCE_BOT_USER={force_user}',
            f'set BOT_LOCATION={location}',
            f'set PATH={pywin32_sys32};{pywin32_dll_dir};%PATH%',
            'set PYTHONUTF8=1',
            'set PYTHONNOUSERSITE=1',
            f'echo [%date% %time%] Launching python... >> "{session_log}" 2>&1',
            f'echo [%date% %time%] PYTHONPATH=%PYTHONPATH% >> "{session_log}" 2>&1',
            f'echo [%date% %time%] PYTHONNOUSERSITE=%PYTHONNOUSERSITE% >> "{session_log}" 2>&1',
            f'"{python_exe}" -s -u -c "import sys; print(\'sys.path=\', sys.path[:5]); import PIL; print(\'PIL=\', PIL.__file__, getattr(PIL,\'__path__\',\'\'))" >> "{session_log}" 2>&1',
            f'"{python_exe}" -s -u "{launcher_path}" --monitor-only{mode_flag} >> "{session_log}" 2>&1',
            f'echo [%date% %time%] Python exited with code %errorlevel% >> "{session_log}" 2>&1',
        ]
        with open(wrapper_bat, "w") as wf:
            wf.write("\r\n".join(wrapper_lines))

    # Build an elevated Python script that uses CreateProcessAsUser
    # to launch each wrapper .bat as the interactive session user.
    launch_items = []
    for s in sessions_to_launch:
        sid = s["session_id"]
        bat = str(Path(r"C:\ProgramData\MHTAgentic") / f"run_monitor_{sid}.bat")
        launch_items.append((sid, bat, s["username"], s["mode"], s["location"]))

    helper_script = r'''
import ctypes, ctypes.wintypes, sys, time

kernel32 = ctypes.windll.kernel32
advapi32 = ctypes.windll.advapi32
wtsapi32 = ctypes.windll.Wtsapi32
userenv = ctypes.windll.userenv

log = open(r"C:\ProgramData\MHTAgentic\monitor_launch.log", "a")
def logmsg(msg):
    import datetime
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log.write(f"[{ts}] {msg}\n")
    log.flush()

logmsg("=== CreateProcessAsUser launcher ===")

# Enable SeDebugPrivilege + SeAssignPrimaryTokenPrivilege
class LUID(ctypes.Structure):
    _fields_ = [("Lo", ctypes.wintypes.DWORD), ("Hi", ctypes.wintypes.LONG)]
class LA(ctypes.Structure):
    _fields_ = [("Luid", LUID), ("Attr", ctypes.wintypes.DWORD)]
class TP(ctypes.Structure):
    _fields_ = [("Count", ctypes.wintypes.DWORD), ("Privs", LA * 3)]

tok = ctypes.wintypes.HANDLE()
advapi32.OpenProcessToken(kernel32.GetCurrentProcess(), 0x0028, ctypes.byref(tok))
tp = TP()
tp.Count = 3
for i, priv_name in enumerate(["SeDebugPrivilege", "SeAssignPrimaryTokenPrivilege", "SeIncreaseQuotaPrivilege"]):
    luid = LUID()
    advapi32.LookupPrivilegeValueW(None, priv_name, ctypes.byref(luid))
    tp.Privs[i].Luid = luid
    tp.Privs[i].Attr = 2  # SE_PRIVILEGE_ENABLED
advapi32.AdjustTokenPrivileges(tok, False, ctypes.byref(tp), 0, None, None)
kernel32.CloseHandle(tok)
logmsg("Privileges adjusted")

# STARTUPINFO
class STARTUPINFOW(ctypes.Structure):
    _fields_ = [
        ("cb", ctypes.wintypes.DWORD),
        ("lpReserved", ctypes.wintypes.LPWSTR),
        ("lpDesktop", ctypes.wintypes.LPWSTR),
        ("lpTitle", ctypes.wintypes.LPWSTR),
        ("dwX", ctypes.wintypes.DWORD),
        ("dwY", ctypes.wintypes.DWORD),
        ("dwXSize", ctypes.wintypes.DWORD),
        ("dwYSize", ctypes.wintypes.DWORD),
        ("dwXCountChars", ctypes.wintypes.DWORD),
        ("dwYCountChars", ctypes.wintypes.DWORD),
        ("dwFillAttribute", ctypes.wintypes.DWORD),
        ("dwFlags", ctypes.wintypes.DWORD),
        ("wShowWindow", ctypes.wintypes.WORD),
        ("cbReserved2", ctypes.wintypes.WORD),
        ("lpReserved2", ctypes.c_void_p),
        ("hStdInput", ctypes.wintypes.HANDLE),
        ("hStdOutput", ctypes.wintypes.HANDLE),
        ("hStdError", ctypes.wintypes.HANDLE),
    ]

class PROCESS_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("hProcess", ctypes.wintypes.HANDLE),
        ("hThread", ctypes.wintypes.HANDLE),
        ("dwProcessId", ctypes.wintypes.DWORD),
        ("dwThreadId", ctypes.wintypes.DWORD),
    ]

TOKEN_DUPLICATE = 0x0002
TOKEN_QUERY = 0x0008
TOKEN_ASSIGN_PRIMARY = 0x0001
TOKEN_ADJUST_DEFAULT = 0x0080
TOKEN_ADJUST_SESSIONID = 0x0100
SecurityImpersonation = 2
TokenPrimary = 1
CREATE_NEW_CONSOLE = 0x00000010
CREATE_UNICODE_ENVIRONMENT = 0x00000400

WTSQueryUserToken = wtsapi32.WTSQueryUserToken
WTSQueryUserToken.argtypes = [ctypes.wintypes.ULONG, ctypes.POINTER(ctypes.wintypes.HANDLE)]
WTSQueryUserToken.restype = ctypes.wintypes.BOOL

DuplicateTokenEx = advapi32.DuplicateTokenEx
DuplicateTokenEx.argtypes = [
    ctypes.wintypes.HANDLE, ctypes.wintypes.DWORD, ctypes.c_void_p,
    ctypes.c_int, ctypes.c_int, ctypes.POINTER(ctypes.wintypes.HANDLE)
]
DuplicateTokenEx.restype = ctypes.wintypes.BOOL

CreateEnvironmentBlock = userenv.CreateEnvironmentBlock
CreateEnvironmentBlock.argtypes = [ctypes.POINTER(ctypes.c_void_p), ctypes.wintypes.HANDLE, ctypes.wintypes.BOOL]
CreateEnvironmentBlock.restype = ctypes.wintypes.BOOL

DestroyEnvironmentBlock = userenv.DestroyEnvironmentBlock
DestroyEnvironmentBlock.argtypes = [ctypes.c_void_p]
DestroyEnvironmentBlock.restype = ctypes.wintypes.BOOL

CreateProcessAsUserW = advapi32.CreateProcessAsUserW
CreateProcessAsUserW.argtypes = [
    ctypes.wintypes.HANDLE, ctypes.wintypes.LPCWSTR, ctypes.wintypes.LPWSTR,
    ctypes.c_void_p, ctypes.c_void_p, ctypes.wintypes.BOOL,
    ctypes.wintypes.DWORD, ctypes.c_void_p, ctypes.wintypes.LPCWSTR,
    ctypes.POINTER(STARTUPINFOW), ctypes.POINTER(PROCESS_INFORMATION)
]
CreateProcessAsUserW.restype = ctypes.wintypes.BOOL

LAUNCHES = ''' + repr(launch_items) + r'''

for session_id, bat_path, bot_name, mode, location in LAUNCHES:
    logmsg(f"Launching {bot_name} ({mode} @ {location}) in session {session_id}")

    # Get the session user's token
    user_token = ctypes.wintypes.HANDLE()
    if not WTSQueryUserToken(session_id, ctypes.byref(user_token)):
        err = kernel32.GetLastError()
        logmsg(f"  ERROR: WTSQueryUserToken failed for session {session_id}, error={err}")
        continue

    # Duplicate to a primary token
    dup_token = ctypes.wintypes.HANDLE()
    desired = TOKEN_QUERY | TOKEN_DUPLICATE | TOKEN_ASSIGN_PRIMARY | TOKEN_ADJUST_DEFAULT | TOKEN_ADJUST_SESSIONID
    if not DuplicateTokenEx(user_token, desired, None, SecurityImpersonation, TokenPrimary, ctypes.byref(dup_token)):
        err = kernel32.GetLastError()
        logmsg(f"  ERROR: DuplicateTokenEx failed, error={err}")
        kernel32.CloseHandle(user_token)
        continue

    # Create environment block for that user
    env_block = ctypes.c_void_p()
    CreateEnvironmentBlock(ctypes.byref(env_block), dup_token, False)

    # Set up STARTUPINFO targeting the session's winsta0\default desktop
    si = STARTUPINFOW()
    si.cb = ctypes.sizeof(STARTUPINFOW)
    si.lpDesktop = "winsta0\\default"

    pi = PROCESS_INFORMATION()

    cmd_line = f'cmd.exe /c "{bat_path}"'
    ok = CreateProcessAsUserW(
        dup_token,
        None,
        cmd_line,
        None, None, False,
        CREATE_NEW_CONSOLE | CREATE_UNICODE_ENVIRONMENT,
        env_block,
        None,
        ctypes.byref(si),
        ctypes.byref(pi),
    )

    if ok:
        logmsg(f"  OK: {bot_name} started, PID={pi.dwProcessId}")
        kernel32.CloseHandle(pi.hProcess)
        kernel32.CloseHandle(pi.hThread)
    else:
        err = kernel32.GetLastError()
        logmsg(f"  ERROR: CreateProcessAsUserW failed, error={err}")

    if env_block:
        DestroyEnvironmentBlock(env_block)
    kernel32.CloseHandle(dup_token)
    kernel32.CloseHandle(user_token)

logmsg("=== All done ===")
log.close()
'''

    # Write the helper script
    helper_path = Path(r"C:\ProgramData\MHTAgentic") / "launch_as_user.py"
    with open(helper_path, "w", encoding="utf-8") as f:
        f.write(helper_script)

    logger.info(f"Wrote CreateProcessAsUser helper: {helper_path}")
    for s in sessions_to_launch:
        logger.info(f"  session {s['session_id']}: {s['username']} -> {s['mode']} @ {s['location']}")

    # WTSQueryUserToken requires SYSTEM, so we use PsExec -s to run the
    # helper as SYSTEM.  The helper then uses CreateProcessAsUser to spawn
    # each bot as the real interactive user (with full UIA access).
    # PsExec -s WITHOUT -i runs in session 0 (services), which is fine —
    # the helper targets user sessions via CreateProcessAsUser.
    if not _PSEXEC_PATH.exists():
        return {"success": False, "error": f"PsExec64 not found at {_PSEXEC_PATH}"}

    # Write a tiny .bat that PsExec runs as SYSTEM
    launcher_bat = Path(r"C:\ProgramData\MHTAgentic") / "run_launcher_helper.bat"
    with open(launcher_bat, "w") as f:
        f.write(f'@echo off\r\n"{python_exe}" -u "{helper_path}"\r\n')

    shell32 = ctypes.windll.shell32
    # Elevated .bat that calls PsExec -s (SYSTEM) to run our helper
    psexec_bat = Path(r"C:\ProgramData\MHTAgentic") / "psexec_launcher.bat"
    with open(psexec_bat, "w") as f:
        f.write(
            f'@echo off\r\n'
            f'"{_PSEXEC_PATH}" -accepteula -nobanner -s -d '
            f'"{python_exe}" -u "{helper_path}"\r\n'
        )

    result = shell32.ShellExecuteW(
        None, "runas",
        str(psexec_bat),
        None, None, 0  # SW_HIDE
    )
    ok = result > 32

    if ok:
        logger.info(f"PsExec -s helper launched for {len(sessions_to_launch)} sessions")
    else:
        logger.error(f"Failed to launch PsExec helper (code {result})")
        return {"success": False, "error": f"UAC elevation failed (code {result})"}

    return {
        "success": True,
        "message": f"Launching bots in {len(sessions_to_launch)} sessions (UAC prompt shown)",
        "sessions": [
            {"username": s["username"], "mode": s["mode"], "location": s["location"], "session_id": s["session_id"]}
            for s in sessions_to_launch
        ],
    }


def wait_for_rdp_session(username: str, timeout: int = 60,
                         abort_flag: Optional[dict] = None) -> Optional[int]:
    """Poll WTS until an RDP session for `username` appears. Returns session_id or None."""
    import time
    deadline = time.time() + timeout
    poll_count = 0
    while time.time() < deadline:
        if abort_flag and abort_flag.get("abort"):
            logger.info(f"[wait_for_rdp_session] Aborted for {username} (stop-all)")
            return None
        all_sessions = _find_rdp_session_ids()
        poll_count += 1
        if poll_count <= 3 or poll_count % 5 == 0:
            logger.info(
                f"[wait_for_rdp_session] Poll #{poll_count} for '{username}' — "
                f"found {len(all_sessions)} sessions: "
                f"{[(s['username'], s['session_id'], s['station']) for s in all_sessions]}"
            )
        for sess in all_sessions:
            if sess["username"].lower() == username.lower():
                logger.info(
                    f"[wait_for_rdp_session] MATCH: {username} → "
                    f"session_id={sess['session_id']}, station={sess['station']}"
                )
                return sess["session_id"]
        time.sleep(2)
    logger.error(f"[wait_for_rdp_session] TIMEOUT: no session for {username} after {timeout}s ({poll_count} polls)")
    return None


def _write_wrapper_bat(username: str, project_root: Path,
                       session_id: int = 0, location: str = "ATTALLA") -> Path:
    """Create a wrapper .bat that sets env vars and launches pythonw for a bot.

    This .bat is what PsExec runs inside the RDP session. It sets up
    PYTHONPATH, FORCE_BOT_USER, BOT_LOCATION, PATH (for pywin32 DLLs),
    and launches launcher.pyw with per-user stderr logging.

    Args:
        username: Bot user (e.g. "ExperityB") — used for FORCE_BOT_USER
                  and per-user stderr log filename.
        project_root: Path to MHTAgentic project directory.
        session_id: Windows session ID — used in the .bat filename.
        location: Clinic location (e.g. "ANNISTON", "ATTALLA").

    Returns:
        Path to the written .bat file.
    """
    import sys

    shared_root = Path(r"C:\MHTAgentic")
    launcher_path = shared_root / "launcher.pyw"
    if not launcher_path.exists():
        launcher_path = project_root / "launcher.pyw"

    python_dir = Path(sys.executable).parent
    pythonw_path = python_dir / "pythonw.exe"
    if not pythonw_path.exists():
        pythonw_path = Path(r"C:\Program Files\Python39\pythonw.exe")

    # Collect site-packages paths — ProgramData copy first for SYSTEM access
    extra_paths = [r"C:\ProgramData\MHTAgentic\site-packages"]
    for p in sys.path:
        if "site-packages" in p:
            extra_paths.append(p)
    pythonpath_str = ";".join(extra_paths)

    pywin32_dll_dir = str(shared_root)
    stderr_log = Path(r"C:\ProgramData\MHTAgentic") / f"bot_{username}_stderr.log"

    wrapper_bat = Path(r"C:\ProgramData\MHTAgentic") / f"run_bot_{session_id}.bat"
    wrapper_lines = [
        "@echo off",
        f'set PYTHONPATH={pythonpath_str}',
        f'set FORCE_BOT_USER={username}',
        f'set BOT_LOCATION={location}',
        f'set PATH={pywin32_dll_dir};%PATH%',
        'set PYTHONUTF8=1',
        f'"{pythonw_path}" "{launcher_path}" 2>> "{stderr_log}"',
    ]
    with open(wrapper_bat, "w") as wf:
        wf.write("\r\n".join(wrapper_lines))

    logger.info(f"[_write_wrapper_bat] Written: {wrapper_bat} (user={username}, sid={session_id})")
    return wrapper_bat


def launch_bot_in_session(session_id: int, project_root: Path,
                          username: str = "", agent_mode: str = "",
                          location: str = "ATTALLA") -> Dict:
    """
    Launch launcher.pyw inside a specific RDP session.

    Strategy: PsExec with -s -i <session_id> to run as SYSTEM in the
    interactive session.  launcher.pyw's run_silent() checks username,
    so we set the FORCE_BOT_USER env var to bypass that check when
    running as SYSTEM.
    """
    import tempfile

    if not _PSEXEC_PATH.exists():
        return {"success": False, "error": f"PsExec64 not found at {_PSEXEC_PATH}"}

    shared_root = Path(r"C:\MHTAgentic")
    launcher_path = shared_root / "launcher.pyw"
    if not launcher_path.exists():
        launcher_path = project_root / "launcher.pyw"
    if not launcher_path.exists():
        return {"success": False, "error": "launcher.pyw not found"}

    log_file = Path(r"C:\ProgramData\MHTAgentic\bot_launch.log")

    wrapper_bat = _write_wrapper_bat(username, project_root, session_id, location=location)

    # Outer batch: calls PsExec to run the wrapper inside the RDP session
    lines = [
        "@echo off",
        f'echo [%date% %time%] Launching bot in session {session_id} (user={username}) >> "{log_file}"',
    ]
    cmd = (
        f'"{_PSEXEC_PATH}" -accepteula -nobanner -i {session_id} -d '
        f'"{wrapper_bat}"'
    )
    lines.append(f'{cmd} >> "{log_file}" 2>&1')
    lines.append(f'echo Exit code: %errorlevel% >> "{log_file}"')

    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".bat", prefix="mht_launch_bot_", delete=False
    )
    tmp.write("\r\n".join(lines))
    tmp.close()

    logger.info(
        f"[launch_bot_in_session] session_id={session_id}, user={username}, "
        f"mode={agent_mode}, wrapper={wrapper_bat}, launcher={launcher_path}"
    )
    logger.info(f"[launch_bot_in_session] Elevating via ShellExecuteW(runas) → {tmp.name}")

    shell32 = ctypes.windll.shell32
    result = shell32.ShellExecuteW(None, "runas", tmp.name, None, None, 0)
    ok = result > 32

    if ok:
        logger.info(f"[launch_bot_in_session] SUCCESS — elevated PsExec launched for session {session_id} (user={username})")
    else:
        logger.error(f"[launch_bot_in_session] FAILED — ShellExecuteW returned {result} for session {session_id} (user={username})")

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
