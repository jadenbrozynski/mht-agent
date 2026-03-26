"""
start_all_clean.py — Open RDP sessions and launch Experity login in each.

Uses the same CreateProcessAsUser approach as start_monitoring (proven to work).
One UAC prompt to approve the elevated helper, then everything runs automatically.

Usage:
    python start_all_clean.py                # all 3 RDPs sequentially
    python start_all_clean.py ExperityB      # single RDP
"""

import sys
import os
import time
import ctypes
import ctypes.wintypes
import subprocess
import sqlite3
from pathlib import Path
from datetime import datetime, timezone

from mhtagentic.db.database import (
    get_config,
    get_all_locations,
    release_all_locations,
    assign_location,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _find_system_python(exe_name="python.exe"):
    """Find Python exe dynamically — prefers system-wide install."""
    # Check Program Files first — accessible by all users
    for base in [r"C:\Program Files", r"C:\Program Files (x86)"]:
        for d in sorted(Path(base).glob("Python*"), reverse=True):
            candidate = d / exe_name
            if candidate.exists():
                return candidate
    import shutil
    found = shutil.which(exe_name)
    if found:
        return Path(found)
    return Path(exe_name)

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
PSEXEC_PATH = Path(r"C:\Program Files\pstools\PsExec64.exe")
if not PSEXEC_PATH.exists():
    PSEXEC_PATH = Path(r"C:\ProgramData\MHTAgentic\PsExec64.exe")
RDP_DIR = Path.home() / "Desktop"
DATA_DIR = Path(r"C:\ProgramData\MHTAgentic")
DB_PATH = DATA_DIR / "mht_data.db"

# Template .rdp file to clone settings from (username, password, etc.)
_RDP_TEMPLATE = RDP_DIR / "ExperityB_MHT.rdp"


def _generate_rdp_files(count: int) -> list:
    """Dynamically generate .rdp files arranged in a grid layout.

    Returns list of (label, rdp_key, rdp_path) tuples.
    """
    import ctypes
    screen_w = ctypes.windll.user32.GetSystemMetrics(0)
    screen_h = ctypes.windll.user32.GetSystemMetrics(1)

    # Read template for username/password/settings
    template_lines = []
    if _RDP_TEMPLATE.exists():
        template_lines = _RDP_TEMPLATE.read_text().strip().splitlines()

    # Extract settings from template (everything except winposstr)
    base_settings = [l for l in template_lines if not l.startswith("winposstr:")]
    # Extract password line for reuse
    password_line = next((l for l in template_lines if l.startswith("password 51:")), "")
    username_line = next((l for l in template_lines if l.startswith("username:")), "username:s:ExperityB")

    # Grid layout: 2 columns, 16:9 aspect ratio windows
    cols = min(count, 2)
    rows = (count + cols - 1) // cols
    cell_w = screen_w // cols
    cell_h = screen_h // rows

    rdp_labels = []  # (label, rdp_key, rdp_path)
    slot_letters = "BCDEFGHIJKLMNOP"

    for i in range(count):
        col = i % cols
        row = i // cols
        x1 = col * cell_w
        y1 = row * cell_h
        x2 = x1 + cell_w
        y2 = y1 + cell_h

        letter = slot_letters[i] if i < len(slot_letters) else str(i)
        rdp_key = f"Experity{letter}"
        label = f"RDP{i + 1}"
        rdp_path = RDP_DIR / f"{rdp_key}_MHT.rdp"

        # Build .rdp file content — force 1280x720 + smart sizing
        has_width = has_height = has_smart = has_screenmode = False
        lines = []
        for l in base_settings:
            if l.startswith("username:"):
                lines.append(username_line)
            elif l.startswith("desktopwidth:"):
                lines.append("desktopwidth:i:1280")
                has_width = True
            elif l.startswith("desktopheight:"):
                lines.append("desktopheight:i:720")
                has_height = True
            elif l.startswith("smart sizing:"):
                lines.append("smart sizing:i:1")
                has_smart = True
            elif l.startswith("screen mode id:"):
                lines.append("screen mode id:i:1")
                has_screenmode = True
            else:
                lines.append(l)
        if not has_width:
            lines.append("desktopwidth:i:1280")
        if not has_height:
            lines.append("desktopheight:i:720")
        if not has_smart:
            lines.append("smart sizing:i:1")
        if not has_screenmode:
            lines.append("screen mode id:i:1")
        lines.append(f"winposstr:s:0,1,{x1},{y1},{x2},{y2}")

        rdp_path.write_text("\n".join(lines) + "\n")
        rdp_labels.append((label, rdp_key, rdp_path))

    return rdp_labels


# Legacy dicts for backward compatibility (populated dynamically at startup)
RDP_FILES = {}
ALL_RDPS = []

STATUS_LABELS = {
    5: "Retrying login",
    10: "Opening Experity",
    20: "Entering username (Screen 1)",
    30: "Clicking Next (Screen 1)",
    40: "Waiting for Okta sign-in",
    50: "Username check (Screen 2)",
    60: "Entering password (Screen 2)",
    70: "Clicking Sign In (Screen 2)",
    80: "Generating OTP code (Screen 3)",
    90: "Entering OTP code (Screen 3)",
    100: "Clicking Verify (Screen 3)",
    110: "Login complete",
    # Post-login phase (launcher.pyw --monitor-only)
    120: "Waiting for Tracking Board",
    130: "EMR found",
    140: "Switching location",
    150: "Location confirmed",
    160: "Location claimed",
    200: "Monitoring active",
    -1: "Error",
    -2: "Skipped (already logged in)",
}

# ---------------------------------------------------------------------------
# SQLite helpers
# ---------------------------------------------------------------------------
def _init_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=10)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS clean_startup_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_seed TEXT NOT NULL,
            rdp_label TEXT NOT NULL,
            session_id INTEGER,
            status_code INTEGER NOT NULL,
            status_msg TEXT DEFAULT '',
            timestamp TEXT NOT NULL,
            duration_ms INTEGER,
            complete INTEGER DEFAULT 0
        )
    """)
    conn.commit()
    conn.close()


def _log_status(run_seed, rdp_label, session_id, status_code, status_msg="",
                duration_ms=None, complete=0):
    try:
        conn = sqlite3.connect(str(DB_PATH), timeout=10)
        conn.execute(
            """INSERT INTO clean_startup_log
               (run_seed, rdp_label, session_id, status_code, status_msg,
                timestamp, duration_ms, complete)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (run_seed, rdp_label, session_id, status_code, status_msg,
             datetime.now(timezone.utc).isoformat(), duration_ms, complete),
        )
        conn.commit()
        conn.close()
    except Exception as e:
        print(f"  DB log error: {e}")


# ---------------------------------------------------------------------------
# Build the elevated helper script (runs as SYSTEM via PsExec)
# ---------------------------------------------------------------------------
def build_helper_script(rdp_labels, run_seed, claimed_sids=None, assignments=None):
    """Write a Python script that SYSTEM runs to find sessions and inject clean_bot.

    Args:
        rdp_labels: list of labels like ['RDP1']
        run_seed: hex run identifier
        claimed_sids: session IDs already used by previous RDPs
        assignments: list of (label, rdp_key, role, location) tuples.
                     If None, defaults to inbound + first DB location.
    """
    claimed_sids = claimed_sids or []
    assignments = assignments or []
    # Build lookup from label -> (role, location) for .bat generation
    _assign_map = {a[0]: (a[2], a[3]) for a in assignments}

    # Find clean_bot.py
    shared_root = Path(r"C:\MHTAgentic")
    bot_script = shared_root / "clean_bot.py"
    if not bot_script.exists():
        bot_script = PROJECT_ROOT / "clean_bot.py"

    # Find launcher.pyw
    launcher_script = shared_root / "launcher.pyw"
    if not launcher_script.exists():
        launcher_script = PROJECT_ROOT / "launcher.pyw"

    # Find python.exe — prefer system-wide install so all users can access it
    system_python = _find_system_python()
    if "Program Files" in str(system_python):
        python_exe = system_python
    else:
        python_dir = Path(sys.executable).parent
        python_exe = python_dir / "python.exe"
        if not python_exe.exists():
            python_exe = system_python

    # Build PYTHONPATH — only include paths accessible by all users
    pd_sp = r"C:\ProgramData\MHTAgentic\site-packages"
    extra_paths = [pd_sp, pd_sp + r"\win32", pd_sp + r"\win32\lib", pd_sp + r"\Pythonwin"]
    for p in sys.path:
        if "site-packages" in p and p not in extra_paths:
            # Skip per-user paths — agent can't access orchestrator's profile
            if "\\Users\\" in p and "\\Users\\Public" not in p:
                continue
            extra_paths.append(p)
    # Include system-wide Python's site-packages
    sys_sp = str(python_exe.parent / "Lib" / "site-packages")
    if sys_sp not in extra_paths:
        extra_paths.append(sys_sp)
    pythonpath_str = ";".join(extra_paths)

    pywin32_dll_dir = str(shared_root)
    pywin32_sys32 = r"C:\ProgramData\MHTAgentic\site-packages\pywin32_system32"

    # Write per-label wrapper .bat files
    bot_labels = repr(rdp_labels)  # e.g. ['RDP1', 'RDP2', 'RDP3']
    claimed_repr = repr(claimed_sids)  # e.g. [34, 35]

    helper_script = r'''
import ctypes, ctypes.wintypes, sys, time, os

kernel32 = ctypes.windll.kernel32
advapi32 = ctypes.windll.advapi32
wtsapi32 = ctypes.windll.Wtsapi32
userenv = ctypes.windll.userenv

log = open(r"C:\ProgramData\MHTAgentic\clean_launch.log", "w")
def logmsg(msg):
    import datetime
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log.write(f"[{ts}] {msg}\n")
    log.flush()
    print(msg, flush=True)

logmsg("=== clean_bot CreateProcessAsUser launcher ===")

# Enable privileges
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
    tp.Privs[i].Attr = 2
advapi32.AdjustTokenPrivileges(tok, False, ctypes.byref(tp), 0, None, None)
kernel32.CloseHandle(tok)
logmsg("Privileges adjusted")

CLAIMED_SIDS = set(''' + claimed_repr + r''')
logmsg(f"Claimed session IDs to skip: {CLAIMED_SIDS}")

# Enumerate RDP sessions
class WTS_SESSION_INFO(ctypes.Structure):
    _fields_ = [
        ("SessionId", ctypes.wintypes.DWORD),
        ("pWinStationName", ctypes.wintypes.LPWSTR),
        ("State", ctypes.wintypes.DWORD),
    ]

SKIP_USERS = {"administrator", "orchestrator"}

def _get_session_username(sid):
    """Get the username for a given session ID via WTSQuerySessionInformationW."""
    WTS_USERNAME = 5
    buf = ctypes.wintypes.LPWSTR()
    size = ctypes.wintypes.DWORD()
    ok = wtsapi32.WTSQuerySessionInformationW(0, sid, WTS_USERNAME, ctypes.byref(buf), ctypes.byref(size))
    if ok and buf.value:
        username = buf.value
        wtsapi32.WTSFreeMemory(buf)
        return username
    return ""

def find_live_sessions():
    """Brute-force probe sessions with WTSQueryUserToken.
    RDP Wrapper may not update WTS state, so try ALL session IDs."""
    WTSQueryUserToken2 = wtsapi32.WTSQueryUserToken
    WTSQueryUserToken2.argtypes = [ctypes.wintypes.ULONG, ctypes.POINTER(ctypes.wintypes.HANDLE)]
    WTSQueryUserToken2.restype = ctypes.wintypes.BOOL

    pInfo2 = ctypes.POINTER(WTS_SESSION_INFO)()
    count2 = ctypes.wintypes.DWORD()
    wtsapi32.WTSEnumerateSessionsW(0, 0, 1, ctypes.byref(pInfo2), ctypes.byref(count2))
    candidate_sids = set()
    for i in range(count2.value):
        s = pInfo2[i]
        sid = s.SessionId
        if sid not in (0, 1, 65536):
            candidate_sids.add(sid)
    wtsapi32.WTSFreeMemory(pInfo2)

    for sid in range(2, 61):
        candidate_sids.add(sid)

    found = []
    for sid in sorted(candidate_sids):
        if sid in CLAIMED_SIDS:
            continue
        tok = ctypes.wintypes.HANDLE()
        ok = WTSQueryUserToken2(sid, ctypes.byref(tok))
        if ok:
            kernel32.CloseHandle(tok)
            username = _get_session_username(sid)
            if username.lower() in SKIP_USERS:
                logmsg(f"  SKIP session {sid} (user={username})")
                continue
            logmsg(f"  LIVE session {sid} (user={username})")
            found.append({"session_id": sid, "state": 0})
        else:
            err = kernel32.GetLastError()
            if err not in (1008, 7022, 5):
                logmsg(f"  session {sid} token err={err}")
    return found

RDP_LABELS = ''' + bot_labels + r'''
RUN_SEED = "''' + run_seed + r'''"
BOT_SCRIPT = r"''' + str(bot_script) + r'''"
PYTHON_EXE = r"''' + str(python_exe) + r'''"
PYTHONPATH = r"''' + pythonpath_str + r'''"
PYWIN32_DLL = r"''' + pywin32_sys32 + r'''"
PYWIN32_DIR = r"''' + pywin32_dll_dir + r'''"
ASSIGN_MAP = ''' + repr(_assign_map) + r'''

# STARTUPINFO / PROCESS_INFORMATION for CreateProcessAsUser
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

def launch_in_session(label, sid):
    """Write wrapper .bat and launch clean_bot in the given session."""
    logmsg(f"Launching {label} in session {sid}...")

    # Write session ID marker
    sid_file = os.path.join(r"C:\ProgramData\MHTAgentic", f"clean_session_{label}.txt")
    with open(sid_file, "w") as f:
        f.write(str(sid))

    # Clear old status file
    status_file = os.path.join(r"C:\ProgramData\MHTAgentic", f"clean_status_{label}.txt")
    if os.path.exists(status_file):
        os.remove(status_file)

    # Write wrapper .bat
    stderr_log = os.path.join(r"C:\ProgramData\MHTAgentic", f"clean_bot_{label}_stderr.log")
    wrapper_bat = os.path.join(r"C:\ProgramData\MHTAgentic", f"clean_bot_{label}.bat")
    role, location = ASSIGN_MAP.get(label, ("inbound", "ANNISTON"))
    with open(wrapper_bat, "w") as f:
        f.write(
            f"@echo off\r\n"
            f"set PYTHONPATH={PYTHONPATH}\r\n"
            f"set PATH={PYWIN32_DLL};{PYWIN32_DIR};%PATH%\r\n"
            f"set PYTHONUTF8=1\r\n"
            f"set PYTHONNOUSERSITE=1\r\n"
            f"set MHT_RDP_LABEL={label}\r\n"
            f"set MHT_RUN_SEED={RUN_SEED}\r\n"
            f"set BOT_LOCATION={location}\r\n"
            f"set BOT_ROLE={role}\r\n"
            f"set MHT_BOT_ROLE={role}\r\n"
            f"set MHT_BOT_LOCATION={location}\r\n"
            f'"{PYTHON_EXE}" -u "{BOT_SCRIPT}" 2>> "{stderr_log}"\r\n'
        )

    # Get session user token
    user_token = ctypes.wintypes.HANDLE()
    if not WTSQueryUserToken(sid, ctypes.byref(user_token)):
        err = kernel32.GetLastError()
        logmsg(f"  ERROR: WTSQueryUserToken failed for session {sid}, error={err}")
        return False

    # Duplicate to primary token
    dup_token = ctypes.wintypes.HANDLE()
    desired = TOKEN_QUERY | TOKEN_DUPLICATE | TOKEN_ASSIGN_PRIMARY | TOKEN_ADJUST_DEFAULT | TOKEN_ADJUST_SESSIONID
    if not DuplicateTokenEx(user_token, desired, None, SecurityImpersonation, TokenPrimary, ctypes.byref(dup_token)):
        err = kernel32.GetLastError()
        logmsg(f"  ERROR: DuplicateTokenEx failed, error={err}")
        kernel32.CloseHandle(user_token)
        return False

    # Create environment block
    env_block = ctypes.c_void_p()
    CreateEnvironmentBlock(ctypes.byref(env_block), dup_token, False)

    # STARTUPINFO — minimize console to prevent focus stealing
    si = STARTUPINFOW()
    si.cb = ctypes.sizeof(STARTUPINFOW)
    si.lpDesktop = "winsta0\\default"
    si.dwFlags = 0x00000001  # STARTF_USESHOWWINDOW
    si.wShowWindow = 7       # SW_SHOWMINNOACTIVE

    pi = PROCESS_INFORMATION()

    cmd_line = f'cmd.exe /c "{wrapper_bat}"'
    ok = CreateProcessAsUserW(
        dup_token, None, cmd_line,
        None, None, False,
        CREATE_NEW_CONSOLE | CREATE_UNICODE_ENVIRONMENT,
        env_block, None,
        ctypes.byref(si), ctypes.byref(pi),
    )

    if ok:
        logmsg(f"  OK: {label} started in session {sid}, PID={pi.dwProcessId}")
        kernel32.CloseHandle(pi.hProcess)
        kernel32.CloseHandle(pi.hThread)
    else:
        err = kernel32.GetLastError()
        logmsg(f"  ERROR: CreateProcessAsUserW failed, error={err}")

    if env_block:
        DestroyEnvironmentBlock(env_block)
    kernel32.CloseHandle(dup_token)
    kernel32.CloseHandle(user_token)
    return bool(ok)

# =====================================================================
# Main loop: process labels ONE AT A TIME, triggered by orchestrator
# =====================================================================
logmsg(f"Waiting for triggers for {len(RDP_LABELS)} label(s): {RDP_LABELS}")

for label in RDP_LABELS:
    trigger_file = os.path.join(r"C:\ProgramData\MHTAgentic", f"launch_trigger_{label}.txt")

    # Wait for orchestrator to write the trigger (RDP is open and ready)
    logmsg(f"[{label}] Waiting for trigger: {trigger_file}")
    for wait_attempt in range(600):  # up to 10 minutes per label
        if os.path.exists(trigger_file):
            break
        time.sleep(1)
    else:
        logmsg(f"[{label}] TIMEOUT waiting for trigger (10 min)")
        continue

    # Remove trigger
    try:
        os.remove(trigger_file)
    except Exception:
        pass
    logmsg(f"[{label}] Trigger received, finding new session...")

    # Find ONE new unclaimed session (retry up to 90s)
    found_sid = None
    for attempt in range(30):
        sessions = find_live_sessions()
        if sessions:
            found_sid = sessions[0]["session_id"]
            logmsg(f"[{label}] Found session {found_sid}")
            break
        logmsg(f"[{label}] No new session yet (attempt {attempt+1}/30)...")
        time.sleep(3)

    if found_sid is None:
        logmsg(f"[{label}] ERROR: No session found after 90s")
        continue

    # Launch bot in this session
    ok = launch_in_session(label, found_sid)
    if ok:
        CLAIMED_SIDS.add(found_sid)
        logmsg(f"[{label}] Claimed session {found_sid}")
    else:
        logmsg(f"[{label}] Launch failed")

logmsg("=== All labels processed ===")
log.close()
'''

    helper_path = DATA_DIR / "clean_launch_helper.py"
    with open(helper_path, "w", encoding="utf-8") as f:
        f.write(helper_script)

    return helper_path


# ---------------------------------------------------------------------------
# Monitor status files
# ---------------------------------------------------------------------------
def monitor_status(rdp_label, run_seed, timeout=300):
    """Poll the per-RDP status file and log transitions to SQLite.

    Waits until status 160 (location confirmed) or a terminal code (-1, -2).
    Login phase: 10-110.  Post-login phase: 120-160.
    After 160, monitoring continues in background inside the RDP session.
    """
    status_file = DATA_DIR / f"clean_status_{rdp_label}.txt"
    session_file = DATA_DIR / f"clean_session_{rdp_label}.txt"
    deadline = time.time() + timeout
    last_status = None
    last_status_time = time.time()
    session_id = None

    while time.time() < deadline:
        # Pick up session ID from helper
        if session_id is None and session_file.exists():
            try:
                session_id = int(session_file.read_text().strip())
                print(f"  Session: {session_id}")
            except Exception:
                pass

        if status_file.exists():
            try:
                content = status_file.read_text().strip()
                code = int(content.split()[0])
                msg = content[len(str(code)):].strip()

                if code != last_status:
                    now = time.time()
                    duration_ms = int((now - last_status_time) * 1000) if last_status is not None else None
                    last_status_time = now

                    label = STATUS_LABELS.get(code, f"Status {code}")
                    extra = f" -- {msg}" if msg and msg != label else ""
                    print(f"  [{code}] {label}{extra}")

                    is_complete = 1 if code in (160, 200) else 0
                    _log_status(
                        run_seed, rdp_label, session_id,
                        code, msg or label,
                        duration_ms=duration_ms,
                        complete=is_complete,
                    )
                    last_status = code

                # Terminal codes: 160 (location confirmed), 200 (monitoring active),
                # -1 (error), -2 (already logged in -> still chain to launcher)
                if code in (160, 200, -1, -2):
                    return code
            except Exception:
                pass

        time.sleep(1)

    print("  TIMEOUT: bot did not finish in time")
    _log_status(run_seed, rdp_label, session_id, -1, "Timeout", complete=0)
    return -1


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def _dismiss_mstsc_dialogs(rdp_key, timeout=30):
    """Wait for mstsc to connect, sending Enter to dismiss any session-select dialogs.

    mstsc may show a 'select session to reconnect' dialog when there are
    disconnected sessions. This function detects those dialogs and sends Enter
    to pick the default and continue connecting.
    """
    import ctypes.wintypes

    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    _user32 = ctypes.windll.user32

    def _get_class(hwnd):
        buf = ctypes.create_unicode_buffer(256)
        _user32.GetClassNameW(hwnd, buf, 256)
        return buf.value

    def _get_title(hwnd):
        length = _user32.GetWindowTextLengthW(hwnd) + 1
        buf = ctypes.create_unicode_buffer(length)
        _user32.GetWindowTextW(hwnd, buf, length)
        return buf.value

    deadline = time.time() + timeout
    enter_sent = 0
    grace_until = time.time() + 8  # let mstsc try to connect on its own first

    while time.time() < deadline:
        # Check if the RDP session connected (window title has the rdp_key name)
        mstsc_windows = []
        connected = False

        def _enum(hwnd, _lp):
            nonlocal connected
            if not _user32.IsWindowVisible(hwnd):
                return True
            if _get_class(hwnd) == "TscShellContainerClass":
                title = _get_title(hwnd)
                mstsc_windows.append((hwnd, title))
                if rdp_key.lower() in title.lower():
                    connected = True
            return True

        _user32.EnumWindows(WNDENUMPROC(_enum), 0)

        if connected:
            print(f"  RDP connected ({rdp_key})")
            return True

        # After grace period, send Enter to dismiss dialogs
        if time.time() > grace_until and enter_sent < 5:
            for hwnd, title in mstsc_windows:
                # Dialog windows say "Remote Desktop Connection" without the server name
                if "remote desktop connection" in title.lower():
                    _user32.SetForegroundWindow(hwnd)
                    time.sleep(0.3)
                    # Send Enter via PostMessage (WM_KEYDOWN + WM_KEYUP for VK_RETURN)
                    VK_RETURN = 0x0D
                    _user32.PostMessageW(hwnd, 0x0100, VK_RETURN, 0)  # WM_KEYDOWN
                    time.sleep(0.05)
                    _user32.PostMessageW(hwnd, 0x0101, VK_RETURN, 0)  # WM_KEYUP
                    enter_sent += 1
                    print(f"  Sent Enter to dismiss dialog: '{title}'")
                    grace_until = time.time() + 5  # wait 5s before next attempt
                    break

        time.sleep(2)

    print(f"  WARNING: mstsc did not connect after {timeout}s")
    return False


def _build_assignments(rdp_count, inbound_count, outbound_count, locations):
    """Build assignment plan: (label, rdp_key, role, location) tuples."""
    rdp_entries = ALL_RDPS[:rdp_count]
    assignments = []
    loc_idx = 0
    for idx, (label, rdp_key) in enumerate(rdp_entries):
        if idx < inbound_count:
            role = "inbound"
            loc = locations[loc_idx]["location_name"] if loc_idx < len(locations) else "ANNISTON"
            loc_idx += 1
        else:
            role = "outbound"
            loc = locations[0]["location_name"] if locations else "ANNISTON"
        assignments.append((label, rdp_key, role, loc))
    return assignments


def main():
    if not PSEXEC_PATH.exists():
        print(f"ERROR: PsExec64.exe not found at {PSEXEC_PATH}")
        sys.exit(1)

    run_seed = os.urandom(6).hex()
    _init_db()

    # --- Read dynamic config from dashboard DB ---
    # Active locations drive the count: 1 inbound RDP per location + 1 outbound RDP
    config = get_config(str(DB_PATH))
    all_locs = get_all_locations(str(DB_PATH))
    locations = [loc for loc in all_locs if loc.get("is_active")]
    if not locations:
        print("WARNING: No active locations — falling back to all locations")
        locations = all_locs

    inbound_count = len(locations)
    outbound_count = 1
    rdp_count = inbound_count + outbound_count

    # Generate .rdp files dynamically — no hardcoded limit
    rdp_slots = _generate_rdp_files(rdp_count)
    print(f"  Generated {len(rdp_slots)} RDP files ({inbound_count} inbound + {outbound_count} outbound)")

    # Populate global dicts for backward compatibility
    global ALL_RDPS, RDP_FILES
    ALL_RDPS = [(label, key) for label, key, _ in rdp_slots]
    RDP_FILES = {key: path for _, key, path in rdp_slots}

    release_all_locations(str(DB_PATH))

    # --- Parse --skip <sid,sid,...> for reboot mode (skip other active sessions) ---
    skip_sids = []
    args = sys.argv[1:]
    if "--skip" in args:
        skip_idx = args.index("--skip")
        if skip_idx + 1 < len(args):
            skip_sids = [int(x) for x in args[skip_idx + 1].split(",") if x.strip()]
            args = args[:skip_idx] + args[skip_idx + 2:]
        else:
            args = args[:skip_idx]

    # --- Parse --role <role> for reboot mode ---
    reboot_role = None
    if "--role" in args:
        role_idx = args.index("--role")
        if role_idx + 1 < len(args):
            reboot_role = args[role_idx + 1]
            args = args[:role_idx] + args[role_idx + 2:]
        else:
            args = args[:role_idx]

    # --- Determine assignment plan ---
    if args:
        # Single RDP override
        rdp_key = args[0]
        match = [(label, key) for label, key in ALL_RDPS if key == rdp_key]
        if not match:
            match = [("RDP1", rdp_key)]
        label, key = match[0]
        role = reboot_role or "inbound"
        loc = locations[0]["location_name"] if locations else "ANNISTON"
        assignments = [(label, key, role, loc)]
    else:
        assignments = _build_assignments(rdp_count, inbound_count, outbound_count, locations)

    total = len(assignments)
    print(f"=== Multi-RDP Clean Startup ===")
    print(f"  Run seed: {run_seed}")
    print(f"  Config: rdp_count={rdp_count}, inbound={inbound_count}, outbound={outbound_count}")
    print(f"  Locations: {[l['location_name'] for l in locations]}")
    print(f"  Assignment plan:")
    for label, rdp_key, role, loc in assignments:
        print(f"    {label} ({rdp_key}) -> {role} @ {loc}")
    print()

    # Find python.exe (used for each helper launch)
    python_dir = Path(sys.executable).parent
    python_exe = python_dir / "python.exe"
    if not python_exe.exists():
        python_exe = _find_system_python()

    shell32 = ctypes.windll.shell32
    if skip_sids:
        print(f"  Skipping session IDs: {skip_sids}")
    results = {}

    # =================================================================
    # Phase 1: Clean old status/trigger files
    # =================================================================
    for label, rdp_key, role, loc in assignments:
        for f in [DATA_DIR / f"clean_status_{label}.txt",
                  DATA_DIR / f"clean_session_{label}.txt",
                  DATA_DIR / f"launch_trigger_{label}.txt"]:
            if f.exists():
                f.unlink()

    # =================================================================
    # Phase 2: Build ONE helper for ALL RDPs, ONE PsExec/UAC call
    # The helper stays running and waits for trigger files per-label.
    # =================================================================
    all_labels = [label for label, _, _, _ in assignments]
    print(f"  Launching elevated helper for {len(all_labels)} RDP(s) (single UAC prompt)...")
    helper_path = build_helper_script(all_labels, run_seed, list(skip_sids),
                                      assignments=assignments)

    psexec_bat = DATA_DIR / "clean_psexec_launcher.bat"
    with open(psexec_bat, "w") as f:
        f.write(
            f'@echo off\r\n'
            f'"{PSEXEC_PATH}" -accepteula -nobanner -s -d '
            f'"{python_exe}" -u "{helper_path}"\r\n'
        )

    uac_result = shell32.ShellExecuteW(None, "runas", str(psexec_bat), None, None, 0)
    if uac_result <= 32:
        print("  ERROR: UAC denied or ShellExecute failed")
        sys.exit(1)
    print(f"  Helper started (waiting for triggers)\n")

    # Give PsExec a moment to start the helper
    time.sleep(3)

    # =================================================================
    # Phase 3: Open RDPs sequentially (fuse) — trigger helper per label
    # =================================================================
    for idx, (label, rdp_key, role, loc) in enumerate(assignments, 1):
        print(f"[{idx}/{total}] {label} ({rdp_key}) — {role} @ {loc}")
        print(f"  {'='*40}")
        start_time = time.time()

        rdp_file = RDP_FILES.get(rdp_key)
        if not rdp_file or not rdp_file.exists():
            print(f"  ERROR: {rdp_key} .rdp file not found")
            results[label] = "FAILED"
            continue

        # Open RDP and wait for connection (dismiss session-select dialogs)
        print(f"  Opening {rdp_key} ({rdp_file.name})...")
        subprocess.Popen(["mstsc.exe", str(rdp_file)], creationflags=subprocess.DETACHED_PROCESS)
        time.sleep(3)
        _dismiss_mstsc_dialogs(rdp_key, timeout=30)

        # Signal the helper to find the new session and launch the bot
        trigger_file = DATA_DIR / f"launch_trigger_{label}.txt"
        trigger_file.write_text("go")
        print(f"  Trigger written — helper will find session and launch bot")

        # Monitor until location confirmed (160) or terminal code
        print(f"  Monitoring (login -> EMR -> location)...")
        code = monitor_status(label, run_seed)
        elapsed = time.time() - start_time

        # Formally claim the location slot in DB after status 160 or 200
        if code in (160, 200):
            assigned = assign_location(str(DB_PATH), rdp_key, role)
            if assigned:
                print(f"  DB: assigned {rdp_key} -> {assigned} ({role})")
            else:
                print(f"  DB: no available location to assign (may already be taken)")

        success = code in (160, 200, -2)
        status = "SUCCESS" if success else "FAILED"
        results[label] = status
        print(f"  Result: {status} ({elapsed:.1f}s)")
        print()

    # Summary
    print(f"{'='*50}")
    print(f"  Summary (seed: {run_seed})")
    print(f"{'='*50}")
    for label, status in results.items():
        marker = "OK" if status == "SUCCESS" else "FAIL"
        print(f"  [{marker}] {label}: {status}")

    failed = sum(1 for s in results.values() if s != "SUCCESS")
    if failed:
        print(f"\n  {failed}/{total} RDP(s) failed.")
        sys.exit(1)
    else:
        print(f"\n  All {total} RDP(s) started successfully.")


if __name__ == "__main__":
    main()
