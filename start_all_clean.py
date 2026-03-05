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

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
PROJECT_ROOT = Path(__file__).resolve().parent
PSEXEC_PATH = Path(r"C:\ProgramData\MHTAgentic\PsExec64.exe")
RDP_DIR = Path(r"C:\Users\jaden\Desktop")
DATA_DIR = Path(r"C:\ProgramData\MHTAgentic")
DB_PATH = DATA_DIR / "mht_data.db"

RDP_FILES = {
    "ExperityB": RDP_DIR / "ExperityB_MHT.rdp",
    "ExperityC": RDP_DIR / "ExperityC_MHT.rdp",
    "ExperityD": RDP_DIR / "ExperityD_MHT.rdp",
}

# (label, rdp_file_key)
RDP_LIST = [
    ("RDP1", "ExperityB"),
    ("RDP2", "ExperityC"),
    ("RDP3", "ExperityD"),
]

STATUS_LABELS = {
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
def build_helper_script(rdp_labels, run_seed):
    """Write a Python script that SYSTEM runs to find sessions and inject clean_bot."""
    # Find clean_bot.py
    shared_root = Path(r"C:\MHTAgentic")
    bot_script = shared_root / "clean_bot.py"
    if not bot_script.exists():
        bot_script = PROJECT_ROOT / "clean_bot.py"

    # Find python.exe
    python_dir = Path(sys.executable).parent
    python_exe = python_dir / "python.exe"
    if not python_exe.exists():
        python_exe = Path(r"C:\Program Files\Python39\python.exe")

    # Build PYTHONPATH
    pd_sp = r"C:\ProgramData\MHTAgentic\site-packages"
    extra_paths = [pd_sp, pd_sp + r"\win32", pd_sp + r"\win32\lib", pd_sp + r"\Pythonwin"]
    for p in sys.path:
        if "site-packages" in p and p not in extra_paths:
            extra_paths.append(p)
    pythonpath_str = ";".join(extra_paths)

    pywin32_dll_dir = str(shared_root)
    pywin32_sys32 = r"C:\ProgramData\MHTAgentic\site-packages\pywin32_system32"

    # Write per-label wrapper .bat files
    bot_labels = repr(rdp_labels)  # e.g. ['RDP1', 'RDP2', 'RDP3']

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

# Enumerate RDP sessions
class WTS_SESSION_INFO(ctypes.Structure):
    _fields_ = [
        ("SessionId", ctypes.wintypes.DWORD),
        ("pWinStationName", ctypes.wintypes.LPWSTR),
        ("State", ctypes.wintypes.DWORD),
    ]

WTSActive = 0
WTSConnected = 1
WTSDisconnected = 4
WTSUserName = 5

pInfo = ctypes.POINTER(WTS_SESSION_INFO)()
count = ctypes.wintypes.DWORD()
wtsapi32.WTSEnumerateSessionsW(0, 0, 1, ctypes.byref(pInfo), ctypes.byref(count))

def find_live_sessions():
    """Brute-force probe sessions with WTSQueryUserToken.
    RDP Wrapper may not update WTS state, so try ALL session IDs."""
    WTSQueryUserToken2 = wtsapi32.WTSQueryUserToken
    WTSQueryUserToken2.argtypes = [ctypes.wintypes.ULONG, ctypes.POINTER(ctypes.wintypes.HANDLE)]
    WTSQueryUserToken2.restype = ctypes.wintypes.BOOL

    # Get all session IDs from WTS
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

    # Also try IDs 2-60 in case WTS doesn't enumerate them
    for sid in range(2, 61):
        candidate_sids.add(sid)

    found = []
    for sid in sorted(candidate_sids):
        tok = ctypes.wintypes.HANDLE()
        ok = WTSQueryUserToken2(sid, ctypes.byref(tok))
        if ok:
            kernel32.CloseHandle(tok)
            logmsg(f"  LIVE session {sid}")
            found.append({"session_id": sid, "state": 0})
        else:
            err = kernel32.GetLastError()
            # Log non-trivial errors
            if err not in (1008, 7022, 5):
                logmsg(f"  session {sid} token err={err}")
    return found

# Retry for up to 90 seconds
sessions = []
for attempt in range(30):
    sessions = find_live_sessions()
    if sessions:
        logmsg(f"  Found {len(sessions)} live session(s)")
        break
    logmsg(f"  No live sessions (attempt {attempt+1}/30), waiting 3s...")
    time.sleep(3)

RDP_LABELS = ''' + bot_labels + r'''
RUN_SEED = "''' + run_seed + r'''"
BOT_SCRIPT = r"''' + str(bot_script) + r'''"
PYTHON_EXE = r"''' + str(python_exe) + r'''"
PYTHONPATH = r"''' + pythonpath_str + r'''"
PYWIN32_DLL = r"''' + pywin32_sys32 + r'''"
PYWIN32_DIR = r"''' + pywin32_dll_dir + r'''"

if not sessions:
    logmsg("ERROR: No RDP sessions found!")
    log.close()
    sys.exit(1)

logmsg(f"Found {len(sessions)} RDP session(s), need {len(RDP_LABELS)} label(s)")

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

# Assign labels to sessions (sorted by session ID)
for i, label in enumerate(RDP_LABELS):
    if i >= len(sessions):
        logmsg(f"WARNING: Not enough sessions for {label}")
        continue

    sess = sessions[i]
    sid = sess["session_id"]
    logmsg(f"Launching {label} in session {sid} (state={sess['state']})...")

    # Write status file marker so monitor knows session ID
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
    with open(wrapper_bat, "w") as f:
        f.write(
            f"@echo off\r\n"
            f"set PYTHONPATH={PYTHONPATH}\r\n"
            f"set PATH={PYWIN32_DLL};{PYWIN32_DIR};%PATH%\r\n"
            f"set PYTHONUTF8=1\r\n"
            f"set PYTHONNOUSERSITE=1\r\n"
            f"set MHT_RDP_LABEL={label}\r\n"
            f"set MHT_RUN_SEED={RUN_SEED}\r\n"
            f'"{PYTHON_EXE}" -u "{BOT_SCRIPT}" 2>> "{stderr_log}"\r\n'
        )

    # Get session user token
    user_token = ctypes.wintypes.HANDLE()
    if not WTSQueryUserToken(sid, ctypes.byref(user_token)):
        err = kernel32.GetLastError()
        logmsg(f"  ERROR: WTSQueryUserToken failed for session {sid}, error={err}")
        continue

    # Duplicate to primary token
    dup_token = ctypes.wintypes.HANDLE()
    desired = TOKEN_QUERY | TOKEN_DUPLICATE | TOKEN_ASSIGN_PRIMARY | TOKEN_ADJUST_DEFAULT | TOKEN_ADJUST_SESSIONID
    if not DuplicateTokenEx(user_token, desired, None, SecurityImpersonation, TokenPrimary, ctypes.byref(dup_token)):
        err = kernel32.GetLastError()
        logmsg(f"  ERROR: DuplicateTokenEx failed, error={err}")
        kernel32.CloseHandle(user_token)
        continue

    # Create environment block
    env_block = ctypes.c_void_p()
    CreateEnvironmentBlock(ctypes.byref(env_block), dup_token, False)

    # STARTUPINFO targeting session desktop
    si = STARTUPINFOW()
    si.cb = ctypes.sizeof(STARTUPINFOW)
    si.lpDesktop = "winsta0\\default"

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

logmsg("=== All launches done ===")
log.close()
'''

    helper_path = DATA_DIR / "clean_launch_helper.py"
    with open(helper_path, "w", encoding="utf-8") as f:
        f.write(helper_script)

    return helper_path


# ---------------------------------------------------------------------------
# Monitor status files
# ---------------------------------------------------------------------------
def monitor_status(rdp_label, run_seed, timeout=180):
    """Poll the per-RDP status file and log transitions to SQLite."""
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

                    label = STATUS_LABELS.get(code, "Unknown")
                    extra = f" -- {msg}" if msg and msg != label else ""
                    print(f"  [{code}] {label}{extra}")

                    is_complete = 1 if code in (110, -2) else 0
                    _log_status(
                        run_seed, rdp_label, session_id,
                        code, msg or label,
                        duration_ms=duration_ms,
                        complete=is_complete,
                    )
                    last_status = code

                if code in (110, -1, -2):
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
def main():
    if not PSEXEC_PATH.exists():
        print(f"ERROR: PsExec64.exe not found at {PSEXEC_PATH}")
        sys.exit(1)

    run_seed = os.urandom(6).hex()
    _init_db()

    # Determine which RDPs to process
    if len(sys.argv) > 1:
        rdp_key = sys.argv[1]
        rdp_entries = [(label, key) for label, key in RDP_LIST if key == rdp_key]
        if not rdp_entries:
            rdp_entries = [("RDP1", rdp_key)]
    else:
        rdp_entries = RDP_LIST

    total = len(rdp_entries)
    print(f"=== Multi-RDP Clean Startup ===")
    print(f"  Run seed: {run_seed}")
    print(f"  RDPs to process: {total}")
    print()

    # Step 1: Open all RDP connections
    print("Step 1: Opening RDP connections...")
    for label, rdp_key in rdp_entries:
        rdp_file = RDP_FILES.get(rdp_key)
        if not rdp_file or not rdp_file.exists():
            print(f"  ERROR: {rdp_key} .rdp file not found")
            continue
        print(f"  Opening {rdp_key} ({rdp_file.name})...")
        subprocess.Popen(["mstsc.exe", str(rdp_file)], creationflags=subprocess.DETACHED_PROCESS)
        time.sleep(3)  # Small gap between RDP opens

    # Step 2: Clean old status files
    for label, _ in rdp_entries:
        for f in [DATA_DIR / f"clean_status_{label}.txt", DATA_DIR / f"clean_session_{label}.txt"]:
            if f.exists():
                f.unlink()

    # Step 3: Build and launch the elevated helper (it retries session detection internally)
    rdp_labels = [label for label, _ in rdp_entries]
    print(f"\nStep 2: Launching clean_bot in sessions (1 UAC prompt)...")
    helper_path = build_helper_script(rdp_labels, run_seed)

    # Find python.exe for the helper
    python_dir = Path(sys.executable).parent
    python_exe = python_dir / "python.exe"
    if not python_exe.exists():
        python_exe = Path(r"C:\Program Files\Python39\python.exe")

    # PsExec -s runs helper as SYSTEM (needed for WTSQueryUserToken)
    psexec_bat = DATA_DIR / "clean_psexec_launcher.bat"
    with open(psexec_bat, "w") as f:
        f.write(
            f'@echo off\r\n'
            f'"{PSEXEC_PATH}" -accepteula -nobanner -s -d '
            f'"{python_exe}" -u "{helper_path}"\r\n'
        )

    shell32 = ctypes.windll.shell32
    result = shell32.ShellExecuteW(None, "runas", str(psexec_bat), None, None, 0)
    if result <= 32:
        print("  ERROR: UAC denied or ShellExecute failed")
        sys.exit(1)

    print("  Helper launched (will retry until sessions appear)...")

    # Step 3: Monitor each RDP's status file
    print(f"\nStep 3: Monitoring progress...")
    results = {}
    for label, rdp_key in rdp_entries:
        print(f"\n[{label}] ({rdp_key})")
        start_time = time.time()
        code = monitor_status(label, run_seed)
        elapsed = time.time() - start_time

        success = code in (110, -2)
        status = "SUCCESS" if success else "FAILED"
        results[label] = status
        print(f"  Result: {status} ({elapsed:.1f}s)")

    # Summary
    print(f"\n{'='*50}")
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
