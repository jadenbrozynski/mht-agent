"""
terminate_all.py -- Cleanly terminate all MHT bot processes and RDP sessions.

Usage:
    python terminate_all.py              # Kill bots + logoff sessions + close mstsc
    python terminate_all.py --bots-only  # Only kill bot processes, keep sessions alive

Steps:
  1. Kill all bot python.exe processes running in RDP sessions (via wmic)
  2. Log off all RDP sessions via elevated WTSLogoffSession
  3. Kill mstsc.exe to dismiss "session ended" error dialogs
"""

import subprocess
import sys
import time
import ctypes
from ctypes import wintypes, byref, POINTER, Structure


def get_rdp_python_pids():
    """Find all python.exe PIDs running in RDP sessions (not Console)."""
    r = subprocess.run(
        ['tasklist', '/V', '/FI', 'IMAGENAME eq python.exe', '/FO', 'CSV'],
        capture_output=True, text=True
    )
    pids = []
    for line in r.stdout.strip().split('\n')[1:]:  # skip header
        parts = line.replace('"', '').split(',')
        if len(parts) >= 3 and 'RDP' in parts[2]:
            pids.append(int(parts[1]))
    return pids


def kill_bot_processes():
    """Kill all python.exe processes in RDP sessions."""
    pids = get_rdp_python_pids()
    if not pids:
        print("  No bot processes found in RDP sessions")
        return

    print(f"  Found {len(pids)} bot process(es): {pids}")
    for pid in pids:
        r = subprocess.run(
            ['wmic', 'process', 'where', f'ProcessId={pid}', 'call', 'terminate'],
            capture_output=True, text=True, timeout=10
        )
        print(f"  PID {pid}: terminated")

    # Verify
    time.sleep(1)
    remaining = get_rdp_python_pids()
    if remaining:
        print(f"  WARNING: {len(remaining)} process(es) still running: {remaining}")
        # Fallback: elevated taskkill
        subprocess.run(
            ['powershell.exe', '-Command',
             f"Start-Process -FilePath 'taskkill' -ArgumentList '/F {' '.join(f'/PID {p}' for p in remaining)}' -Verb RunAs -Wait"],
            capture_output=True, text=True, timeout=15
        )
    else:
        print("  All bot processes killed")


class WTS_SESSION_INFO(Structure):
    _fields_ = [
        ('SessionId', wintypes.DWORD),
        ('pWinStationName', wintypes.LPWSTR),
        ('State', wintypes.DWORD),
    ]


def get_rdp_session_ids():
    """Enumerate all RDP sessions (not Console, not Services, not Listener)."""
    wtsapi32 = ctypes.WinDLL('wtsapi32', use_last_error=True)
    PWTS_SESSION_INFO = POINTER(WTS_SESSION_INFO)

    pSessionInfo = PWTS_SESSION_INFO()
    count = wintypes.DWORD(0)

    result = wtsapi32.WTSEnumerateSessionsW(
        wintypes.HANDLE(0), 0, 1, byref(pSessionInfo), byref(count)
    )
    sessions = []
    if result:
        states = {0: 'Active', 1: 'Connected', 4: 'Disconnected'}
        for i in range(count.value):
            s = pSessionInfo[i]
            name = s.pWinStationName or ''
            # Skip console, services, and listener sessions
            if s.SessionId in (0, 1, 65536):
                continue
            if 'Listen' in name:
                continue
            state = states.get(s.State, f'State={s.State}')
            sessions.append((s.SessionId, name, state))
        wtsapi32.WTSFreeMemory(pSessionInfo)
    return sessions


def logoff_sessions():
    """Log off all RDP sessions via elevated PowerShell."""
    sessions = get_rdp_session_ids()
    if not sessions:
        print("  No RDP sessions found")
        return

    for sid, name, state in sessions:
        print(f"  Session {sid} ({name}) - {state}")

    # Build the logoff PS script with discovered session IDs
    sids = [str(s[0]) for s in sessions]
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
    for sid in sids:
        ps_code += f"[WTS]::WTSLogoffSession([IntPtr]::Zero, {sid}, $true) | Out-Null\n"

    script_path = r'C:\ProgramData\MHTAgentic\_terminate_logoff.ps1'
    with open(script_path, 'w') as f:
        f.write(ps_code)

    r = subprocess.run(
        ['powershell.exe', '-Command',
         f"Start-Process -Verb RunAs -Wait -FilePath 'powershell.exe' -ArgumentList '-ExecutionPolicy','Bypass','-File','{script_path}'"],
        capture_output=True, text=True, timeout=30
    )

    # Wait for sessions to fully terminate
    time.sleep(2)

    # Verify
    remaining = get_rdp_session_ids()
    if remaining:
        print(f"  WARNING: {len(remaining)} session(s) still active")
    else:
        print("  All RDP sessions logged off")


def kill_mstsc():
    """Kill mstsc.exe processes to dismiss 'session ended' error dialogs."""
    r = subprocess.run(
        ['powershell.exe', '-Command',
         "Get-Process mstsc -ErrorAction SilentlyContinue | Stop-Process -Force"],
        capture_output=True, text=True, timeout=10
    )
    print("  mstsc processes closed")


def main():
    bots_only = '--bots-only' in sys.argv

    print("[1/3] Killing bot processes...")
    kill_bot_processes()

    if bots_only:
        print("\n--bots-only: skipping session logoff")
        return

    print("\n[2/3] Logging off RDP sessions...")
    logoff_sessions()

    print("\n[3/3] Closing RDP client windows...")
    kill_mstsc()

    print("\nDone.")


if __name__ == '__main__':
    main()
