"""
Session isolation guard for multi-RDP deployments.

Ensures window operations only affect windows in the current Windows Session,
preventing cross-session interference when multiple bot instances run on the
same machine under different user accounts (ExperityB, ExperityC).
"""

import ctypes
import ctypes.wintypes
import logging
import os

logger = logging.getLogger("mhtagentic.session_guard")

kernel32 = ctypes.windll.kernel32
user32 = ctypes.windll.user32


def get_current_session_id() -> int:
    """Get the Windows Session ID for the current process."""
    session_id = ctypes.wintypes.DWORD()
    pid = kernel32.GetCurrentProcessId()
    kernel32.ProcessIdToSessionId(pid, ctypes.byref(session_id))
    return session_id.value


def _get_window_session_id(hwnd) -> int:
    """Get the Session ID that owns a given window handle."""
    pid = ctypes.wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    session_id = ctypes.wintypes.DWORD()
    kernel32.ProcessIdToSessionId(pid.value, ctypes.byref(session_id))
    return session_id.value


def is_current_session_window(hwnd) -> bool:
    """Check if a window belongs to the current session."""
    try:
        return _get_window_session_id(hwnd) == get_current_session_id()
    except Exception:
        return False


def is_rdp_session() -> bool:
    """Return True if the current process is running in an RDP session (not console)."""
    try:
        session_id = get_current_session_id()
        # Session 0 is services, Session 1 is typically the console.
        # More reliable: query the WinStation name via WTS API.
        wts = ctypes.windll.Wtsapi32
        buf = ctypes.wintypes.LPWSTR()
        size = ctypes.wintypes.DWORD()
        WTSClientProtocolType = 16
        if wts.WTSQuerySessionInformationW(
            0, session_id, WTSClientProtocolType, ctypes.byref(buf), ctypes.byref(size)
        ):
            # Protocol: 0 = console, 2 = RDP
            protocol = ctypes.cast(buf, ctypes.POINTER(ctypes.c_ushort)).contents.value
            wts.WTSFreeMemory(buf)
            return protocol == 2
    except Exception:
        pass
    # Fallback: check SESSIONNAME environment variable
    return os.environ.get("SESSIONNAME", "Console").upper().startswith("RDP")


# ---------------------------------------------------------------------------
# Drop-in replacements for cross-session-unsafe window search functions
# ---------------------------------------------------------------------------

def session_find_elements(**kwargs):
    """Drop-in replacement for pywinauto.findwindows.find_elements() that
    only returns elements whose top-level window belongs to the current session."""
    from pywinauto import findwindows
    my_session = get_current_session_id()
    elements = findwindows.find_elements(**kwargs)
    filtered = []
    for elem in elements:
        try:
            if _get_window_session_id(elem.handle) == my_session:
                filtered.append(elem)
        except Exception:
            pass
    return filtered


def session_get_all_windows():
    """Drop-in replacement for pygetwindow.getAllWindows() — current session only."""
    import pygetwindow as gw
    my_session = get_current_session_id()
    return [w for w in gw.getAllWindows() if _get_window_session_id(w._hWnd) == my_session]


def session_get_windows_with_title(title: str):
    """Drop-in replacement for pygetwindow.getWindowsWithTitle() — current session only."""
    import pygetwindow as gw
    my_session = get_current_session_id()
    return [w for w in gw.getWindowsWithTitle(title) if _get_window_session_id(w._hWnd) == my_session]


def session_desktop_windows(backend='uia'):
    """Drop-in replacement for Desktop(backend=...).windows() — current session only."""
    from pywinauto import Desktop
    my_session = get_current_session_id()
    desktop = Desktop(backend=backend)
    filtered = []
    for w in desktop.windows():
        try:
            if _get_window_session_id(w.handle) == my_session:
                filtered.append(w)
        except Exception:
            pass
    return filtered


def session_connect(backend='uia', timeout=5, **kwargs):
    """Session-safe Application.connect().

    Instead of letting pywinauto search all sessions, find the handle via
    session_find_elements() first, then connect by handle.

    Accepts the same keyword args as findwindows.find_elements()
    (title, title_re, class_name, etc.).

    Returns (app, win) tuple or raises TimeoutError.
    """
    import time as _time
    from pywinauto import Application

    deadline = _time.time() + timeout
    while True:
        elems = session_find_elements(backend=backend, **kwargs)
        if elems:
            handle = elems[0].handle
            app = Application(backend=backend).connect(handle=handle, timeout=3)
            win = app.window(handle=handle)
            return app, win
        if _time.time() >= deadline:
            raise TimeoutError(
                f"session_connect timed out after {timeout}s — no window matched "
                f"{kwargs} in current session"
            )
        _time.sleep(0.25)


def session_wait_for_window(timeout=30, poll=0.5, backend='uia', **kwargs):
    """Poll until a window matching kwargs appears in the current session.

    Returns (app, win) as soon as the window is found.
    Raises TimeoutError if it doesn't appear within timeout seconds.

    This replaces patterns like:
        time.sleep(N)
        app = Application(...).connect(title=..., timeout=M)
    """
    return session_connect(backend=backend, timeout=timeout, **kwargs)
