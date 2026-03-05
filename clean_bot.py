"""
clean_bot.py — Runs inside the RDP session via PsExec.
Opens Experity, enters username, clicks Next, then handles Okta login.

Environment variables (set by start_all_clean.py wrapper bat):
    MHT_RDP_LABEL   — e.g. "RDP1", "RDP2", "RDP3"  (default: "RDP1")
    MHT_RUN_SEED    — 12-char hex run identifier

Status codes (written to C:\\ProgramData\\MHTAgentic\\clean_status_<label>.txt):
    10  Opening Experity
    20  Entering username (Screen 1)
    30  Clicking Next (Screen 1)
    40  Waiting for Okta sign-in (Screen 2)
    50  Entering username (Screen 2, if not pre-filled)
    60  Entering password (Screen 2)
    70  Clicking Sign In (Screen 2)
    80  Generating OTP code (Screen 3)
    90  Entering OTP code (Screen 3)
   100  Clicking Verify (Screen 3)
   110  Login complete
    -1  Error
    -2  Skipped (already logged in)
"""

import sys
import os
import time
import subprocess
from pathlib import Path

# Register pywin32 DLL directories BEFORE importing pywinauto
for _dll_dir in [
    r"C:\MHTAgentic",
    r"C:\ProgramData\MHTAgentic\site-packages\pywin32_system32",
    r"C:\ProgramData\MHTAgentic\site-packages\win32",
]:
    if os.path.isdir(_dll_dir):
        try:
            os.add_dll_directory(_dll_dir)
        except (OSError, AttributeError):
            pass

# Also try user site-packages pywin32_system32
for p in sys.path:
    _candidate = os.path.join(p, "pywin32_system32")
    if os.path.isdir(_candidate):
        try:
            os.add_dll_directory(_candidate)
        except (OSError, AttributeError):
            pass

# ---------------------------------------------------------------------------
# Per-RDP env vars
# ---------------------------------------------------------------------------
RDP_LABEL = os.environ.get("MHT_RDP_LABEL", "RDP1")
RUN_SEED = os.environ.get("MHT_RUN_SEED", "")

# Load credentials from .env
_env_path = Path(r"C:\MHTAgentic\config\.env")
if not _env_path.exists():
    _env_path = Path(__file__).resolve().parent / "config" / ".env"

STORED_USERNAME = ""
STORED_PASSWORD = ""
OKTA_TOTP_SECRET = ""
if _env_path.exists():
    for line in _env_path.read_text().splitlines():
        line = line.strip()
        if line.startswith("EXPERITY_USERNAME="):
            STORED_USERNAME = line.split("=", 1)[1]
        elif line.startswith("EXPERITY_PASSWORD="):
            STORED_PASSWORD = line.split("=", 1)[1]
        elif line.startswith("OKTA_TOTP_SECRET="):
            OKTA_TOTP_SECRET = line.split("=", 1)[1]

if not STORED_USERNAME:
    STORED_USERNAME = "JBROZYNSKI@STHRN"

EXPERITY_SHORTCUT = r"C:\Users\Public\Desktop\Experity EMR - PROD.lnk"

# Per-RDP status file
STATUS_FILE = Path(rf"C:\ProgramData\MHTAgentic\clean_status_{RDP_LABEL}.txt")
STATUS_FILE.parent.mkdir(parents=True, exist_ok=True)


def set_status(code, msg=""):
    STATUS_FILE.write_text(f"{code} {msg}".strip())
    seed_tag = f" seed={RUN_SEED}" if RUN_SEED else ""
    print(f"[clean_bot:{RDP_LABEL}{seed_tag}] STATUS {code}: {msg}", flush=True)


def _click_into(element):
    """Click into an element — click_input first, fall back to coordinate click."""
    try:
        element.click_input()
    except Exception:
        try:
            r = element.rectangle()
            import pyautogui
            pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
        except Exception:
            pass
    time.sleep(0.2)


def kill_existing_apps():
    """Kill any leftover Experity/Edge windows in THIS session only."""
    try:
        from mhtagentic.desktop.session_guard import get_current_session_id
        my_session = get_current_session_id()
        # Only kill processes in our session
        import csv, io
        result = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
        )
        targets = {"msedge.exe", "msedgewebview2.exe", "pythonw.exe"}
        for row in csv.reader(io.StringIO(result.stdout)):
            if len(row) >= 3 and row[0].strip('"').lower() in targets:
                pid = row[1].strip('"').strip()
                session = row[2].strip('"').strip()
                if session == str(my_session):
                    subprocess.run(["taskkill", "/F", "/PID", pid], capture_output=True, timeout=5)
    except Exception:
        pass
    time.sleep(2)


def main():
    try:
        # Import session-safe window finder
        from mhtagentic.desktop.session_guard import session_find_elements

        # Kill leftover apps first
        kill_existing_apps()

        # ── STATUS 10: Open Experity ──
        set_status(10, "Opening Experity")

        if not Path(EXPERITY_SHORTCUT).exists():
            set_status(-1, f"Shortcut not found: {EXPERITY_SHORTCUT}")
            return

        subprocess.Popen(["cmd", "/c", "start", "", EXPERITY_SHORTCUT], shell=True)

        # Wait for Experity window to appear (session-safe)
        from pywinauto import Application

        win = None
        for i in range(60):  # 60 seconds
            try:
                elems = session_find_elements(title_re=".*Experity.*", backend="uia")
                if elems:
                    handle = elems[0].handle
                    app = Application(backend="uia").connect(handle=handle, timeout=3)
                    win = app.window(handle=handle)
                    break
            except Exception:
                pass
            time.sleep(1)

        if not win:
            set_status(-1, "Experity window not found after 60s")
            return

        print(f"[clean_bot:{RDP_LABEL}] Experity window found: '{win.window_text()}'", flush=True)

        # ── STATUS 20: Enter username ──
        set_status(20, "Entering username")

        # Wait for edit fields
        edits = []
        for _ in range(15):
            edits = win.descendants(control_type="Edit")
            if edits:
                break
            time.sleep(1)

        if not edits:
            # Graceful skip: Experity is open but no login fields → already logged in
            print(f"[clean_bot:{RDP_LABEL}] No edit fields found — already logged in, skipping", flush=True)
            set_status(-2, "Already logged in")
            return

        field = edits[0]

        # Click into the field
        try:
            field.click_input()
        except Exception:
            try:
                r = field.rectangle()
                import pyautogui
                pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
            except Exception:
                pass
        time.sleep(0.2)

        # Clear and paste username
        import pyperclip
        field.type_keys("^a", pause=0.05)
        time.sleep(0.1)
        pyperclip.copy(STORED_USERNAME)
        time.sleep(0.1)
        field.type_keys("^v", pause=0.05)
        time.sleep(0.3)

        # Verify
        actual = ""
        try:
            actual = field.get_value() or field.window_text() or ""
            actual = actual.strip()
        except Exception:
            pass

        if actual == STORED_USERNAME:
            print(f"[clean_bot:{RDP_LABEL}] Username entered and verified", flush=True)
        else:
            print(f"[clean_bot:{RDP_LABEL}] Username entered (got '{actual}', expected '{STORED_USERNAME}')", flush=True)

        # ── STATUS 30: Click Next ──
        set_status(30, "Clicking Next")
        time.sleep(0.5)

        clicked = False
        for btn in win.descendants(control_type="Button"):
            if btn.window_text() == "Next":
                try:
                    btn.click_input()
                    clicked = True
                except Exception:
                    try:
                        btn.click()
                        clicked = True
                    except Exception:
                        r = btn.rectangle()
                        import pyautogui
                        pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
                        clicked = True
                break

        if not clicked:
            print(f"[clean_bot:{RDP_LABEL}] Next button not found, pressing Enter", flush=True)
            field.type_keys("{ENTER}")

        time.sleep(2)

        # ── STATUS 40: Wait for Okta sign-in (Screen 2) ──
        set_status(40, "Waiting for Okta sign-in")

        win2 = None
        for _ in range(20):
            for pat in [".*Sign In.*", ".*Experity.*", ".*Okta.*"]:
                try:
                    elems2 = session_find_elements(title_re=pat, backend="uia")
                    if elems2:
                        handle2 = elems2[0].handle
                        app2 = Application(backend="uia").connect(handle=handle2, timeout=3)
                        win2 = app2.window(handle=handle2)
                        break
                except Exception:
                    pass
            if win2:
                break
            time.sleep(1)

        if not win2:
            set_status(-1, "Okta sign-in window not found after 20s")
            return

        print(f"[clean_bot:{RDP_LABEL}] Screen 2 found: '{win2.window_text()}'", flush=True)

        # Wait for edit fields on Screen 2
        edits2 = []
        for _ in range(15):
            edits2 = win2.descendants(control_type="Edit")
            if edits2:
                break
            time.sleep(1)

        if not edits2:
            # Graceful skip: Okta screen open but no fields → already authenticated
            print(f"[clean_bot:{RDP_LABEL}] Screen 2: no edit fields — already authenticated, skipping", flush=True)
            set_status(-2, "Already authenticated at Okta")
            return

        print(f"[clean_bot:{RDP_LABEL}] Screen 2: {len(edits2)} edit field(s)", flush=True)

        # ── STATUS 50: Check username / enter if needed ──
        if len(edits2) >= 2:
            # Two fields: first = username, second = password
            username_field = edits2[0]
            password_field = edits2[1]

            # Check if username is pre-filled
            existing_val = ""
            try:
                existing_val = (username_field.get_value() or username_field.window_text() or "").strip()
            except Exception:
                pass

            if existing_val == STORED_USERNAME:
                print(f"[clean_bot:{RDP_LABEL}] Screen 2: username already pre-filled", flush=True)
                set_status(50, "Username pre-filled")
            else:
                set_status(50, "Entering username")
                _click_into(username_field)
                username_field.type_keys("^a", pause=0.05)
                time.sleep(0.1)
                pyperclip.copy(STORED_USERNAME)
                time.sleep(0.1)
                username_field.type_keys("^v", pause=0.05)
                time.sleep(0.3)
                print(f"[clean_bot:{RDP_LABEL}] Screen 2: username entered", flush=True)
        else:
            # Single field — check if it's username or password
            single_field = edits2[0]
            field_name = ""
            try:
                field_name = (single_field.element_info.name or "").lower()
            except Exception:
                pass

            if "pass" in field_name:
                # It's the password field — username was pre-filled on a previous screen
                password_field = single_field
                set_status(50, "Username pre-filled (single password field)")
                print(f"[clean_bot:{RDP_LABEL}] Screen 2: single field is password, username was pre-filled", flush=True)
            else:
                # It's the username field
                set_status(50, "Entering username")
                _click_into(single_field)
                single_field.type_keys("^a", pause=0.05)
                time.sleep(0.1)
                pyperclip.copy(STORED_USERNAME)
                time.sleep(0.1)
                single_field.type_keys("^v", pause=0.05)
                time.sleep(0.3)
                password_field = single_field  # will need to re-find after submit
                print(f"[clean_bot:{RDP_LABEL}] Screen 2: username entered (single field)", flush=True)

        # ── STATUS 60: Enter password ──
        set_status(60, "Entering password")
        _click_into(password_field)
        password_field.type_keys("^a", pause=0.05)
        time.sleep(0.1)
        pyperclip.copy(STORED_PASSWORD)
        time.sleep(0.1)
        password_field.type_keys("^v", pause=0.05)
        time.sleep(0.3)
        print(f"[clean_bot:{RDP_LABEL}] Screen 2: password entered", flush=True)

        # ── STATUS 70: Click Sign In ──
        set_status(70, "Clicking Sign In")
        time.sleep(0.5)

        signed_in = False
        for btn in win2.descendants(control_type="Button"):
            btn_text = btn.window_text()
            if btn_text in ("Sign In", "Sign in"):
                try:
                    btn.click_input()
                    signed_in = True
                except Exception:
                    try:
                        btn.click()
                        signed_in = True
                    except Exception:
                        r = btn.rectangle()
                        import pyautogui
                        pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
                        signed_in = True
                break

        if not signed_in:
            print(f"[clean_bot:{RDP_LABEL}] Sign In button not found, pressing Enter", flush=True)
            password_field.type_keys("{ENTER}")

        time.sleep(3)

        # ── STATUS 80: Generate OTP code (Screen 3) ──
        set_status(80, "Generating OTP code")

        import pyotp
        if not OKTA_TOTP_SECRET:
            set_status(-1, "OKTA_TOTP_SECRET not set in .env")
            return

        # Wait for the verify/OTP screen (session-safe)
        win3 = None
        for _ in range(20):
            for pat in [".*Verify.*", ".*Sign In.*", ".*Okta.*", ".*Authentication.*"]:
                try:
                    elems3 = session_find_elements(title_re=pat, backend="uia")
                    if elems3:
                        handle3 = elems3[0].handle
                        app3 = Application(backend="uia").connect(handle=handle3, timeout=3)
                        win3 = app3.window(handle=handle3)
                        break
                except Exception:
                    pass
            if win3:
                break
            time.sleep(1)

        if not win3:
            set_status(-1, "OTP/Verify window not found after 20s")
            return

        print(f"[clean_bot:{RDP_LABEL}] Screen 3 found: '{win3.window_text()}'", flush=True)

        # Click "Google Authenticator" link if visible
        try:
            for link in win3.descendants(control_type="Hyperlink"):
                if "google" in link.window_text().lower():
                    link.click_input()
                    print(f"[clean_bot:{RDP_LABEL}] Clicked Google Authenticator link", flush=True)
                    time.sleep(2)
                    break
        except Exception:
            pass

        # Generate TOTP code
        totp = pyotp.TOTP(OKTA_TOTP_SECRET)
        otp_code = totp.now()
        print(f"[clean_bot:{RDP_LABEL}] OTP code generated", flush=True)

        # ── STATUS 90: Enter OTP code ──
        set_status(90, "Entering OTP code")

        # Find the OTP input field
        otp_field = None
        for _ in range(10):
            edits3 = win3.descendants(control_type="Edit")
            if edits3:
                # Try to find one labeled "Enter Code", otherwise use first edit
                for e in edits3:
                    try:
                        name = (e.element_info.name or "").lower()
                        if "code" in name or "enter" in name:
                            otp_field = e
                            break
                    except Exception:
                        pass
                if not otp_field:
                    otp_field = edits3[0]
                break
            time.sleep(1)

        if not otp_field:
            set_status(-1, "Screen 3: no OTP input field found")
            return

        _click_into(otp_field)
        otp_field.type_keys("^a", pause=0.05)
        time.sleep(0.1)
        pyperclip.copy(otp_code)
        time.sleep(0.1)
        otp_field.type_keys("^v", pause=0.05)
        time.sleep(0.3)
        print(f"[clean_bot:{RDP_LABEL}] OTP code entered", flush=True)

        # ── STATUS 100: Click Verify ──
        set_status(100, "Clicking Verify")
        time.sleep(0.5)

        verified = False
        for btn in win3.descendants(control_type="Button"):
            btn_text = btn.window_text()
            if btn_text in ("Verify", "verify"):
                try:
                    btn.click_input()
                    verified = True
                except Exception:
                    try:
                        btn.click()
                        verified = True
                    except Exception:
                        r = btn.rectangle()
                        import pyautogui
                        pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
                        verified = True
                break

        if not verified:
            print(f"[clean_bot:{RDP_LABEL}] Verify button not found, pressing Enter", flush=True)
            otp_field.type_keys("{ENTER}")

        set_status(110, "Login complete")

    except Exception as e:
        set_status(-1, f"Error: {e}")
        raise


if __name__ == "__main__":
    main()
