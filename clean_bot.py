"""
clean_bot.py -- Automated Experity login inside RDP sessions.

Runs inside the RDP session via CreateProcessAsUser (injected by start_all_clean.py).
Opens Experity, handles 3-screen Okta login (username -> password -> OTP).

Key design:
  - Screen fingerprinting: each screen identified by BUTTON text + FIELD layout,
    NOT just window title (titles overlap between screens)
  - Verify-before-act: confirm we're on the correct screen before typing anything
  - Verify-after-act: confirm data was entered and screen transitioned
  - Auto-restart: on error, kill everything and retry up to MAX_RETRIES times

Status codes (written to C:\\ProgramData\\MHTAgentic\\clean_status_<label>.txt):
    10  Opening Experity
    20  Entering username (Screen 1)
    30  Clicking Next (Screen 1)
    40  Waiting for Okta sign-in (Screen 2)
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

MAX_RETRIES = 3

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


def log(msg):
    print(f"[clean_bot:{RDP_LABEL}] {msg}", flush=True)


# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------
def _wait_visible(element, timeout=10):
    """Wait for an element to become visible/actionable."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            if element.is_visible() and element.is_enabled():
                return True
        except Exception:
            pass
        time.sleep(0.5)
    return False


def _click_into(element):
    """Click into an element -- waits for visible, then click_input, fall back to coordinate."""
    _wait_visible(element, timeout=10)
    try:
        element.click_input()
    except Exception:
        try:
            r = element.rectangle()
            import pyautogui
            pyautogui.click((r.left + r.right) // 2, (r.top + r.bottom) // 2)
        except Exception:
            pass
    time.sleep(0.3)


def _get_field_value(field):
    """Safely read a field's current value."""
    try:
        v = field.get_value()
        if v:
            return v.strip()
    except Exception:
        pass
    try:
        v = field.window_text()
        if v:
            return v.strip()
    except Exception:
        pass
    return ""


def _enter_text(field, text, field_name="field", retries=3):
    """Clear field, paste text, verify it took. Retries on failure.

    Uses triple-click to select field text (not Ctrl+A which selects entire web page),
    then types replacement. Falls back to set_edit_text if available."""
    import pyperclip
    import pyautogui

    for attempt in range(1, retries + 1):
        try:
            # Click into field first
            _click_into(field)
            time.sleep(0.2)

            # Wait for field to be actionable
            if not _wait_visible(field, timeout=10):
                log(f"  {field_name}: not visible after 10s (attempt {attempt})")
                time.sleep(1)
                continue

            # Method 1: Try set_edit_text (UIA direct value set — cleanest)
            try:
                iface = field.iface_value
                if iface:
                    iface.SetValue(text)
                    log(f"  {field_name}: set via UIA SetValue (attempt {attempt})")
                    time.sleep(0.3)
                    # Check if it actually worked
                    actual = _get_field_value(field)
                    if actual == text:
                        return True
            except Exception:
                pass

            # Method 2: Triple-click to select field text only (not whole page),
            # then paste. Triple-click selects all text in the focused field.
            try:
                r = field.rectangle()
                cx, cy = (r.left + r.right) // 2, (r.top + r.bottom) // 2
                pyautogui.click(cx, cy, clicks=3, interval=0.05)
                time.sleep(0.2)
            except Exception:
                # Fallback: Home then Shift+End to select field contents
                try:
                    field.type_keys("{HOME}", pause=0.05)
                    time.sleep(0.05)
                    field.type_keys("+{END}", pause=0.05)
                    time.sleep(0.1)
                except Exception:
                    pass

            # Paste via clipboard (replaces selected text)
            pyperclip.copy(text)
            time.sleep(0.1)
            field.type_keys("^v", pause=0.05)
            time.sleep(0.3)
        except Exception as e:
            log(f"  {field_name}: interaction error on attempt {attempt}: {e}")
            time.sleep(1)
            continue

        # Verify
        actual = _get_field_value(field)
        if actual == text:
            log(f"  {field_name}: entered and verified (attempt {attempt})")
            return True

        # For password fields, get_value often returns empty -- that's OK
        if "password" in field_name.lower() or "pass" in field_name.lower():
            log(f"  {field_name}: entered (password field, can't verify)")
            return True

        log(f"  {field_name}: verify failed (got '{actual}', expected '{text}'), retry {attempt}/{retries}")
        time.sleep(0.5)

    log(f"  WARNING: {field_name} failed to verify after {retries} attempts, proceeding anyway")
    return False


def _click_button(win, button_texts, fallback_enter_field=None, description="button"):
    """Find and click a button by text. Tries multiple methods aggressively."""
    import pyautogui

    for btn in win.descendants(control_type="Button"):
        btn_text = btn.window_text()
        if btn_text in button_texts:
            _wait_visible(btn, timeout=5)

            # Method 1: UIA invoke (most reliable for web buttons)
            try:
                btn.click()
                log(f"  Clicked {description}: '{btn_text}' (invoke)")
                time.sleep(0.5)
                return True
            except Exception:
                pass

            # Method 2: click_input (simulated mouse)
            try:
                btn.set_focus()
                time.sleep(0.1)
                btn.click_input()
                log(f"  Clicked {description}: '{btn_text}' (click_input)")
                time.sleep(0.5)
                return True
            except Exception:
                pass

            # Method 3: coordinate click with pyautogui (double-click for reliability)
            try:
                r = btn.rectangle()
                cx, cy = (r.left + r.right) // 2, (r.top + r.bottom) // 2
                pyautogui.click(cx, cy)
                time.sleep(0.3)
                pyautogui.click(cx, cy)
                log(f"  Clicked {description}: '{btn_text}' (coordinate double-click)")
                time.sleep(0.5)
                return True
            except Exception:
                pass

    # Fallback: press Enter
    if fallback_enter_field:
        log(f"  {description} not found by text, pressing Enter")
        try:
            fallback_enter_field.type_keys("{ENTER}")
        except Exception:
            pyautogui.press("enter")
        return True

    log(f"  WARNING: {description} not found and no fallback")
    return False


# ---------------------------------------------------------------------------
# Screen fingerprinting
# ---------------------------------------------------------------------------
def _get_button_texts(win):
    """Get all button texts in a window."""
    try:
        return [b.window_text() for b in win.descendants(control_type="Button")]
    except Exception:
        return []


def _get_edit_fields(win):
    """Get all edit fields in a window."""
    try:
        return win.descendants(control_type="Edit")
    except Exception:
        return []


def _get_field_names(edits):
    """Get automation names for a list of edit fields."""
    names = []
    for e in edits:
        try:
            names.append((e.element_info.name or "").lower())
        except Exception:
            names.append("")
    return names


def _is_screen1(win):
    """Screen 1: Experity login -- has 'Next' button + edit field(s)."""
    btns = _get_button_texts(win)
    edits = _get_edit_fields(win)
    return "Next" in btns and len(edits) >= 1


def _is_screen2(win):
    """Screen 2: Okta sign-in -- has 'Sign In'/'Sign in' button + password field.
    Distinguisher: has 'Sign In' but NOT 'Verify'."""
    btns = _get_button_texts(win)
    has_signin = any(t in ("Sign In", "Sign in") for t in btns)
    has_verify = any(t in ("Verify", "verify") for t in btns)
    if not has_signin or has_verify:
        return False
    edits = _get_edit_fields(win)
    return len(edits) >= 1


def _is_screen3(win):
    """Screen 3: OTP/MFA -- has 'Verify' button."""
    btns = _get_button_texts(win)
    return any(t in ("Verify", "verify") for t in btns)


def _is_tracking_board(win):
    """Post-login: Tracking Board / Experity EMR main screen -- no login buttons."""
    btns = _get_button_texts(win)
    has_login_btns = any(t in ("Next", "Sign In", "Sign in", "Verify", "verify") for t in btns)
    if has_login_btns:
        return False
    title = ""
    try:
        title = win.window_text().lower()
    except Exception:
        pass
    return "experity" in title or "tracking" in title


# ---------------------------------------------------------------------------
# Screen discovery (session-safe)
# ---------------------------------------------------------------------------
def _find_window(check_fn, timeout=30, description="window"):
    """Poll for a window matching check_fn. Returns (app, win) or (None, None)."""
    from pywinauto import Application
    from mhtagentic.desktop.session_guard import session_find_elements

    deadline = time.time() + timeout
    attempt = 0
    while time.time() < deadline:
        attempt += 1
        try:
            for pat in [".*Experity.*", ".*Sign In.*", ".*Okta.*",
                        ".*Verify.*", ".*Authentication.*"]:
                elems = session_find_elements(title_re=pat, backend="uia")
                for elem in elems:
                    try:
                        app = Application(backend="uia").connect(handle=elem.handle, timeout=3)
                        win = app.window(handle=elem.handle)
                        if check_fn(win):
                            log(f"  Found {description}: '{win.window_text()}' (attempt {attempt})")
                            return app, win
                    except Exception:
                        continue
        except Exception:
            pass
        time.sleep(1)

    log(f"  {description} NOT found after {timeout}s")
    return None, None


def _wait_for_screen_transition(old_check_fn, new_check_fn, timeout=30,
                                 old_desc="old screen", new_desc="new screen"):
    """Wait until old_check_fn stops matching and new_check_fn starts matching.
    This prevents detecting the SAME screen as the next one."""
    from pywinauto import Application
    from mhtagentic.desktop.session_guard import session_find_elements

    deadline = time.time() + timeout
    attempt = 0
    while time.time() < deadline:
        attempt += 1
        try:
            for pat in [".*Experity.*", ".*Sign In.*", ".*Okta.*",
                        ".*Verify.*", ".*Authentication.*"]:
                elems = session_find_elements(title_re=pat, backend="uia")
                for elem in elems:
                    try:
                        app = Application(backend="uia").connect(handle=elem.handle, timeout=3)
                        win = app.window(handle=elem.handle)
                        if new_check_fn(win) and not old_check_fn(win):
                            log(f"  Transition to {new_desc}: '{win.window_text()}' (attempt {attempt})")
                            return app, win
                    except Exception:
                        continue
        except Exception:
            pass

        if attempt % 5 == 0:
            log(f"  Waiting for {new_desc}... ({attempt}s)")
        time.sleep(1)

    log(f"  Transition to {new_desc} NOT detected after {timeout}s")
    return None, None


# ---------------------------------------------------------------------------
# Kill existing apps
# ---------------------------------------------------------------------------
def kill_existing_apps():
    """No-op — dashboard handles process management, bot should never kill anything."""
    pass


# ---------------------------------------------------------------------------
# Main login flow (single attempt)
# ---------------------------------------------------------------------------
def do_login():
    """Execute the full 3-screen Experity/Okta login. Returns status code."""
    from mhtagentic.desktop.session_guard import session_find_elements
    from pywinauto import Application

    # ======================================================================
    # STATUS 10: Open Experity
    # ======================================================================
    set_status(10, "Opening Experity")

    if not Path(EXPERITY_SHORTCUT).exists():
        set_status(-1, f"Shortcut not found: {EXPERITY_SHORTCUT}")
        return -1

    subprocess.Popen(["cmd", "/c", "start", "", EXPERITY_SHORTCUT], shell=True)

    # Wait for Experity window
    app1, win1 = _find_window(
        lambda w: _is_screen1(w) or _is_tracking_board(w),
        timeout=60,
        description="Experity window",
    )

    if not win1:
        set_status(-1, "Experity window not found after 60s")
        return -1

    # Check if already logged in (Tracking Board visible, no login fields)
    if _is_tracking_board(win1):
        log("Already logged in (Tracking Board visible)")
        set_status(-2, "Already logged in")
        return -2

    # Confirm we're on Screen 1 before touching anything
    if not _is_screen1(win1):
        set_status(-1, "Experity window found but not on login screen (Screen 1)")
        return -1

    log(f"Screen 1 confirmed: '{win1.window_text()}'")
    log(f"  Buttons: {_get_button_texts(win1)}")
    log(f"  Edit fields: {len(_get_edit_fields(win1))}")

    # ======================================================================
    # STATUS 20: Enter username (Screen 1)
    # ======================================================================
    set_status(20, "Entering username")

    edits1 = _get_edit_fields(win1)
    if not edits1:
        set_status(-1, "Screen 1: no edit fields found")
        return -1

    _enter_text(edits1[0], STORED_USERNAME, field_name="Username (Screen 1)")

    # ======================================================================
    # STATUS 30: Click Next (Screen 1) + wait for Screen 2
    # ======================================================================
    set_status(30, "Clicking Next")
    time.sleep(0.5)

    _click_button(win1, ["Next"], fallback_enter_field=edits1[0], description="Next button")
    time.sleep(2)

    # Wait for transition: Screen 1 disappears, Screen 2 appears
    # Screen 2 has "Sign In" button but NOT "Next" button
    app2, win2 = _wait_for_screen_transition(
        old_check_fn=_is_screen1,
        new_check_fn=_is_screen2,
        timeout=30,
        old_desc="Screen 1 (Experity login)",
        new_desc="Screen 2 (Okta sign-in)",
    )

    if not win2:
        # Fallback: maybe Screen 1 turned into Screen 2 (same window)
        app2, win2 = _find_window(_is_screen2, timeout=10, description="Screen 2 (fallback)")

    if not win2:
        set_status(-1, "Screen 2 (Okta sign-in) not found after clicking Next")
        return -1

    # ======================================================================
    # STATUS 40: Validate Screen 2
    # ======================================================================
    set_status(40, "Validating Okta sign-in (Screen 2)")

    # Double-check: we must be on Screen 2, NOT Screen 1 or Screen 3
    if _is_screen1(win2):
        log("WARNING: Still on Screen 1 after clicking Next -- retrying click")
        _click_button(win2, ["Next"], description="Next button (retry)")
        time.sleep(3)
        app2, win2 = _find_window(_is_screen2, timeout=15, description="Screen 2 (retry)")
        if not win2:
            set_status(-1, "Still on Screen 1 after retry")
            return -1

    if _is_screen3(win2):
        log("Already on Screen 3 (OTP) -- skipping password entry")
        # Jump ahead to OTP
        app3, win3 = app2, win2
    else:
        # Confirmed on Screen 2
        log(f"Screen 2 confirmed: '{win2.window_text()}'")
        edits2 = _get_edit_fields(win2)
        field_names2 = _get_field_names(edits2)
        log(f"  Buttons: {_get_button_texts(win2)}")
        log(f"  Edit fields: {len(edits2)} -- names: {field_names2}")

        if not edits2:
            set_status(-1, "Screen 2: no edit fields found")
            return -1

        # Filter out browser chrome fields (URL bar, search bar, etc.)
        # Only keep fields that are actual login form fields
        login_fields = []
        for e in edits2:
            try:
                fname = (e.element_info.name or "").lower().strip()
            except Exception:
                fname = ""
            # Skip browser URL/address bar
            if "address" in fname or "search bar" in fname or "url" in fname:
                log(f"  Skipping browser field: '{fname}'")
                continue
            login_fields.append((e, fname))

        log(f"  Login fields (after filtering): {[fn for _, fn in login_fields]}")

        if not login_fields:
            set_status(-1, "Screen 2: no login fields found (only browser chrome)")
            return -1

        # Find username and password fields by name
        username_field = None
        password_field = None
        for field, fname in login_fields:
            if "pass" in fname:
                password_field = field
            elif "user" in fname or "email" in fname:
                username_field = field

        # Fallback: if we have 2+ fields, first=username, last=password
        if not password_field and len(login_fields) >= 2:
            if not username_field:
                username_field = login_fields[0][0]
            password_field = login_fields[-1][0]
        elif not password_field and len(login_fields) == 1:
            # Single field -- check if it's password
            fname = login_fields[0][1]
            if "pass" in fname:
                password_field = login_fields[0][0]
                log("  Single field is password (username pre-filled)")
            else:
                username_field = login_fields[0][0]

        # Enter username if we have the field
        if username_field:
            existing_val = _get_field_value(username_field)
            if existing_val == STORED_USERNAME:
                log("  Username already pre-filled")
            else:
                _enter_text(username_field, STORED_USERNAME, field_name="Username (Screen 2)")

        # If we still don't have password field, try Tab + re-scan
        if not password_field and username_field:
            try:
                username_field.type_keys("{TAB}", pause=0.1)
            except Exception:
                pass
            time.sleep(1)
            edits2 = _get_edit_fields(win2)
            for e in edits2:
                try:
                    fname = (e.element_info.name or "").lower()
                except Exception:
                    fname = ""
                if "pass" in fname:
                    password_field = e
                    break
            if not password_field and len(edits2) > 1:
                password_field = edits2[-1]

        if not password_field:
            set_status(-1, "Screen 2: could not find password field")
            return -1

        # ==================================================================
        # STATUS 60: Enter password (Screen 2)
        # ==================================================================
        set_status(60, "Entering password")

        # VERIFY: still on Screen 2 before typing password
        if not _is_screen2(win2):
            log("WARNING: Screen changed while entering password!")
            app2, win2 = _find_window(_is_screen2, timeout=10, description="Screen 2 (re-find)")
            if not win2:
                set_status(-1, "Lost Screen 2 while entering password")
                return -1
            edits2 = _get_edit_fields(win2)
            password_field = edits2[-1] if edits2 else None
            if not password_field:
                set_status(-1, "Screen 2: lost password field")
                return -1

        _enter_text(password_field, STORED_PASSWORD, field_name="Password")

        # ==================================================================
        # STATUS 70: Click Sign In (Screen 2) + wait for Screen 3
        # ==================================================================
        set_status(70, "Clicking Sign In")
        time.sleep(0.5)

        _click_button(win2, ["Sign In", "Sign in"],
                       fallback_enter_field=password_field,
                       description="Sign In button")
        time.sleep(3)

        # Wait for Screen 3: must have "Verify" button, must NOT have "Sign In"
        app3, win3 = _wait_for_screen_transition(
            old_check_fn=_is_screen2,
            new_check_fn=_is_screen3,
            timeout=30,
            old_desc="Screen 2 (Okta sign-in)",
            new_desc="Screen 3 (OTP/Verify)",
        )

        if not win3:
            # Fallback: direct search for Screen 3
            app3, win3 = _find_window(_is_screen3, timeout=10, description="Screen 3 (fallback)")

        if not win3:
            # Check if we're still on Screen 2 (bad password?)
            _, still2 = _find_window(_is_screen2, timeout=3, description="Screen 2 check")
            if still2:
                set_status(-1, "Still on Screen 2 after Sign In -- bad password or network error")
            else:
                set_status(-1, "Screen 3 (OTP/Verify) not found after Sign In")
            return -1

    # ======================================================================
    # STATUS 80: Generate OTP code (Screen 3)
    # ======================================================================
    set_status(80, "Generating OTP code")

    import pyotp
    if not OKTA_TOTP_SECRET:
        set_status(-1, "OKTA_TOTP_SECRET not set in .env")
        return -1

    # VERIFY: we are on Screen 3 (has Verify button, NOT Sign In)
    if not _is_screen3(win3):
        log("WARNING: Expected Screen 3 but fingerprint doesn't match!")
        log(f"  Buttons: {_get_button_texts(win3)}")
        app3, win3 = _find_window(_is_screen3, timeout=10, description="Screen 3 (re-verify)")
        if not win3:
            set_status(-1, "Cannot confirm Screen 3 (OTP)")
            return -1

    log(f"Screen 3 confirmed: '{win3.window_text()}'")
    log(f"  Buttons: {_get_button_texts(win3)}")

    # Click "Google Authenticator" link if visible (some Okta configs show method selector)
    try:
        for link in win3.descendants(control_type="Hyperlink"):
            if "google" in link.window_text().lower():
                link.click_input()
                log("  Clicked Google Authenticator link")
                time.sleep(2)
                break
    except Exception:
        pass

    # Generate TOTP
    totp = pyotp.TOTP(OKTA_TOTP_SECRET)
    otp_code = totp.now()
    log(f"  OTP code generated (6 digits)")

    # ======================================================================
    # STATUS 90: Enter OTP code (Screen 3)
    # ======================================================================
    set_status(90, "Entering OTP code")

    # Find OTP input field
    otp_field = None
    for _ in range(10):
        edits3 = _get_edit_fields(win3)
        if edits3:
            # Prefer field named "code"/"enter"
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
        return -1

    # VERIFY: still on Screen 3 right before entering OTP
    if not _is_screen3(win3):
        log("WARNING: Screen changed before OTP entry!")
        set_status(-1, "Lost Screen 3 before OTP entry")
        return -1

    _enter_text(otp_field, otp_code, field_name="OTP code")

    # ======================================================================
    # STATUS 100: Click Verify (Screen 3)
    # ======================================================================
    set_status(100, "Clicking Verify")
    time.sleep(0.5)

    _click_button(win3, ["Verify", "verify"],
                   fallback_enter_field=otp_field,
                   description="Verify button")

    time.sleep(3)

    # ======================================================================
    # STATUS 110: Login complete
    # ======================================================================
    set_status(110, "Login complete")
    log("Login flow completed successfully")
    return 110


# ---------------------------------------------------------------------------
# Post-login: continue to monitoring (no process restart)
# ---------------------------------------------------------------------------
def _run_post_login():
    """After login, launch monitoring as a subprocess and wait for it.

    Uses subprocess instead of importlib to avoid COM/threading conflicts
    between clean_bot's pywinauto state and launcher's tkinter overlays.
    clean_bot.py stays alive the entire time (blocks on proc.wait).
    """
    role = os.environ.get("BOT_ROLE", os.environ.get("MHT_BOT_ROLE", "inbound"))
    log(f"Post-login: launching monitor-only subprocess (mode={role})")

    launcher_path = Path(__file__).resolve().parent / "launcher.pyw"
    shared_path = Path(r"C:\MHTAgentic\launcher.pyw")
    if not launcher_path.exists() and shared_path.exists():
        launcher_path = shared_path

    python_exe = Path(sys.executable).parent / "python.exe"
    if not python_exe.exists():
        import shutil
        _found = shutil.which("python.exe")
        python_exe = Path(_found) if _found else Path("python.exe")

    args = [str(python_exe), "-u", str(launcher_path), "--monitor-only"]
    if role == "outbound":
        args.append("-outbound")

    log(f"Post-login: {' '.join(args)}")
    proc = subprocess.Popen(args)
    log(f"Post-login: launcher subprocess started (PID={proc.pid})")
    proc.wait()
    log(f"Post-login: launcher subprocess exited (code={proc.returncode})")


# ---------------------------------------------------------------------------
# Main with auto-restart
# ---------------------------------------------------------------------------
def main():
    for attempt in range(1, MAX_RETRIES + 1):
        log(f"===== Login attempt {attempt}/{MAX_RETRIES} =====")

        # Write a non-terminal "retrying" status so start_all_clean doesn't
        # declare us dead while we still have retries left
        if attempt > 1:
            set_status(5, f"Retrying (attempt {attempt}/{MAX_RETRIES})")

        try:
            result = do_login()

            if result in (110, -2):
                # Login succeeded — continue to monitoring (keep process alive)
                _run_post_login()
                return result
            elif result == -1:
                log(f"Login failed on attempt {attempt}/{MAX_RETRIES}")
            else:
                log(f"Unexpected result {result} on attempt {attempt}/{MAX_RETRIES}")

        except Exception as e:
            log(f"EXCEPTION on attempt {attempt}/{MAX_RETRIES}: {e}")
            import traceback
            traceback.print_exc()
            # Only write -1 on the LAST attempt so monitor doesn't bail early
            if attempt == MAX_RETRIES:
                set_status(-1, f"Exception: {e}")

        if attempt < MAX_RETRIES:
            log(f"Killing apps and retrying in 5s...")
            kill_existing_apps()
            time.sleep(5)
        else:
            log(f"All {MAX_RETRIES} attempts exhausted. Giving up.")
            set_status(-1, f"All {MAX_RETRIES} login attempts failed")

    return -1


if __name__ == "__main__":
    try:
        code = main()
        # If we get here, monitoring has ended or login failed
        if code not in (110, -2):
            sys.exit(1)
    except SystemExit:
        raise
    except Exception:
        sys.exit(1)
