"""
MHT Agentic - Experity EMR Patient Data Extraction

Automates patient data extraction from Experity EMR for MHT SmarTest
behavioral health assessments (PHQ-9, GAD-7).

Usage:
    pythonw launcher.pyw          # Run via desktop shortcut (silent)
    python launcher.pyw           # Run with console output

Features:
    - Monitors Waiting Room for qualified patients (age 12+)
    - Extracts demographics from patient charts
    - Generates MHT API JSON files for assessment triggers
    - Tracks roomed/discharged patients
    - Visual overlays showing extraction status
"""

import sys
import time
import os
import tkinter as tk
import threading
import logging
from pathlib import Path
from datetime import datetime

logger = logging.getLogger("mhtagentic.launcher")

# Add project root to path
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(SCRIPT_DIR))

# Load .env file if present
_env_path = SCRIPT_DIR / "config" / ".env"
if _env_path.exists():
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

# Ensure pywin32 DLLs are findable (Python 3.8+ restricts DLL search)
_pywin32_dll_dirs = [
    os.path.join(os.environ.get("PYTHONPATH", "").split(";")[0], "pywin32_system32"),
    r"C:\ProgramData\MHTAgentic\site-packages\pywin32_system32",
    r"C:\ProgramData\MHTAgentic\site-packages\win32",
    str(SCRIPT_DIR),
]
for _dll_dir in _pywin32_dll_dirs:
    if os.path.isdir(_dll_dir):
        try:
            os.add_dll_directory(_dll_dir)
        except (OSError, AttributeError):
            pass

from mhtagentic.desktop.control_overlay import (
    ControlOverlay,
    reset_control_overlay,
    register_overlay_window,
    unregister_overlay_window,
    DemoStatusOverlay,
    DemoExtractedDataOverlay
)
from mhtagentic.desktop.automation import DesktopAutomation
from mhtagentic.desktop.analytics import (
    get_analytics,
    reset_analytics,
    ProcessingStage,
    ProcessingResult
)
from mhtagentic.db import MHTDatabase, EventStatus

# Global for debug
_automation = None
_control = None
_recorder = None
_debug_dir = SCRIPT_DIR / "output" / "debug"
_debug_dir.mkdir(parents=True, exist_ok=True)


class TargetOverlay:
    """Shows a visual overlay box where automation will click/type."""

    def __init__(self):
        self.root = None
        self.window = None
        self.label_window = None
        self._thread = None

    def show(self, x, y, width=200, height=40, label="Target", color="#FF6B6B", duration=2.0):
        """Show overlay at position for duration seconds."""
        self._thread = threading.Thread(
            target=self._show_overlay,
            args=(x, y, width, height, label, color, duration),
            daemon=True
        )
        self._thread.start()
        time.sleep(0.3)

    def _show_overlay(self, x, y, width, height, label, color, duration):
        try:
            self.root = tk.Tk()
            self.root.withdraw()

            self.window = tk.Toplevel(self.root)
            self.window.overrideredirect(True)
            self.window.attributes("-topmost", True)
            self.window.attributes("-alpha", 0.7)

            box_x = x - width // 2
            box_y = y - height // 2
            self.window.geometry(f"{width}x{height}+{box_x}+{box_y}")

            canvas = tk.Canvas(self.window, width=width, height=height, highlightthickness=0, bg=color)
            canvas.pack(fill=tk.BOTH, expand=True)

            border = 4
            canvas.create_rectangle(border, border, width - border, height - border, fill="black", outline="")

            try:
                self.window.attributes("-transparentcolor", "black")
            except:
                pass

            self.label_window = tk.Toplevel(self.root)
            self.label_window.overrideredirect(True)
            self.label_window.attributes("-topmost", True)

            lbl = tk.Label(self.label_window, text=f">>> {label}", font=("Segoe UI", 10, "bold"), bg=color, fg="white", padx=8, pady=4)
            lbl.pack()

            self.label_window.update_idletasks()
            lbl_width = lbl.winfo_reqwidth()
            self.label_window.geometry(f"+{x - lbl_width // 2}+{box_y - 30}")

            self.root.after(int(duration * 1000), self._close)
            self.root.mainloop()
        except:
            pass

    def _close(self):
        try:
            if self.label_window:
                self.label_window.destroy()
            if self.window:
                self.window.destroy()
            if self.root:
                self.root.quit()
                self.root.destroy()
        except:
            pass


def show_target(x, y, label, color="#FF6B6B", width=250, height=45, duration=1.5):
    """Helper to show a target overlay."""
    overlay = TargetOverlay()
    overlay.show(x, y, width, height, label, color, duration)
    return overlay


def _start_monitoring(control, outbound_worker=None, demo_mode=False):
    """Start monitoring the Waiting Room with periodic refresh and patient tracking."""
    import pyautogui
    import threading
    import tkinter as tk
    from pywinauto import Application, findwindows
    import re

    REFRESH_INTERVAL = 15  # 15 seconds
    tracked_patients = {}  # Store patient data
    previous_roomed_patients = {}  # Track roomed patients for discharge detection

    # Database for event tracking — use shared ProgramData path so dashboard can read it
    from mhtagentic import OUTPUT_DIR as _OUTPUT_DIR
    db = MHTDatabase(_OUTPUT_DIR / "mht_data.db")
    patient_event_ids = {}  # {patient_name_upper: event_id}

    # === ELEMENT DETECTION HELPERS (replace fixed sleeps) ===
    def wait_for_window(title_re, timeout=5):
        """Wait for a window to appear, return it immediately when found. No fixed sleep."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                elements = _sfind(title_re=title_re, backend='uia')
                if elements:
                    app = Application(backend='uia').connect(handle=elements[0].handle, timeout=1)
                    return app.window(handle=elements[0].handle)
            except:
                pass
            time.sleep(0.05)  # 50ms polling interval
        return None

    def wait_for_window_close(title_re, timeout=3):
        """Wait until a window disappears. No fixed sleep."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                elements = _sfind(title_re=title_re, backend='uia')
                if not elements:
                    return True
            except:
                return True
            time.sleep(0.05)
        return False

    def wait_for_element(parent_win, timeout=3, **kwargs):
        """Wait for a child element to appear in a window. Returns element or None."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                elem = parent_win.child_window(**kwargs)
                if elem.exists():
                    return elem
            except:
                pass
            time.sleep(0.05)
        return None

    # === RESOLUTION-INDEPENDENT HELPERS ===

    def extract_demographics_fields(demo_win):
        """
        Extract demographics fields by matching Text labels to nearby Edit fields.
        Resolution-independent: uses label text and relative position instead of absolute coords.
        """
        fields = {}
        try:
            edits = demo_win.descendants(control_type='Edit')
            texts = demo_win.descendants(control_type='Text')

            # Build list of labels we care about
            # Experity labels include colons (e.g. "First Name:")
            label_map = {
                'First Name:': 'first_name',
                'Last Name:': 'last_name',
                'Birthday:': 'dob',
                'Cell Phone#:': 'cell_phone',
                'Home Phone#:': 'home_phone',
                'Email:': 'email',
                'Address 1:': 'address1',
                'Zip:': 'zip',
            }

            # Find label positions
            label_positions = []
            for t in texts[:50]:
                try:
                    txt = t.window_text().strip()
                    if txt in label_map:
                        rect = t.rectangle()
                        label_positions.append({
                            'label': txt,
                            'key': label_map[txt],
                            'x': rect.left,
                            'y': (rect.top + rect.bottom) // 2,
                        })
                except:
                    pass

            # Debug: dump ALL text elements found in the demographics window
            all_texts_found = []
            for t in texts[:50]:
                try:
                    all_texts_found.append(t.window_text().strip())
                except:
                    pass
            print(f"[extract] ALL texts in demo window: {all_texts_found}", flush=True)
            print(f"[extract] labels found: {[lp['label'] for lp in label_positions]}", flush=True)
            if not label_positions:
                print("[extract] NO labels found in demographics window!", flush=True)
                return fields

            # Get window midpoint to distinguish left vs right side
            win_rect = demo_win.rectangle()
            mid_x = (win_rect.left + win_rect.right) // 2

            # Build edit field positions
            edit_entries = []
            for edit in edits[:30]:
                try:
                    txt = edit.window_text()
                    if not txt:
                        continue
                    rect = edit.rectangle()
                    edit_entries.append({
                        'text': txt,
                        'x': rect.left,
                        'y': (rect.top + rect.bottom) // 2,
                        'side': 'left' if rect.left < mid_x else 'right',
                    })
                except:
                    pass

            # Match each label to the nearest Edit at approximately the same Y
            for lbl in label_positions:
                lbl_side = 'left' if lbl['x'] < mid_x else 'right'
                best_edit = None
                best_dist = 999999
                for ed in edit_entries:
                    # Must be on the same side of the window
                    if ed['side'] != lbl_side:
                        continue
                    y_dist = abs(ed['y'] - lbl['y'])
                    if y_dist < 40 and y_dist < best_dist:
                        best_dist = y_dist
                        best_edit = ed
                if best_edit:
                    fields[lbl['key']] = best_edit['text']
        except Exception as e:
            control.add_log(f"extract_demographics_fields error: {str(e)[:40]}")

        return fields

    def close_demographics_window(demo_win):
        """
        Close a Demographics popup window using pywinauto element detection.
        Falls back to clicking the bottom-center of the window rect.
        """
        try:
            # Strategy 1: Find Close/X button among descendants
            buttons = demo_win.descendants(control_type='Button')
            for btn in buttons:
                try:
                    btn_text = btn.window_text()
                    if btn_text in ['Close', 'X', 'close', 'Cancel']:
                        btn.click_input()
                        return
                except:
                    continue

            # Strategy 2: Look for a button near the bottom of the window
            win_rect = demo_win.rectangle()
            bottom_buttons = []
            for btn in buttons:
                try:
                    rect = btn.rectangle()
                    # Button in the bottom third of the window
                    if rect.top > win_rect.top + (win_rect.bottom - win_rect.top) * 0.7:
                        bottom_buttons.append((rect, btn))
                except:
                    continue

            if bottom_buttons:
                # Click the most centered bottom button
                win_cx = (win_rect.left + win_rect.right) // 2
                bottom_buttons.sort(key=lambda b: abs(((b[0].left + b[0].right) // 2) - win_cx))
                rect = bottom_buttons[0][0]
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                return

            # Strategy 3: Fall back to bottom-center of window
            cx = (win_rect.left + win_rect.right) // 2
            cy = win_rect.bottom - 30
            pyautogui.click(cx, cy)
        except Exception as e:
            control.add_log(f"close_demographics_window error: {str(e)[:40]}")

    def close_chart_window(chart_win):
        """
        Close a chart window:
        1. Find and click the 'Close' button (sidebar close) via pywinauto
        2. Find the 'Close Chart' confirmation button and click it
        """
        from pywinauto import findwindows, Application as PwApp
        try:
            # Step 1: Click the sidebar "Close" button
            try:
                close_btn = chart_win.child_window(title='Close', control_type='Button')
                if close_btn.exists():
                    print(f"[close_chart] clicking sidebar 'Close' button", flush=True)
                    close_btn.click_input()
                    time.sleep(0.5)
            except Exception as e:
                print(f"[close_chart] sidebar close err: {e}", flush=True)
                # Fallback: top-right area click
                win_rect = chart_win.rectangle()
                pyautogui.click(win_rect.right - 50, win_rect.top + 95)
                time.sleep(0.5)

            # Step 2: Find and click the "Close Chart" button
            # It appears either in the same window or in a new "Close Chart - ..." window
            close_clicked = False

            # Try in the existing chart window first
            try:
                cc_btn = chart_win.child_window(title='Close Chart', control_type='Button')
                if cc_btn.exists():
                    print(f"[close_chart] found 'Close Chart' in chart_win, clicking", flush=True)
                    cc_btn.click_input()
                    close_clicked = True
            except:
                pass

            # Try finding a "Close Chart" window
            if not close_clicked:
                try:
                    cc_win = wait_for_window('.*Close Chart.*', timeout=3)
                    if cc_win:
                        print(f"[close_chart] found 'Close Chart' window", flush=True)
                        try:
                            cc_btn = cc_win.child_window(title='Close Chart', control_type='Button')
                            if cc_btn.exists():
                                print(f"[close_chart] clicking 'Close Chart' button in dialog", flush=True)
                                cc_btn.click_input()
                                close_clicked = True
                        except:
                            # Click all buttons with "close" in text
                            for btn in cc_win.descendants(control_type='Button'):
                                try:
                                    bt = btn.window_text()
                                    if 'close chart' in bt.lower():
                                        print(f"[close_chart] clicking button '{bt}'", flush=True)
                                        btn.click_input()
                                        close_clicked = True
                                        break
                                except:
                                    continue
                except Exception as e:
                    print(f"[close_chart] Close Chart window search err: {e}", flush=True)

            if not close_clicked:
                print(f"[close_chart] WARNING: could not find Close Chart button", flush=True)

        except Exception as e:
            print(f"[close_chart] error: {e}", flush=True)
            control.add_log(f"close_chart_window error: {str(e)[:40]}")

    # === BACKGROUND ERROR DETECTION THREAD ===
    # Runs continuously to catch Application Error popups without blocking main process
    error_monitor_running = True

    def _check_logout_popup():
        """Aggressively detect and dismiss the Experity 'Confirm Log Out' popup.
        Uses multiple strategies since it can be an in-app web dialog."""
        import pyautogui
        control.add_log("[LOGOUT CHECK] Scanning for logout popup...")
        # Also log to file so orchestrator can read it
        try:
            with open(r"C:\ProgramData\MHTAgentic\logout_popup_debug.log", "a") as _lf:
                import datetime as _dt
                _lf.write(f"[{_dt.datetime.now():%H:%M:%S}] Scanning for logout popup...\n")
                _lf.flush()
        except:
            pass
        try:
            # Strategy 1: pywinauto — find any window/dialog with logout text
            from pywinauto import Desktop
            desktop = Desktop(backend='uia')
            all_titles = []
            for w in desktop.windows():
                try:
                    all_titles.append(w.window_text() or "(no title)")
                except:
                    pass
            control.add_log(f"[LOGOUT CHECK] Windows found: {all_titles[:10]}")
            for w in desktop.windows():
                try:
                    title = w.window_text() or ""
                    # Check window title
                    if 'confirm' in title.lower() and 'log' in title.lower():
                        for btn in w.descendants(control_type='Button'):
                            if btn.window_text() == 'No':
                                btn.click_input()
                                control.add_log("[POPUP] Dismissed logout dialog (window title match)")
                                return True
                    # Check inside Experity windows for the dialog as child elements
                    if 'Tracking Board' in title or 'Experity' in title or 'DocuTAP' in title:
                        for btn in w.descendants(control_type='Button'):
                            try:
                                if btn.window_text() == 'No':
                                    # Verify this No button is near logout text
                                    rect = btn.rectangle()
                                    # Check nearby text elements
                                    for t in w.descendants(control_type='Text'):
                                        txt = (t.window_text() or "").lower()
                                        if 'log out' in txt or 'confirm log' in txt:
                                            btn.click_input()
                                            control.add_log("[POPUP] Dismissed in-app logout dialog")
                                            return True
                            except:
                                continue
                except:
                    continue
        except:
            pass

        try:
            # Strategy 2: pyautogui — scan screen pixels for the dialog
            # The dialog has a gray box with "Yes" and "No" buttons
            # Take screenshot and look for the confirmation dialog pattern
            import PIL.Image
            screenshot = pyautogui.screenshot()
            width, height = screenshot.size
            # Search for "No" button — it's typically a gray button ~80px wide
            # in the right portion of a centered dialog
            # Use pixel color matching: the dialog has a light gray background
            # with distinct button borders
        except:
            pass

        try:
            # Strategy 3: Send keyboard shortcut to dismiss
            # If a JS confirm dialog is focused, Tab moves to No, Enter clicks it
            # Or Alt+N selects No directly
            import pyautogui
            # Check if there's a dialog by looking for a specific window class
            import ctypes
            user32 = ctypes.windll.user32
            control.add_log("[LOGOUT CHECK] Strategy 3: Win32 dialog scan...")
            # Find dialog windows with class "#32770" (standard Windows dialog)
            hwnd = user32.FindWindowW("#32770", None)
            control.add_log(f"[LOGOUT CHECK] FindWindowW #32770 result: hwnd={hwnd}")
            if hwnd:
                title_buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, title_buf, 256)
                title = title_buf.value.lower()
                if 'confirm' in title or 'log out' in title:
                    # Send Tab to move from Yes to No, then Enter
                    user32.SetForegroundWindow(hwnd)
                    time.sleep(0.1)
                    pyautogui.press('tab')
                    time.sleep(0.1)
                    pyautogui.press('enter')
                    control.add_log("[POPUP] Dismissed logout dialog via keyboard (Win32 dialog)")
                    return True
            # Also try finding by enumerating all #32770 dialogs
            import ctypes.wintypes
            WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
            found_dialogs = []
            def enum_cb(hwnd, lparam):
                cls_buf = ctypes.create_unicode_buffer(64)
                user32.GetClassNameW(hwnd, cls_buf, 64)
                if cls_buf.value == '#32770' and user32.IsWindowVisible(hwnd):
                    found_dialogs.append(hwnd)
                return True
            user32.EnumWindows(WNDENUMPROC(enum_cb), 0)
            for hwnd in found_dialogs:
                title_buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, title_buf, 256)
                t = title_buf.value.lower()
                if 'confirm' in t or 'log' in t:
                    user32.SetForegroundWindow(hwnd)
                    time.sleep(0.1)
                    pyautogui.press('tab')
                    time.sleep(0.1)
                    pyautogui.press('enter')
                    control.add_log(f"[POPUP] Dismissed dialog '{title_buf.value}' via keyboard")
                    return True
        except:
            pass
        return False

    def background_error_monitor():
        """Background thread that continuously monitors for error, birthday, and logout popups."""
        from pywinauto import Desktop
        control.add_log("[ERROR MONITOR] === NEW CODE v2 — logout popup detection active ===")
        _popup_titles = ('Application Error', 'Error', 'Birthday', 'Experity', 'PROD')
        _popup_buttons = ('OK', 'Ok', 'Close', 'Yes', 'Continue', 'Accept')
        _logout_scan_count = 0
        while error_monitor_running and not control.is_killed:
            try:
                # Check for Experity logout popup first
                _check_logout_popup()

                desktop = Desktop(backend='uia')
                for w in desktop.windows():
                    try:
                        title = w.window_text()
                        if not title:
                            continue
                        # Match error popups and birthday modals
                        if any(kw in title for kw in _popup_titles) or 'Error' == title:
                            # Skip the main Tracking Board / chart windows
                            if 'Tracking Board' in title and 'Birthday' not in title:
                                continue
                            buttons = w.descendants(control_type='Button')
                            for btn in buttons:
                                try:
                                    bt = btn.window_text()
                                    if bt in _popup_buttons:
                                        btn.click_input()
                                        control.add_log(f"[POPUP MONITOR] Auto-dismissed '{title[:40]}' via [{bt}]")
                                        time.sleep(0.3)
                                        break
                                except:
                                    continue
                    except:
                        continue
            except:
                pass
            time.sleep(0.8)  # Check every 0.8 seconds

    # Start background error monitor thread
    error_monitor_thread = threading.Thread(target=background_error_monitor, daemon=True)
    error_monitor_thread.start()
    control.add_log("Background error monitor started")

    def dismiss_popup_dialogs():
        """
        Check for and dismiss common popup dialogs that may appear when opening a patient chart.
        OPTIMIZED: Reduced timeouts and sleeps for faster popup dismissal.
        """
        popup_dismissed = False
        popup_buttons = ['OK', 'Ok', 'Yes', 'Continue', 'Cancel', 'No', 'Dismiss', 'Got it', 'Accept']

        def check_window_for_popups(win):
            nonlocal popup_dismissed
            try:
                buttons = win.descendants(control_type='Button')
                for btn in buttons:
                    try:
                        btn_text = btn.window_text()
                        if btn_text in popup_buttons:
                            rect = btn.rectangle()
                            # Use window-relative bounds to filter valid popup buttons
                            _wr = win.rectangle()
                            _ww = _wr.right - _wr.left
                            _wh = _wr.bottom - _wr.top
                            if (rect.left > _wr.left + int(_ww * 0.15)
                                and rect.right < _wr.right - int(_ww * 0.15)
                                and rect.top > _wr.top + int(_wh * 0.1)
                                and rect.bottom < _wr.bottom - int(_wh * 0.1)):
                                btn.click_input()
                                control.add_log(f"Dismissed popup: {btn_text}")
                                popup_dismissed = True
                                time.sleep(0.2)
                                return True
                    except:
                        continue
            except:
                pass
            return False

        # Method 1: Check Tracking Board window (reduced timeout)
        try:
            tb_app = Application(backend='uia').connect(title='Tracking Board', timeout=1)
            tb_win = tb_app.window(title='Tracking Board')
            if check_window_for_popups(tb_win):
                return True
        except:
            pass

        # Method 2: Check Experity/error windows
        try:
            from pywinauto import Desktop
            desktop = Desktop(backend='uia')
            for w in desktop.windows():
                try:
                    title = w.window_text()
                    if 'Experity' in title or 'PROD' in title or ' - P' in title or 'Birthday' in title or 'Application Error' in title or 'Error' in title:
                        if check_window_for_popups(w):
                            return True
                except:
                    continue
        except:
            pass

        return popup_dismissed

    def create_mht_api_json(patient_data):
        """
        Create a JSON file in MHT API format (Modify Appointment request) for sending to MHT SmarTest.
        This triggers behavioral health assessments (PHQ-9, GAD-7) for eligible patients.

        API Endpoint: https://appmw.mhtech.com/api/modify-appointment/
        Request Type: PUT
        """
        import json
        from datetime import datetime
        from pathlib import Path

        # Add clinic location to raw data so outbound knows where to switch
        bot_location = os.environ.get("BOT_LOCATION", "ATTALLA")
        patient_data['clinic_location'] = bot_location

        # Create DB event with raw data FIRST (status=PENDING)
        event_id = db.create_event(patient_data, kind="patient_extraction")

        # Output directory for API JSON files
        api_output_dir = SCRIPT_DIR / "output" / "mht_api"
        api_output_dir.mkdir(parents=True, exist_ok=True)

        # Extract patient info from extracted data
        patient_id = patient_data.get('mrn', '')
        first_name = patient_data.get('first_name', '')
        last_name = patient_data.get('last_name', '')
        dob = patient_data.get('dob', '')  # Format: MM/DD/YYYY from Experity
        gender = patient_data.get('gender', '')
        email = patient_data.get('email', '')
        cell_phone = patient_data.get('cell_phone', '')
        race = patient_data.get('race', '')
        ethnicity = patient_data.get('ethnicity', '')
        language = patient_data.get('language', 'English')
        insurance = patient_data.get('insurance', '')

        # Convert DOB from MM/DD/YYYY to YYYY-MM-DD format for API
        dob_formatted = ''
        if dob:
            try:
                # Handle various date formats
                for fmt in ['%m/%d/%Y', '%m-%d-%Y', '%Y-%m-%d']:
                    try:
                        dob_obj = datetime.strptime(dob, fmt)
                        dob_formatted = dob_obj.strftime('%Y-%m-%d')
                        break
                    except:
                        continue
            except:
                dob_formatted = dob

        # Convert gender to API format
        gender_map = {'Male': 'M', 'Female': 'F', 'M': 'M', 'F': 'F', 'Unknown': 'Unknown'}
        gender_formatted = gender_map.get(gender, gender[:1].upper() if gender else '')

        # Format phone number (remove non-digits, ensure 10 digits)
        phone_formatted = ''.join(filter(str.isdigit, cell_phone)) if cell_phone else ''
        if len(phone_formatted) > 10:
            phone_formatted = phone_formatted[-10:]  # Take last 10 digits

        # Map ethnicity to API values
        ethnicity_map = {
            'NOT HISPANIC OR LATINO': 'Not Hispanic or Latino',
            'HISPANIC OR LATINO': 'Hispanic or Latino',
            'HISPANIC': 'Hispanic or Latino',
            'LATINO': 'Hispanic or Latino',
        }
        ethnicity_formatted = ethnicity_map.get(ethnicity.upper() if ethnicity else '', ethnicity)

        # Map race to API values
        race_map = {
            'WHITE': 'White',
            'BLACK': 'Black or African American',
            'AFRICAN AMERICAN': 'Black or African American',
            'ASIAN': 'Asian',
            'NATIVE AMERICAN': 'American Indian or Alaska Native',
            'AMERICAN INDIAN': 'American Indian or Alaska Native',
            'PACIFIC ISLANDER': 'Native Hawaiian or Other Pacific Islander',
        }
        race_formatted = race_map.get(race.upper() if race else '', race)

        # Map language
        language_map = {'English': 'English', 'Spanish': 'Spanish', 'ENGLISH': 'English', 'SPANISH': 'Spanish'}
        language_formatted = language_map.get(language, language)

        # Generate unique appointment ID based on patient and timestamp
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        appointment_id = f"{patient_id}_{timestamp}" if patient_id else f"APT_{timestamp}"

        # Current appointment date/time
        appointment_datetime = datetime.now().strftime('%Y-%m-%dT%H:%M:%S')

        # Build MHT API JSON structure (Modify Appointment format)
        bot_location = os.environ.get("BOT_LOCATION", "ATTALLA")
        mht_api_payload = {
            "clinic_id": 110,  # Southern Immediate Care clinic ID (placeholder, get from MHT)
            "clinic_location": bot_location,
            "patient": {
                "patient_id": patient_id,
                "patient_first_name": first_name,
                "patient_middle_name": "",
                "patient_last_name": last_name,
                "patient_date_of_birth": dob_formatted,
                "patient_sex": gender_formatted,
                "patient_email": email if email else "",
                "patient_ethnicity": ethnicity_formatted,
                "patient_mobile": phone_formatted,
                "patient_race": race_formatted,
                "patient_preferred_language": language_formatted,
                "patient_country_code": "+1",
                "patient_insurance": [insurance] if insurance else []
            },
            "clinician": {
                "clinician_email": "",  # To be filled from assigned provider
                "clinician_first_name": "",
                "clinician_last_name": "",
                "clinician_alternate_id": ""  # NPI
            },
            "appointment": {
                "appointment_id": appointment_id,
                "appointment_date": appointment_datetime,
                "appointment_reason": "EST PC"  # Primary Care appointment
            },
            "_metadata": {
                "generated_at": datetime.now().isoformat(),
                "sent_at": datetime.now().isoformat(),
                "source": "MHTAgentic Experity Automation",
                "clinic_location": bot_location,
                "patient_name_full": patient_data.get('name', f"{last_name}, {first_name}"),
                "extraction_status": "complete",
                "expired": False,
                "expired_at": None,
                "raw_extracted_data": patient_data
            }
        }

        # Save JSON file
        safe_name = f"{last_name}_{first_name}".replace(' ', '_').replace(',', '')[:30]
        filename = f"mht_api_{safe_name}_{timestamp}.json"
        filepath = api_output_dir / filename

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(mht_api_payload, f, indent=2, ensure_ascii=False)
            control.add_log(f"MHT API JSON saved: {filename}")

            # Update DB with converted data (status=CONVERTED)
            db.update_event_converted(event_id, mht_api_payload)

            # Track this patient's JSON file and event_id for expiration when discharged
            patient_name_full = patient_data.get('name', f"{last_name}, {first_name}")
            active_patient_jsons[patient_name_full.upper()] = str(filepath)
            patient_event_ids[patient_name_full.upper()] = event_id

            return str(filepath)
        except Exception as e:
            control.add_log(f"Error saving MHT API JSON: {str(e)[:30]}")
            db.record_error(event_id, str(e))
            return None

    def expire_patient_assessment(patient_name):
        """
        Mark a patient's assessment as expired when they are discharged (leave Roomed Patients).
        Updates the existing JSON file with expired=True and expired_at timestamp.
        Also updates the database event status to EXPIRED.
        """
        import json
        from datetime import datetime

        patient_name_upper = patient_name.upper()

        if patient_name_upper not in active_patient_jsons:
            control.add_log(f"No active JSON found for discharged patient: {patient_name[:20]}")
            return None

        json_path = active_patient_jsons[patient_name_upper]

        try:
            # Read existing JSON
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            # Update expiration fields
            data['_metadata']['expired'] = True
            data['_metadata']['expired_at'] = datetime.now().isoformat()

            # Save updated JSON
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

            # Expire DB event (status=EXPIRED)
            if patient_name_upper in patient_event_ids:
                db.expire_event(patient_event_ids[patient_name_upper])
                del patient_event_ids[patient_name_upper]

            control.add_log(f"Assessment EXPIRED for: {patient_name[:20]}")
            update_patient_log(f"  >> EXPIRED: {patient_name[:25]}")

            # Remove from active tracking
            del active_patient_jsons[patient_name_upper]

            return json_path
        except Exception as e:
            control.add_log(f"Error expiring assessment: {str(e)[:30]}")
            return None

    # Track patients with active (non-expired) JSON files
    # Key: patient name (uppercase), Value: JSON file path
    active_patient_jsons = {}

    # Connect to window - exclude Setup windows, use specific pattern
    # IMPORTANT: After OTP verification, EMR may take up to 60 seconds to fully load
    # We retry patiently instead of failing immediately
    control.set_step("Connecting to Experity...")
    from pywinauto import findwindows
    try:
        from mhtagentic.desktop.session_guard import session_find_elements as _sfind
    except ImportError:
        _sfind = findwindows.find_elements

    app = None
    win = None
    max_wait_seconds = 60
    retry_interval = 5

    for attempt in range(max_wait_seconds // retry_interval):
        remaining = max_wait_seconds - (attempt * retry_interval)
        control.set_step(f"Waiting for EMR... ({remaining}s)")
        control.add_log(f"Connection attempt {attempt + 1}, {remaining}s remaining...")

        try:
            # Try to connect to main Tracking Board (not Setup)
            elements = _sfind(title_re='.*Tracking Board.*', backend='uia')
            # Find the one that's NOT "Setup"
            target_handle = None
            for elem in elements:
                if 'Setup' not in elem.name:
                    target_handle = elem.handle
                    break
            if target_handle:
                app = Application(backend='uia').connect(handle=target_handle, timeout=3)
                win = app.window(handle=target_handle)
                control.add_log("Connected to Tracking Board!")
                break
        except Exception as e1:
            pass

        # Try Experity window as fallback (session-safe)
        try:
            _exp_elems = _sfind(title_re='.*Experity.*', backend='uia')
            if _exp_elems:
                app = Application(backend='uia').connect(handle=_exp_elems[0].handle, timeout=3)
                win = app.window(handle=_exp_elems[0].handle)
                control.add_log("Connected to Experity window!")
                break
        except Exception as e2:
            pass

        # No window found yet - wait and retry
        control.add_log(f"EMR not ready yet, waiting {retry_interval}s...")
        time.sleep(retry_interval)

    if not win:
        control.set_status("Error")
        control.set_step("Could not connect to EMR")
        control.add_log("ERROR: Failed to connect to Experity/Tracking Board after 60 seconds")
        return

    control.add_log(f"Connected: {win.window_text()}")

    # Log window rect for resolution debugging
    try:
        win_rect = win.rectangle()
        control.add_log(f"Window rect: left={win_rect.left}, top={win_rect.top}, right={win_rect.right}, bottom={win_rect.bottom}")
        control.add_log(f"Window size: {win_rect.right - win_rect.left}x{win_rect.bottom - win_rect.top}")
    except:
        pass

    # Find UI elements
    control.set_step("Finding UI elements...")

    # Refresh button
    refresh_btn = win.child_window(title='Refresh', control_type='Button')
    refresh_rect = refresh_btn.rectangle()
    REFRESH_X = (refresh_rect.left + refresh_rect.right) // 2
    REFRESH_Y = (refresh_rect.top + refresh_rect.bottom) // 2
    control.add_log(f"Refresh: ({REFRESH_X}, {REFRESH_Y})")

    # Waiting Room tab
    waiting_room_tab = win.child_window(title_re='.*Waiting Room.*', control_type='TabItem')
    wr_rect = waiting_room_tab.rectangle()
    WR_X, WR_Y = wr_rect.left, wr_rect.top
    WR_WIDTH = wr_rect.right - wr_rect.left
    WR_HEIGHT = wr_rect.bottom - wr_rect.top
    control.add_log(f"Waiting Room: ({WR_X}, {WR_Y})")

    # Waiting Room data group (contains patient rows)
    try:
        waiting_room_group = win.child_window(title_re='.*Waiting Room.*', control_type='Group')
        wr_group_rect = waiting_room_group.rectangle()
        control.add_log(f"Waiting Room Group found")
    except:
        wr_group_rect = None
        control.add_log("Waiting Room Group not found")

    control.set_status("Monitoring")
    control.set_step("Starting monitors...")

    # Demo overlay references (used in finally block)
    demo_status_overlay = None
    demo_extracted_overlay = None

    if demo_mode:
        # === DEMO MODE: Headless — no overlays, just scrape ===
        control.hide()

        demo_status_overlay = DemoStatusOverlay()
        demo_status_overlay.start()
        # No extracted data overlay — run silent
        demo_extracted_overlay = None

        # Initialize all overlay variables to None so monitoring loop references work
        patient_log_root = patient_log_text = cycle_label = None
        roomed_overlay_root = roomed_text = None
        discharged_overlay_root = discharged_text = None
        extracted_data_root = extracted_data_text = None
        wr_overlay_root = patient_overlay_root = None
        patient_overlay_toplevels = []
        extracted_patients_list = []
        roomed_patients = []
        previous_patients = {}
        cycle_history = {}
        current_view_cycle = [0]
        current_cycle_num = [0]
        analytics_labels = {}

        # Redefine callbacks to route to demo overlays
        def update_patient_log(message, new_cycle=None):
            pass  # No patient log in demo mode

        def update_extracted_data(patient_data):
            pass  # No visual overlay — data goes to DB only

        def update_roomed(message):
            pass

        def update_discharged(message):
            pass

        def show_patient_overlays(patients):
            # Set qualification data on each patient (no visual overlays)
            for p in patients:
                name = p.get('name', 'Patient')
                qualified, reason = check_qualification(p)
                p['qualified'] = qualified
                p['reason'] = reason
                patient_status = p.get('status', '').upper()
                patient_notes = p.get('notes', '').strip()
                if name in processed_patients:
                    p['extracted'] = True
                elif 'LOGGING' in patient_status:
                    p['extracted'] = False
                    p['qualified'] = False
                    p['reason'] = "Logging - waiting to complete"
                elif 'LOGPENDING' in patient_status:
                    p['extracted'] = False
                    p['qualified'] = False
                    p['reason'] = "LogPending - waiting for status update"
                elif 'CHARTING' in patient_status and not patient_notes:
                    p['extracted'] = False
                    p['qualified'] = False
                    p['reason'] = "Charting - waiting for notes"
                elif qualified:
                    p['extracted'] = False
                else:
                    p['extracted'] = False

        def clear_patient_overlays():
            pass

        def lift_info_overlays():
            pass

        def update_analytics_cache():
            pass

        def update_analytics_tab():
            pass

        # Wrap control.set_step to feed demo status bar
        # Route outbound steps (contain "Step X/21") to the Outbound side
        _orig_set_step = control.set_step
        def _demo_set_step(text):
            _orig_set_step(text)
            if 'Step ' in text and '/21' in text:
                demo_status_overlay.update_status(inbound_text="Idle", outbound_text=text)
            else:
                demo_status_overlay.update_status(inbound_text=text)
        control.set_step = _demo_set_step

        # Analytics still runs (writes to DB) but no dashboard overlay
        analytics = get_analytics()
        analytics.start_session()
        control.add_log(f"Demo mode - analytics session started: {analytics.current_session.session_id}")

    else:

        # === LEFT SIDE PATIENT LOG OVERLAY WITH TABS ===
        patient_log_root = None
        patient_log_text = None
        cycle_label = None
        cycle_history = {}  # Store data for each cycle: {cycle_num: [messages]}
        current_view_cycle = [0]  # Use list to allow modification in nested functions
        current_cycle_num = [0]  # Track the actual current cycle
        current_tab = ["patients"]  # Track current tab: "patients" or "analytics"
        analytics_labels = {}  # Store analytics label references for updates

        def create_patient_log_overlay():
            nonlocal patient_log_root, patient_log_text, cycle_label, analytics_labels
            patient_log_root = tk.Tk()
            patient_log_root.withdraw()

            log_win = tk.Toplevel(patient_log_root)
            log_win.overrideredirect(True)
            log_win.attributes('-topmost', True)
            log_win.attributes('-alpha', 0.9)

            # Dynamic sizing based on screen resolution
            screen_width = patient_log_root.winfo_screenwidth()
            screen_height = patient_log_root.winfo_screenheight()
            _s = min(screen_width / 1920, screen_height / 1080)
            log_width = int(380 * _s)
            log_height = int(350 * _s)
            y_pos = screen_height - log_height - int(100 * _s)
            log_win.geometry(f"{log_width}x{log_height}+{int(30 * _s)}+{y_pos}")

            # Dark theme colors
            bg_color = "#1a1a24"
            card_color = "#252532"
            accent_color = "#5cb85c"
            accent_analytics = "#4ECDC4"
            text_primary = "#f0f0f0"
            text_secondary = "#9a9a9a"

            log_win.configure(bg=bg_color)

            # Register for global hide/show
            register_overlay_window(log_win)

            # Title bar with navigation (draggable)
            title_frame = tk.Frame(log_win, bg=card_color)
            title_frame.pack(fill=tk.X)

            # Make draggable
            def start_drag(event):
                log_win._drag_x = event.x
                log_win._drag_y = event.y
            def do_drag(event):
                x = log_win.winfo_x() + (event.x - log_win._drag_x)
                y = log_win.winfo_y() + (event.y - log_win._drag_y)
                log_win.geometry(f"+{x}+{y}")
            title_frame.bind("<Button-1>", start_drag)
            title_frame.bind("<B1-Motion>", do_drag)

            # Title
            title_label = tk.Label(title_frame, text="MHT Monitor", font=("Segoe UI", 11, "bold"),
                                   bg=card_color, fg=text_primary)
            title_label.pack(side=tk.LEFT, padx=10, pady=8)
            title_label.bind("<Button-1>", start_drag)
            title_label.bind("<B1-Motion>", do_drag)

            # Live indicator
            live_label = tk.Label(title_frame, text="● LIVE", font=("Segoe UI", 9, "bold"),
                                 bg=card_color, fg=accent_color)
            live_label.pack(side=tk.RIGHT, padx=10, pady=8)

            # === TAB BAR ===
            tab_frame = tk.Frame(log_win, bg=bg_color)
            tab_frame.pack(fill=tk.X)

            # Tab buttons container
            tabs_container = tk.Frame(tab_frame, bg=bg_color)
            tabs_container.pack(fill=tk.X, padx=10, pady=(10, 0))

            # Patients tab button
            patients_tab_btn = tk.Label(tabs_container, text="  Patients  ", font=("Segoe UI", 10, "bold"),
                                        bg=accent_color, fg="#1a1a24", cursor="hand2", padx=15, pady=6)
            patients_tab_btn.pack(side=tk.LEFT, padx=(0, 5))

            # Analytics tab button
            analytics_tab_btn = tk.Label(tabs_container, text="  Analytics  ", font=("Segoe UI", 10),
                                         bg=card_color, fg=text_secondary, cursor="hand2", padx=15, pady=6)
            analytics_tab_btn.pack(side=tk.LEFT)

            # === CONTENT FRAMES ===
            # Patients content frame
            patients_frame = tk.Frame(log_win, bg=bg_color)
            patients_frame.pack(fill=tk.BOTH, expand=True)

            # Analytics content frame (hidden initially)
            analytics_frame = tk.Frame(log_win, bg=bg_color)

            # === PATIENTS TAB CONTENT ===
            # Navigation for cycles
            nav_frame = tk.Frame(patients_frame, bg=bg_color)
            nav_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

            back_btn = tk.Label(nav_frame, text="◀", font=("Segoe UI", 12, "bold"),
                               bg=bg_color, fg="#888888", cursor="hand2")
            back_btn.pack(side=tk.LEFT)
            back_btn.bind("<Button-1>", lambda e: nav_cycle(-1))

            cycle_label = tk.Label(nav_frame, text="Cycle 0", font=("Segoe UI", 10),
                                  bg=bg_color, fg=accent_color)
            cycle_label.pack(side=tk.LEFT, padx=10)

            fwd_btn = tk.Label(nav_frame, text="▶", font=("Segoe UI", 12, "bold"),
                              bg=bg_color, fg="#888888", cursor="hand2")
            fwd_btn.pack(side=tk.LEFT)
            fwd_btn.bind("<Button-1>", lambda e: nav_cycle(1))

            # Patient log text area
            text_container = tk.Frame(patients_frame, bg=bg_color)
            text_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

            patient_log_text = tk.Text(text_container, bg=card_color, fg=text_primary,
                                       font=("Consolas", 9), wrap=tk.WORD,
                                       highlightthickness=0, borderwidth=0)
            patient_log_text.pack(fill=tk.BOTH, expand=True)
            patient_log_text.insert(tk.END, "Waiting for first cycle...\n")
            patient_log_text.config(state=tk.DISABLED)

            # === ANALYTICS TAB CONTENT (Compact) ===
            def create_stat_row(parent, label, key, value_color=accent_analytics):
                row = tk.Frame(parent, bg=card_color)
                row.pack(fill=tk.X, padx=8, pady=1)
                tk.Label(row, text=label, font=("Segoe UI", 8), bg=card_color, fg=text_secondary,
                        anchor="w", width=16).pack(side=tk.LEFT)
                val_label = tk.Label(row, text="0", font=("Segoe UI", 8, "bold"), bg=card_color,
                                    fg=value_color, anchor="e")
                val_label.pack(side=tk.RIGHT)
                analytics_labels[key] = val_label

            def create_section_header(parent, title, color=accent_analytics):
                header = tk.Frame(parent, bg=bg_color)
                header.pack(fill=tk.X, padx=8, pady=(8, 2))
                tk.Frame(header, bg=color, width=3, height=12).pack(side=tk.LEFT, padx=(0, 5))
                tk.Label(header, text=title, font=("Segoe UI", 9, "bold"), bg=bg_color,
                        fg=text_primary).pack(side=tk.LEFT)

            # Session Stats Section
            create_section_header(analytics_frame, "Session Stats")
            session_card = tk.Frame(analytics_frame, bg=card_color)
            session_card.pack(fill=tk.X, padx=8, pady=2)

            create_stat_row(session_card, "Duration:", "duration")
            create_stat_row(session_card, "Patients:", "total_patients", "#5cb85c")
            create_stat_row(session_card, "Success Rate:", "success_rate", "#5cb85c")
            create_stat_row(session_card, "Avg Time:", "avg_time", "#f0ad4e")
            create_stat_row(session_card, "Per Hour:", "patients_per_hour", "#9c27b0")

            # Processing Breakdown Section
            create_section_header(analytics_frame, "Breakdown", "#FF9800")
            breakdown_card = tk.Frame(analytics_frame, bg=card_color)
            breakdown_card.pack(fill=tk.X, padx=8, pady=2)

            create_stat_row(breakdown_card, "Successful:", "successful", "#5cb85c")
            create_stat_row(breakdown_card, "Partial:", "partial", "#f0ad4e")
            create_stat_row(breakdown_card, "Failed:", "failed", "#d9534f")
            create_stat_row(breakdown_card, "WR Scans:", "scans", accent_analytics)

            # Value Assessment Section
            create_section_header(analytics_frame, "Value", "#9c27b0")
            value_card = tk.Frame(analytics_frame, bg=card_color)
            value_card.pack(fill=tk.X, padx=8, pady=2)

            create_stat_row(value_card, "Efficiency:", "efficiency", "#5cb85c")
            create_stat_row(value_card, "Time Saved:", "time_saved", accent_analytics)
            create_stat_row(value_card, "Weekly Saved:", "weekly_hours", "#9c27b0")

            # Recommendation (compact)
            rec_frame = tk.Frame(analytics_frame, bg=bg_color)
            rec_frame.pack(fill=tk.X, padx=8, pady=(5, 3))
            analytics_labels['recommendation'] = tk.Label(rec_frame, text="Gathering data...",
                                                          font=("Segoe UI", 8), bg=bg_color,
                                                          fg=text_secondary, wraplength=log_width - 40, justify="left")
            analytics_labels['recommendation'].pack(fill=tk.X)

            # === TAB SWITCHING LOGIC ===
            def switch_to_patients():
                current_tab[0] = "patients"
                analytics_frame.pack_forget()
                patients_frame.pack(fill=tk.BOTH, expand=True)
                patients_tab_btn.config(bg=accent_color, fg="#1a1a24", font=("Segoe UI", 10, "bold"))
                analytics_tab_btn.config(bg=card_color, fg=text_secondary, font=("Segoe UI", 10))

            def switch_to_analytics():
                current_tab[0] = "analytics"
                patients_frame.pack_forget()
                analytics_frame.pack(fill=tk.BOTH, expand=True)
                analytics_tab_btn.config(bg=accent_analytics, fg="#1a1a24", font=("Segoe UI", 10, "bold"))
                patients_tab_btn.config(bg=card_color, fg=text_secondary, font=("Segoe UI", 10))
                # Trigger an analytics update
                update_analytics_tab()

            patients_tab_btn.bind("<Button-1>", lambda e: switch_to_patients())
            analytics_tab_btn.bind("<Button-1>", lambda e: switch_to_analytics())

            # Hover effects for tabs
            def tab_hover_enter(btn, is_active_check):
                if current_tab[0] != is_active_check:
                    btn.config(bg="#3a3a4e")
            def tab_hover_leave(btn, active_color, is_active_check):
                if current_tab[0] != is_active_check:
                    btn.config(bg=card_color)
                else:
                    btn.config(bg=active_color)

            patients_tab_btn.bind("<Enter>", lambda e: tab_hover_enter(patients_tab_btn, "patients"))
            patients_tab_btn.bind("<Leave>", lambda e: tab_hover_leave(patients_tab_btn, accent_color, "patients"))
            analytics_tab_btn.bind("<Enter>", lambda e: tab_hover_enter(analytics_tab_btn, "analytics"))
            analytics_tab_btn.bind("<Leave>", lambda e: tab_hover_leave(analytics_tab_btn, accent_analytics, "analytics"))

            patient_log_root.mainloop()

        def nav_cycle(direction):
            """Navigate between cycles."""
            if not cycle_history:
                return
            new_cycle = current_view_cycle[0] + direction
            if new_cycle >= 1 and new_cycle <= current_cycle_num[0]:
                current_view_cycle[0] = new_cycle
                display_cycle(new_cycle)

        def display_cycle(cycle_num):
            """Display data for a specific cycle."""
            if patient_log_root and patient_log_text:
                patient_log_root.after(0, lambda: _do_display_cycle(cycle_num))

        def _do_display_cycle(cycle_num):
            if patient_log_text and cycle_label:
                patient_log_text.config(state=tk.NORMAL)
                patient_log_text.delete(1.0, tk.END)
                if cycle_num in cycle_history:
                    for msg in cycle_history[cycle_num]:
                        patient_log_text.insert(tk.END, msg + "\n")
                patient_log_text.config(state=tk.DISABLED)
                # Update cycle label
                is_live = cycle_num == current_cycle_num[0]
                cycle_label.config(text=f"Cycle {cycle_num}" + (" (LIVE)" if is_live else ""))

        log_thread = threading.Thread(target=create_patient_log_overlay, daemon=True)
        log_thread.start()
        time.sleep(0.2)

        def update_patient_log(message, new_cycle=None):
            """Add message to patient log overlay."""
            is_new_cycle = new_cycle is not None

            if is_new_cycle:
                current_cycle_num[0] = new_cycle
                current_view_cycle[0] = new_cycle
                cycle_history[new_cycle] = []

            if current_cycle_num[0] > 0:
                if current_cycle_num[0] not in cycle_history:
                    cycle_history[current_cycle_num[0]] = []
                cycle_history[current_cycle_num[0]].append(message)

            # Only update display if viewing current cycle
            if current_view_cycle[0] == current_cycle_num[0]:
                if patient_log_text and patient_log_root:
                    try:
                        patient_log_root.after(0, lambda: _do_log_update(message, is_new_cycle))
                    except:
                        pass

        def _do_log_update(message, clear_first=False):
            if patient_log_text and cycle_label:
                patient_log_text.config(state=tk.NORMAL)
                # Clear display if starting new cycle
                if clear_first:
                    patient_log_text.delete(1.0, tk.END)
                # If this is viewing current cycle, append
                if current_view_cycle[0] == current_cycle_num[0]:
                    patient_log_text.insert(tk.END, message + "\n")
                    patient_log_text.see(tk.END)
                cycle_label.config(text=f"Cycle {current_cycle_num[0]} (LIVE)")
                patient_log_text.config(state=tk.DISABLED)

        # Simple analytics data cache (updated in background)
        analytics_cache = {
            'duration': '0 min', 'total_patients': '0', 'success_rate': '0%',
            'avg_time': '0s', 'patients_per_hour': '0', 'successful': '0',
            'partial': '0', 'failed': '0', 'scans': '0', 'efficiency': '0x',
            'time_saved': '0s', 'weekly_hours': '0h', 'recommendation': 'Gathering data...'
        }

        def update_analytics_cache():
            """Update the analytics cache (called from main thread)."""
            try:
                stats = analytics.get_current_stats()
                if stats:
                    analytics_cache['duration'] = f"{stats.get('duration_minutes', 0):.1f} min"
                    analytics_cache['total_patients'] = str(stats.get('total_patients_processed', 0))
                    analytics_cache['success_rate'] = f"{stats.get('success_rate', 0):.0f}%"
                    analytics_cache['avg_time'] = f"{stats.get('avg_time_per_patient_seconds', 0):.1f}s"
                    analytics_cache['patients_per_hour'] = f"{stats.get('patients_per_hour', 0):.1f}"
                    analytics_cache['successful'] = str(stats.get('successful_extractions', 0))
                    analytics_cache['partial'] = str(stats.get('partial_extractions', 0))
                    analytics_cache['failed'] = str(stats.get('failed_extractions', 0))
                    analytics_cache['scans'] = str(stats.get('waiting_room_scans', 0))
            except Exception as e:
                logger.debug(f"Analytics cache update error: {e}")

        def update_analytics_tab():
            """Update the analytics tab display from cache."""
            if not patient_log_root or not analytics_labels:
                return
            try:
                def _do_update():
                    for key, val in analytics_cache.items():
                        if key in analytics_labels:
                            analytics_labels[key].config(text=val)
                patient_log_root.after(0, _do_update)
            except Exception:
                pass

        # === ROOMED PATIENTS OVERLAY (Top Right) ===
        roomed_overlay_root = None
        roomed_text = None
        roomed_patients = []

        def create_roomed_overlay():
            nonlocal roomed_overlay_root, roomed_text
            while True:
                try:
                    roomed_overlay_root = tk.Tk()
                    roomed_overlay_root.withdraw()

                    roomed_win = tk.Toplevel(roomed_overlay_root)
                    roomed_win.overrideredirect(True)
                    roomed_win.attributes('-topmost', True)
                    roomed_win.attributes('-alpha', 0.9)

                    screen_width = roomed_overlay_root.winfo_screenwidth()
                    screen_height = roomed_overlay_root.winfo_screenheight()
                    _s = min(screen_width / 1920, screen_height / 1080)
                    rw = int(320 * _s)
                    rh = int(200 * _s)
                    roomed_win.geometry(f'{rw}x{rh}+{screen_width - rw - int(30 * _s)}+{int(180 * _s)}')
                    roomed_win.configure(bg='#1a1a24')

                    # Register for global hide/show
                    register_overlay_window(roomed_win)

                    title_frame = tk.Frame(roomed_win, bg='#252532')
                    title_frame.pack(fill=tk.X)
                    tk.Label(title_frame, text='Roomed Patients', font=('Segoe UI', 11, 'bold'),
                            bg='#252532', fg='#f0f0f0').pack(side=tk.LEFT, padx=10, pady=8)

                    tk.Frame(roomed_win, bg='#FF9800', height=2).pack(fill=tk.X)

                    text_frame = tk.Frame(roomed_win, bg='#1a1a24')
                    text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

                    roomed_text = tk.Text(text_frame, bg='#252532', fg='#f0f0f0',
                                          font=('Consolas', 9), wrap=tk.WORD,
                                          highlightthickness=0, borderwidth=0)
                    roomed_text.pack(fill=tk.BOTH, expand=True)
                    roomed_text.insert(tk.END, 'Waiting for patients to be roomed...\n')
                    roomed_text.config(state=tk.DISABLED)

                    roomed_overlay_root.mainloop()
                except Exception as e:
                    print(f"Roomed overlay error: {e}")
                    time.sleep(0.5)
                    continue
                break

        roomed_thread = threading.Thread(target=create_roomed_overlay, daemon=True)
        roomed_thread.start()
        time.sleep(0.15)

        def update_roomed(message):
            if roomed_text and roomed_overlay_root:
                try:
                    roomed_overlay_root.after(0, lambda: _do_roomed_update(message))
                except:
                    pass

        def _do_roomed_update(message):
            try:
                if roomed_text:
                    roomed_text.config(state=tk.NORMAL)
                    roomed_text.insert(tk.END, message + "\n")
                    roomed_text.see(tk.END)
                    roomed_text.config(state=tk.DISABLED)
            except:
                pass

        # === DISCHARGED PATIENTS OVERLAY (Below Roomed - Top Right) ===
        discharged_overlay_root = None
        discharged_text = None

        def create_discharged_overlay():
            nonlocal discharged_overlay_root, discharged_text
            while True:
                try:
                    discharged_overlay_root = tk.Tk()
                    discharged_overlay_root.withdraw()

                    discharged_win = tk.Toplevel(discharged_overlay_root)
                    discharged_win.overrideredirect(True)
                    discharged_win.attributes('-topmost', True)
                    discharged_win.attributes('-alpha', 0.9)

                    screen_width = discharged_overlay_root.winfo_screenwidth()
                    screen_height = discharged_overlay_root.winfo_screenheight()
                    _s = min(screen_width / 1920, screen_height / 1080)
                    dw = int(320 * _s)
                    dh = int(170 * _s)
                    # Position below roomed overlay
                    roomed_bottom = int(180 * _s) + int(200 * _s) + int(8 * _s)
                    discharged_win.geometry(f'{dw}x{dh}+{screen_width - dw - int(30 * _s)}+{roomed_bottom}')
                    discharged_win.configure(bg='#1a1a24')

                    register_overlay_window(discharged_win)

                    title_frame = tk.Frame(discharged_win, bg='#252532')
                    title_frame.pack(fill=tk.X)
                    tk.Label(title_frame, text='Discharged Patients', font=('Segoe UI', 11, 'bold'),
                            bg='#252532', fg='#f0f0f0').pack(side=tk.LEFT, padx=10, pady=8)

                    tk.Frame(discharged_win, bg='#9C27B0', height=2).pack(fill=tk.X)  # Purple accent

                    text_frame = tk.Frame(discharged_win, bg='#1a1a24')
                    text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

                    discharged_text = tk.Text(text_frame, bg='#252532', fg='#f0f0f0',
                                              font=('Consolas', 9), wrap=tk.WORD,
                                              highlightthickness=0, borderwidth=0)
                    discharged_text.pack(fill=tk.BOTH, expand=True)
                    discharged_text.insert(tk.END, 'Tracking discharges...\n')
                    discharged_text.config(state=tk.DISABLED)

                    discharged_overlay_root.mainloop()
                except Exception as e:
                    print(f"Discharged overlay error: {e}")
                    time.sleep(0.5)
                    continue
                break

        discharged_thread = threading.Thread(target=create_discharged_overlay, daemon=True)
        discharged_thread.start()
        time.sleep(0.15)

        def update_discharged(message):
            if discharged_text and discharged_overlay_root:
                try:
                    discharged_overlay_root.after(0, lambda: _do_discharged_update(message))
                except:
                    pass

        def _do_discharged_update(message):
            try:
                if discharged_text:
                    discharged_text.config(state=tk.NORMAL)
                    discharged_text.insert(tk.END, message + "\n")
                    discharged_text.see(tk.END)
                    discharged_text.config(state=tk.DISABLED)
            except:
                pass

        # === EXTRACTED PATIENT DATA OVERLAY (Top Left) ===
        extracted_data_root = None
        extracted_data_text = None
        extracted_patients_list = []  # Store all extracted patient data

        def create_extracted_data_overlay():
            nonlocal extracted_data_root, extracted_data_text
            extracted_data_root = tk.Tk()
            extracted_data_root.withdraw()

            data_win = tk.Toplevel(extracted_data_root)
            data_win.overrideredirect(True)
            data_win.attributes('-topmost', True)
            data_win.attributes('-alpha', 0.9)

            # Dynamic sizing - position in TOP LEFT corner
            _sw = extracted_data_root.winfo_screenwidth()
            _sh = extracted_data_root.winfo_screenheight()
            _s = min(_sw / 1920, _sh / 1080)
            data_win.geometry(f'{int(360 * _s)}x{int(320 * _s)}+{int(30 * _s)}+{int(30 * _s)}')
            data_win.configure(bg='#1a1a24')

            title_frame = tk.Frame(data_win, bg='#252532')
            title_frame.pack(fill=tk.X)
            tk.Label(title_frame, text='Extracted Patient Data', font=('Segoe UI', 11, 'bold'),
                    bg='#252532', fg='#f0f0f0').pack(side=tk.LEFT, padx=10, pady=8)

            # Patient count
            count_label = tk.Label(title_frame, text='0 patients', font=('Segoe UI', 9),
                                  bg='#252532', fg='#5cb85c')
            count_label.pack(side=tk.RIGHT, padx=10, pady=8)

            tk.Frame(data_win, bg='#2196F3', height=2).pack(fill=tk.X)  # Blue accent

            text_frame = tk.Frame(data_win, bg='#1a1a24')
            text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

            extracted_data_text = tk.Text(text_frame, bg='#252532', fg='#f0f0f0',
                                          font=('Consolas', 9), wrap=tk.WORD,
                                          highlightthickness=0, borderwidth=0)
            extracted_data_text.pack(fill=tk.BOTH, expand=True)
            extracted_data_text.insert(tk.END, 'Waiting for qualified patients...\n')
            extracted_data_text.config(state=tk.DISABLED)

            # Store count label for updates
            data_win.count_label = count_label

            register_overlay_window(data_win)

            extracted_data_root.mainloop()

        extracted_data_thread = threading.Thread(target=create_extracted_data_overlay, daemon=True)
        extracted_data_thread.start()
        time.sleep(0.15)

        def update_extracted_data(patient_data):
            """Add a new patient's extracted data to the overlay."""
            if extracted_data_text and extracted_data_root:
                extracted_patients_list.append(patient_data)
                try:
                    extracted_data_root.after(0, lambda: _do_extracted_update(patient_data))
                except:
                    pass

        def _do_extracted_update(patient_data):
            if extracted_data_text:
                extracted_data_text.config(state=tk.NORMAL)

                # Format patient data nicely
                name = patient_data.get('name', 'Unknown')
                first = patient_data.get('first_name', '')
                last = patient_data.get('last_name', '')
                dob = patient_data.get('dob', 'N/A')
                mrn = patient_data.get('mrn', 'N/A')
                phone = patient_data.get('cell_phone', 'N/A')
                email = patient_data.get('email', 'N/A')
                gender = patient_data.get('gender', 'N/A')
                insurance = patient_data.get('insurance', 'N/A')
                race = patient_data.get('race', 'N/A')
                ethnicity = patient_data.get('ethnicity', 'N/A')
                language = patient_data.get('language', 'N/A')
                pharmacy = patient_data.get('pharmacy', 'N/A')

                # Clear on first entry
                if len(extracted_patients_list) == 1:
                    extracted_data_text.delete(1.0, tk.END)

                # Add separator between patients
                if len(extracted_patients_list) > 1:
                    extracted_data_text.insert(tk.END, "-" * 40 + "\n")

                # Display formatted data
                extracted_data_text.insert(tk.END, f"Name: {last}, {first}\n")
                extracted_data_text.insert(tk.END, f"DOB: {dob}\n")
                extracted_data_text.insert(tk.END, f"MRN: {mrn}\n")
                extracted_data_text.insert(tk.END, f"Phone: {phone}\n")
                extracted_data_text.insert(tk.END, f"Email: {email}\n")
                extracted_data_text.insert(tk.END, f"Gender: {gender}\n")
                extracted_data_text.insert(tk.END, f"Insurance: {insurance}\n")
                extracted_data_text.insert(tk.END, f"Race: {race}\n")
                extracted_data_text.insert(tk.END, f"Ethnicity: {ethnicity}\n")
                extracted_data_text.insert(tk.END, f"Language: {language}\n")
                extracted_data_text.insert(tk.END, f"Pharmacy: {pharmacy}\n")

                extracted_data_text.see(tk.END)
                extracted_data_text.config(state=tk.DISABLED)

        # Previous patient tracking for roomed detection
        previous_patients = {}

        # === WAITING ROOM HEADER OVERLAY ===
        wr_overlay_root = None

        def create_wr_overlay():
            nonlocal wr_overlay_root
            wr_overlay_root = tk.Tk()
            wr_overlay_root.withdraw()

            overlay = tk.Toplevel(wr_overlay_root)
            overlay.overrideredirect(True)
            overlay.attributes('-topmost', True)
            overlay.attributes('-alpha', 0.5)
            overlay.geometry(f"{WR_WIDTH}x{WR_HEIGHT}+{WR_X}+{WR_Y}")

            frame = tk.Frame(overlay, bg='#00FF00', highlightthickness=3, highlightbackground='#00FF00')
            frame.pack(fill=tk.BOTH, expand=True)
            inner = tk.Frame(frame, bg='black')
            inner.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)
            try:
                overlay.attributes('-transparentcolor', 'black')
            except:
                pass

            label_win = tk.Toplevel(wr_overlay_root)
            label_win.overrideredirect(True)
            label_win.attributes('-topmost', True)
            tk.Label(label_win, text="MONITORING", font=("Segoe UI", 9, "bold"),
                    bg='#00AA00', fg='white', padx=6, pady=3).pack()
            label_win.geometry(f"+{WR_X + WR_WIDTH + 5}+{WR_Y}")

            register_overlay_window(overlay)
            register_overlay_window(label_win)

            wr_overlay_root.mainloop()

        wr_thread = threading.Thread(target=create_wr_overlay, daemon=True)
        wr_thread.start()
        time.sleep(0.15)

        control.add_log("Overlays active")
        update_patient_log("=== MONITORING ACTIVE ===")

        # === ANALYTICS INITIALIZATION ===
        analytics = get_analytics()
        analytics.start_session()
        control.add_log(f"Analytics session started: {analytics.current_session.session_id}")

        # Callback to update analytics UI
        def update_analytics_display(stats):
            control.add_log(f"Analytics callback! Patients: {stats.get('total_patients_processed', 0) if stats else 0}")
            try:
                update_analytics_cache()
                update_analytics_tab()
            except Exception as e:
                control.add_log(f"Analytics callback err: {e}")

        analytics.set_stats_callback(update_analytics_display)

        # === PATIENT ROW OVERLAY FUNCTIONS ===
        # Single shared root for all patient overlays (created in main thread context)
        patient_overlay_root = None
        patient_overlay_toplevels = []  # Store just the Toplevel windows

        def init_patient_overlay_root():
            """Initialize the shared root for patient overlays."""
            nonlocal patient_overlay_root
            patient_overlay_root = tk.Tk()
            patient_overlay_root.withdraw()
            patient_overlay_root.mainloop()

        # Start the overlay root in background thread
        overlay_root_thread = threading.Thread(target=init_patient_overlay_root, daemon=True)
        overlay_root_thread.start()
        time.sleep(0.15)

        def create_patient_row_overlay(rect, name, color="#00FF00"):
            """Create overlay border around a patient row. Returns list of Toplevel windows."""
            toplevels = []

            # Use actual rect dimensions from DataItem
            row_x = rect.left
            row_y = rect.top
            row_width = rect.right - rect.left
            row_height = rect.bottom - rect.top
            border_width = 3

            def create_borders():
                nonlocal toplevels
                # Top border
                top = tk.Toplevel(patient_overlay_root)
                top.overrideredirect(True)
                top.attributes('-topmost', True)
                top.attributes('-alpha', 0.8)
                top.geometry(f"{row_width}x{border_width}+{row_x}+{row_y}")
                top.configure(bg=color)
                toplevels.append(top)
                register_overlay_window(top)

                # Bottom border
                bottom = tk.Toplevel(patient_overlay_root)
                bottom.overrideredirect(True)
                bottom.attributes('-topmost', True)
                bottom.attributes('-alpha', 0.8)
                bottom.geometry(f"{row_width}x{border_width}+{row_x}+{row_y + row_height - border_width}")
                bottom.configure(bg=color)
                toplevels.append(bottom)
                register_overlay_window(bottom)

                # Left border
                left = tk.Toplevel(patient_overlay_root)
                left.overrideredirect(True)
                left.attributes('-topmost', True)
                left.attributes('-alpha', 0.8)
                left.geometry(f"{border_width}x{row_height}+{row_x}+{row_y}")
                left.configure(bg=color)
                toplevels.append(left)
                register_overlay_window(left)

                # Right border
                right = tk.Toplevel(patient_overlay_root)
                right.overrideredirect(True)
                right.attributes('-topmost', True)
                right.attributes('-alpha', 0.8)
                right.geometry(f"{border_width}x{row_height}+{row_x + row_width - border_width}+{row_y}")
                right.configure(bg=color)
                toplevels.append(right)
                register_overlay_window(right)

            # Schedule creation in the overlay root's thread
            if patient_overlay_root:
                patient_overlay_root.after(0, create_borders)
                time.sleep(0.05)  # Give time for creation

            return toplevels

        def clear_patient_overlays():
            """Clear patient row overlays (red/blue borders)."""
            nonlocal patient_overlay_toplevels

            def do_clear():
                for tl in patient_overlay_toplevels:
                    try:
                        unregister_overlay_window(tl)
                        tl.destroy()
                    except:
                        pass

            if patient_overlay_root:
                patient_overlay_root.after(0, do_clear)
                time.sleep(0.1)  # Give time for overlays to clear

            patient_overlay_toplevels = []

        def lift_info_overlays():
            """Bring info overlays (patient log, roomed, discharged) to top of z-order."""
            def do_lift(root):
                if root:
                    try:
                        for child in root.winfo_children():
                            try:
                                child.lift()
                                child.attributes('-topmost', True)
                            except:
                                pass
                    except:
                        pass

            # Lift all info overlays above patient row overlays
            for root in [patient_log_root, roomed_overlay_root, discharged_overlay_root, extracted_data_root]:
                if root:
                    try:
                        root.after(0, lambda r=root: do_lift(r))
                    except:
                        pass

        def show_patient_overlays(patients):
            """Show overlay borders around patient rows with qualification colors.
            Colors: GREEN = extracted, BLUE = qualified (not yet extracted), RED = not qualified
            """
            nonlocal patient_overlay_toplevels

            # Create new overlays for each patient
            for p in patients:
                rect = p.get('rect')
                name = p.get('name', 'Patient')
                if rect:
                    # Check qualification
                    qualified, reason = check_qualification(p)
                    p['qualified'] = qualified
                    p['reason'] = reason

                    # Determine color based on status
                    patient_status = p.get('status', '').upper()
                    patient_notes = p.get('notes', '').strip()

                    if name in processed_patients:
                        color = "#4CAF50"  # GREEN - already extracted
                        p['extracted'] = True
                    elif 'LOGGING' in patient_status:
                        # LOGGING status means patient is still being logged in - skip
                        color = "#FF9800"  # ORANGE - still logging in, wait
                        p['extracted'] = False
                        p['qualified'] = False  # Mark as not qualified until logged
                        p['reason'] = "Logging - waiting to complete"
                    elif 'LOGPENDING' in patient_status:
                        # LOGPENDING status means patient registration is pending - skip
                        color = "#FF9800"  # ORANGE - log pending, wait for completion
                        p['extracted'] = False
                        p['qualified'] = False  # Mark as not qualified until status changes
                        p['reason'] = "LogPending - waiting for status update"
                    elif 'CHARTING' in patient_status and not patient_notes:
                        # Charting but no notes yet - highlight orange, skip for now
                        color = "#FF9800"  # ORANGE - charting with no notes, wait
                        p['extracted'] = False
                        p['qualified'] = False  # Mark as not qualified until notes added
                        p['reason'] = "Charting - waiting for notes"
                    elif qualified:
                        # LOGGED status patients with notes are eligible if they qualify
                        color = "#2196F3"  # BLUE - qualified, not yet extracted
                        p['extracted'] = False
                    else:
                        color = "#F44336"  # RED - not qualified
                        p['extracted'] = False

                    # Create overlay and track the toplevels
                    new_toplevels = create_patient_row_overlay(rect, name, color)
                    patient_overlay_toplevels.extend(new_toplevels)

            # Lift info overlays above patient row overlays
            lift_info_overlays()

    def find_patient_rows():
        """Find patient data rows in the Waiting Room with full details."""
        patients = []

        try:
            # Find the Waiting Room group
            wr_group = win.child_window(title_re='.*Waiting Room.*', control_type='Group')

            # Get DataItems - each represents a patient row
            data_items = wr_group.descendants(control_type='DataItem')

            for di in data_items:
                rect = di.rectangle()
                row_data = {'rect': rect, 'row_y': rect.top, 'data_item': di}

                # Get data from children of this DataItem
                children = di.descendants()
                for child in children:
                    try:
                        text = child.window_text()
                        if not text or 'Row' not in text:
                            continue

                        match = re.match(r'Row (\d+), Column (\d+): (.*)', text)
                        if match:
                            row_num = int(match.group(1))
                            col_num = int(match.group(2))
                            value = match.group(3).strip()

                            row_data['row_num'] = row_num

                            if value:
                                if col_num == 7:  # Patient name
                                    row_data['name'] = value
                                    # Store name element rect for clicking
                                    row_data['name_rect'] = child.rectangle()
                                elif col_num == 10:  # Status (Charting, Logged, Waiting, etc.)
                                    row_data['status'] = value
                                elif col_num == 11:  # Age
                                    row_data['age_text'] = value
                                    try:
                                        row_data['age'] = int(value.split()[0])
                                    except:
                                        row_data['age'] = 0
                                elif col_num == 14:  # Notes
                                    row_data['notes'] = value
                    except:
                        pass

                # Only add if we found a patient name
                if 'name' in row_data and row_data['name']:
                    patients.append(row_data)

            # Sort by row number
            patients.sort(key=lambda p: p.get('row_num', 0))

        except Exception as e:
            control.add_log(f"Find patients error: {str(e)[:40]}")

        return patients

    def find_roomed_patients():
        """Find patient data rows in the Roomed Patients section."""
        patients = []

        try:
            # Find the Roomed Patients group
            roomed_group = win.child_window(title_re='.*Roomed Patients.*', control_type='Group')

            # Get DataItems - each represents a row
            data_items = roomed_group.descendants(control_type='DataItem')

            for di in data_items:
                row_data = {'rect': di.rectangle()}

                # Get data from children of this DataItem
                children = di.descendants()
                for child in children:
                    try:
                        text = child.window_text()
                        if not text or 'Row' not in text:
                            continue

                        match = re.match(r'Row (\d+), Column (\d+): (.*)', text)
                        if match:
                            row_num = int(match.group(1))
                            col_num = int(match.group(2))
                            value = match.group(3).strip()

                            row_data['row_num'] = row_num

                            if value:
                                if col_num == 6:  # Patient name
                                    row_data['name'] = value
                                    # Store name element rect for clicking (like find_patient_rows)
                                    row_data['name_rect'] = child.rectangle()
                                elif col_num == 1:  # Room
                                    row_data['room'] = value
                                elif col_num == 10:  # Age
                                    row_data['age_text'] = value
                                    try:
                                        row_data['age'] = int(value.split()[0])
                                    except:
                                        row_data['age'] = 0
                                elif col_num == 0:  # Time in room
                                    row_data['time'] = value
                    except:
                        pass

                # Store chart icon rect from Image elements in the row
                try:
                    images = di.descendants(control_type='Image')
                    if images:
                        row_data['chart_icon_rect'] = images[0].rectangle()
                except:
                    pass

                # Only add if we found a patient name
                if 'name' in row_data and row_data['name']:
                    patients.append(row_data)

            # Sort by row number
            patients.sort(key=lambda p: p.get('row_num', 0))

        except Exception as e:
            control.add_log(f"Find roomed error: {str(e)[:40]}")

        return patients

    def check_qualification(patient):
        """
        Check if patient qualifies for GAD-7 & PHQ-9.
        Returns: (qualified: bool, reason: str)

        QUALIFIED: Anyone with a chart (status = LOGGED, CHARTING, etc.)
        DISQUALIFIED: No chart yet (status = LOGGING, LOGPENDING, WAITING, etc.)
        NO age, insurance, or visit type restrictions - we take everyone WITH a chart
        """
        status = patient.get('status', '').upper()

        # Statuses that indicate NO CHART yet - must wait
        if not status:
            return False, "No status - waiting"
        if 'LOGGING' in status:
            return False, "Logging - no chart yet"
        if 'LOGPENDING' in status:
            return False, "LogPending - no chart yet"
        if 'WAITING' in status and 'CHARTING' not in status:
            return False, "Waiting - no chart yet"

        # Statuses that indicate HAS CHART - qualified
        if 'LOGGED' in status or 'CHARTING' in status:
            return True, "Has chart - eligible"

        # Unknown status - assume has chart if status exists
        return True, f"Status: {status} - eligible"

    # Track patients we've already processed to avoid re-clicking
    processed_patients = set()

    def click_qualified_patient(patient):
        """
        Extract patient demographics via Demographics popup + chart.
        EVENT-DRIVEN: Detects when windows/elements appear instead of fixed sleeps.
        Uses stored data_item ref to find chart icon without re-searching.
        """
        name = patient.get('name', 'Unknown')
        name_rect = patient.get('name_rect')
        data_item = patient.get('data_item')

        if not name_rect:
            control.add_log(f"No name_rect for {name}")
            return None

        cx = (name_rect.left + name_rect.right) // 2
        cy = (name_rect.top + name_rect.bottom) // 2

        control.add_log(f"Clicking {name} at ({cx}, {cy})")
        control.set_step(f"Extracting {name[:15]}...")

        clear_patient_overlays()

        # Click patient name and WAIT for Demographics window to appear (no fixed sleep)
        print(f"[extract] clicking {name} at ({cx},{cy})", flush=True)
        pyautogui.click(cx, cy)
        patient_data = {'name': name, 'clicked': True}

        try:
            # Wait for Demographics window - proceeds instantly when it appears
            demo_win = wait_for_window('.*Demographics.*', timeout=5)
            if not demo_win:
                print(f"[extract] no demo win on 1st try, dismissing popups", flush=True)
                dismiss_popup_dialogs()
                demo_win = wait_for_window('.*Demographics.*', timeout=3)

            if not demo_win:
                print(f"[extract] Demographics window NEVER opened for {name}", flush=True)
                control.add_log(f"Demographics window didn't open for {name}")
                close_demographics_window(demo_win) if demo_win else None
                return patient_data

            print(f"[extract] Demographics window found for {name}", flush=True)
            control.set_step("Extracting...")

            # BATCH EXTRACT: Get all elements at once using label-based detection
            demo_fields = extract_demographics_fields(demo_win)
            print(f"[extract] demo_fields={demo_fields}", flush=True)
            patient_data.update(demo_fields)

            texts = demo_win.descendants(control_type='Text')
            radios = demo_win.descendants(control_type='RadioButton')

            for txt_elem in texts[:30]:
                try:
                    txt = txt_elem.window_text()
                    if txt and 'Patient Number:' in txt:
                        patient_data['mrn'] = txt.split(':')[-1].strip()
                        break
                except:
                    pass

            try:
                _demo_rect = demo_win.rectangle()
                _gender_threshold = _demo_rect.top + int((_demo_rect.bottom - _demo_rect.top) * 0.6)
                for rb in radios:
                    txt = rb.window_text()
                    if txt in ['Male', 'Female', 'Unknown']:
                        rect = rb.rectangle()
                        if rect.top < _gender_threshold:
                            try:
                                legacy = rb.legacy_properties()
                                state = legacy.get('State', 0)
                                if (state & 16) != 0:
                                    patient_data['gender'] = txt
                                    break
                            except:
                                pass
            except:
                pass

            # NOTE: Race/Ethnicity/Language are ONLY extracted from the chart Demographics tab,
            # not from this popup — the popup layout causes false matches.

            extracted = [f"{k}={v}" for k, v in patient_data.items() if k not in ['name', 'clicked']]
            control.add_log(f"Extracted: {', '.join(extracted)}")

            # Close demographics and WAIT for it to disappear (no fixed sleep)
            print(f"[extract] closing demographics for {name}", flush=True)
            close_demographics_window(demo_win)
            closed = wait_for_window_close('.*Demographics.*', timeout=3)
            print(f"[extract] demographics closed={closed}", flush=True)

        except Exception as e:
            print(f"[extract] Demographics error: {e}", flush=True)
            control.add_log(f"Demographics error: {str(e)[:30]}")
            if demo_win:
                close_demographics_window(demo_win)
            wait_for_window_close('.*Demographics.*', timeout=2)

        # === INSURANCE EXTRACTION - RE-SCAN TO FIND CHART ICON ===
        control.set_step("Opening chart...")

        try:
            # Always re-scan the Waiting Room to find the patient's row fresh
            # (stored data_item refs go stale after Demographics popup closes)
            chart_icon = None
            try:
                wr_group = win.child_window(title_re='.*Waiting Room.*', control_type='Group')
                data_items = wr_group.descendants(control_type='DataItem')
                patient_name = name.upper()
                for di in data_items:
                    try:
                        di_text = di.window_text()
                        if patient_name in di_text.upper():
                            images_in_row = di.descendants(control_type='Image')
                            if images_in_row:
                                chart_icon = images_in_row[0]
                            control.add_log(f"Re-scanned: found {name} row")
                            break
                    except:
                        continue
            except:
                control.add_log("Re-scan failed, trying stored ref")

            # Last resort: try stored data_item reference
            if not chart_icon and data_item:
                try:
                    images_in_row = data_item.descendants(control_type='Image')
                    if images_in_row:
                        chart_icon = images_in_row[0]
                except:
                    control.add_log("Stored data_item also stale")

            if chart_icon:
                chart_rect = chart_icon.rectangle()
                chart_cx = (chart_rect.left + chart_rect.right) // 2
                chart_cy = (chart_rect.top + chart_rect.bottom) // 2

                # Click chart icon and WAIT for chart window to appear (no fixed sleep)
                pyautogui.click(chart_cx, chart_cy)
                control.set_step("Loading chart...")

                name_part = name.split(",")[0].strip()
                chart_title_re = f'.*{re.escape(name_part)}.*'
                chart_win = wait_for_window(chart_title_re, timeout=8)

                if not chart_win:
                    dismiss_popup_dialogs()
                    chart_win = wait_for_window(chart_title_re, timeout=4)

                if not chart_win:
                    control.add_log("Chart didn't open - skipping insurance")
                    raise Exception("Chart failed to open")

                # Wait for Demographics tab to actually exist before clicking
                demo_tab_clicked = False
                demo_deadline = time.time() + 10  # 10 second timeout
                while time.time() < demo_deadline:
                    try:
                        tab_items = chart_win.descendants(control_type='TabItem')
                        for elem in tab_items:
                            tab_name = elem.element_info.name or ''
                            if 'demograph' in tab_name.lower():
                                # Verify the element is visible and has a valid rect
                                rect = elem.rectangle()
                                if rect.right > rect.left and rect.bottom > rect.top:
                                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                    demo_tab_clicked = True
                                    control.add_log("Demographics tab found and clicked")
                                    break
                    except Exception:
                        pass
                    if demo_tab_clicked:
                        break
                    time.sleep(0.2)
                if not demo_tab_clicked:
                    control.add_log("Demographics tab never appeared after 10s - skipping insurance")
                    raise Exception("Demographics tab not found")

                # Wait for tab content by checking for text elements to appear
                deadline = time.time() + 3
                texts = []
                while time.time() < deadline:
                    try:
                        texts = chart_win.descendants(control_type='Text')
                        if len(texts) > 10:
                            break
                    except:
                        pass
                    time.sleep(0.05)

                # Extract insurance from window title
                try:
                    title = chart_win.window_text()
                    if title and ' - ' in title:
                        parts = title.split(' - ')
                        if len(parts) >= 2:
                            insurance_part = parts[-1].strip()
                            if insurance_part and insurance_part not in ['NEW']:
                                patient_data['insurance'] = insurance_part
                                control.add_log(f"Insurance from title: {insurance_part}")
                except:
                    pass

                # Extract from chart text elements
                try:
                    if not texts:
                        texts = chart_win.descendants(control_type='Text')

                    # Primary Insurance
                    if 'insurance' not in patient_data or patient_data.get('insurance') in ['SELF PAY', 'None']:
                        found_primary_label = False
                        for t in texts[:100]:
                            try:
                                txt = t.window_text()
                                if txt:
                                    if 'Primary Insurance' in txt:
                                        found_primary_label = True
                                    elif found_primary_label and txt.strip() and txt.strip() != 'None':
                                        if any(ins in txt.upper() for ins in ['MEDICAID', 'MEDICARE', 'BCBS', 'BLUE', 'AETNA', 'CIGNA', 'UNITED', 'HUMANA', 'ANTHEM', 'KAISER']):
                                            patient_data['insurance'] = txt.strip()
                                            control.add_log(f"Insurance: {txt.strip()}")
                                        found_primary_label = False
                            except:
                                pass

                    # Race/Ethnicity/Language from chart Cultural Information section
                    # Position-based: find label, then value at same y to its right
                    # Only match elements under the "Cultural Information" header
                    chart_texts_pos = []
                    cultural_info_y = None
                    for t in chart_win.descendants(control_type='Text'):
                        try:
                            txt = t.window_text()
                            if txt and txt.strip():
                                rect = t.rectangle()
                                entry = {'txt': txt.strip(), 'x': rect.left, 'y': rect.top}
                                chart_texts_pos.append(entry)
                                if txt.strip() == 'Cultural Information':
                                    cultural_info_y = rect.top
                        except:
                            pass

                    if cultural_info_y is not None:
                        # Only look at elements below "Cultural Information" header
                        cultural_texts = [e for e in chart_texts_pos if e['y'] > cultural_info_y and e['y'] < cultural_info_y + 150]
                        for label_key, data_key in [('Race:', 'race'), ('Ethnicity:', 'ethnicity'), ('Language:', 'language')]:
                            for tp in cultural_texts:
                                if tp['txt'] == label_key:
                                    ly, lx = tp['y'], tp['x']
                                    for vp in cultural_texts:
                                        if abs(vp['y'] - ly) < 20 and vp['x'] > lx + 30 and vp['txt'] != label_key:
                                            patient_data[data_key] = vp['txt']
                                            control.add_log(f"Chart {data_key}: {vp['txt']}")
                                            break
                                    break
                    else:
                        control.add_log("Cultural Information section not found in chart")
                except Exception as ins_err:
                    control.add_log(f"Insurance extraction error: {str(ins_err)[:30]}")

                # Close sidebar then chart, WAIT for chart to disappear
                print(f"[extract] closing chart window", flush=True)
                close_chart_window(chart_win)
                chart_closed = wait_for_window_close(chart_title_re, timeout=3)
                print(f"[extract] chart closed={chart_closed}", flush=True)
                control.add_log("Chart closed")

            else:
                print(f"[extract] chart icon NOT found", flush=True)
                control.add_log("Chart icon not found")

        except Exception as chart_err:
            print(f"[extract] Chart error: {chart_err}", flush=True)
            control.add_log(f"Chart error: {str(chart_err)[:30]}")
            try:
                if chart_win:
                    close_chart_window(chart_win)
            except:
                pass

        return patient_data

    def find_next_qualified_patient(patients):
        """Find the next qualified patient that hasn't been processed yet."""
        for p in patients:
            name = p.get('name', 'Unknown')
            qualified = p.get('qualified', False)

            if qualified and name not in processed_patients:
                return p
        return None

    def process_one_qualified_patient(patient):
        """Process a single qualified patient and return extracted data."""
        name = patient.get('name', 'Unknown')
        _extract_start = time.time()

        # Start analytics
        try:
            analytics.start_patient_processing(name, '')
        except Exception as e:
            control.add_log(f"Analytics start err: {e}")

        # Click and get data
        data = click_qualified_patient(patient)

        # Only count as successfully processed if we got real demographic data
        # (not just {'name': ..., 'clicked': True} from a failed extraction)
        has_real_data = data and (data.get('dob') or data.get('first_name') or data.get('last_name'))

        if has_real_data:
            # Record extraction duration
            data['_extract_seconds'] = round(time.time() - _extract_start, 1)
            processed_patients.add(name)
            update_patient_log(f"  >> Extracted data for {name}")
            update_extracted_data(data)
            control.set_step(f"Extracted: {name[:15]}")

            # Create JSON
            json_path = create_mht_api_json(data)
            if json_path:
                update_patient_log(f"  >> JSON created")

            # Update analytics
            analytics.end_patient_processing(ProcessingResult.SUCCESS)

            # Force analytics UI update
            stats = analytics.get_current_stats()
            if stats:
                control.add_log(f"Patients: {stats.get('total_patients_processed', 0)}, Success: {stats.get('successful_extractions', 0)}")

            return data

        # Failed — do NOT add to processed_patients so bot retries next cycle
        control.add_log(f"Extraction failed for {name} - will retry next cycle")
        analytics.end_patient_processing(ProcessingResult.FAILED)
        return None

    def process_roomed_qualified_patient(roomed_patient, tb_win):
        """
        Process a qualified patient that moved to Roomed Patients section before we could extract data.

        This handles the edge case where a patient was in Waiting Room, qualified for GAD-7/PHQ-9,
        but moved to Roomed Patients before we could click on them to extract their demographics.

        Steps:
        1. Click on the patient's name in Roomed Patients to open Demographics
        2. Extract patient data
        3. Close Demographics
        4. Click on the chart icon in Roomed Patients section
        5. Extract insurance info and close chart
        """
        name = roomed_patient.get('name', 'Unknown')

        if name in processed_patients:
            control.add_log(f"Roomed patient {name} already processed, skipping")
            return None

        control.add_log(f"Processing roomed qualified patient: {name}")
        control.set_step(f"Processing roomed patient: {name[:15]}...")
        update_patient_log(f">> Processing roomed patient: {name}")

        # Clear overlays before clicking
        clear_patient_overlays()

        patient_data = {'name': name, 'from_roomed': True}
        chart_icon = None  # Not used with new rect-based approach

        try:
            # Check if we have rect data from find_roomed_patients
            row_rect = roomed_patient.get('rect')
            name_rect = roomed_patient.get('name_rect')

            if name_rect:
                # Use the name_rect directly - most accurate for clicking the name
                cx = (name_rect.left + name_rect.right) // 2
                cy = (name_rect.top + name_rect.bottom) // 2
                control.add_log(f"Using name_rect for {name}, clicking at ({cx}, {cy})")
            elif row_rect:
                # Use the row rect center as fallback
                cx = (row_rect.left + row_rect.right) // 2
                cy = (row_rect.top + row_rect.bottom) // 2
                control.add_log(f"Using row_rect for {name}, clicking at ({cx}, {cy})")
            else:
                # Fallback: Find the Roomed Patients group and search for patient
                control.add_log(f"No rect provided, searching for {name} in Roomed Patients...")
                roomed_group = tb_win.child_window(title_re='.*Roomed Patients.*', control_type='Group')
                roomed_data_items = roomed_group.descendants(control_type='DataItem')

                patient_name = name.upper()
                found_rect = None

                # Find the patient row
                for di in roomed_data_items:
                    try:
                        children = di.descendants()
                        found_name = False

                        for child in children:
                            child_text = child.window_text()
                            if child_text and patient_name in child_text.upper():
                                found_name = True
                                found_rect = child.rectangle()
                                break

                        if found_name:
                            if not found_rect:
                                found_rect = di.rectangle()
                            control.add_log(f"Found {name} in Roomed section")
                            break

                    except Exception as row_err:
                        continue

                if not found_rect:
                    control.add_log(f"Could not find {name} in Roomed Patients section")
                    return None

                cx = (found_rect.left + found_rect.right) // 2
                cy = (found_rect.top + found_rect.bottom) // 2

            control.add_log(f"Clicking roomed patient {name} at ({cx}, {cy})")
            pyautogui.click(cx, cy)
            time.sleep(0.1)
            pyautogui.click(cx, cy)  # Double click to ensure it registers

            # Wait for Demographics/patient window to appear (no fixed sleep)
            control.set_step("Loading demographics...")
            name_part = name.split(',')[0].upper()
            demo_win_roomed = wait_for_window(f'.*{re.escape(name_part)}.*', timeout=5)
            if not demo_win_roomed:
                dismiss_popup_dialogs()
                demo_win_roomed = wait_for_window(f'.*{re.escape(name_part)}.*', timeout=3)

            # Extract data from Demographics window
            try:
                demo_win = demo_win_roomed
                if not demo_win:
                    control.add_log("ERROR: Could not find Demographics window!")
                    return None

                control.add_log(f"Connected to Demographics for roomed patient")

                # Extract demographics fields using label-based detection
                demo_fields = extract_demographics_fields(demo_win)
                patient_data.update(demo_fields)
                for k, v in demo_fields.items():
                    control.add_log(f"{k}: {v}")

                # Get MRN
                texts = demo_win.descendants(control_type='Text')
                for txt_elem in texts[:30]:
                    try:
                        txt = txt_elem.window_text()
                        if txt and 'Patient Number:' in txt:
                            mrn = txt.split(':')[-1].strip()
                            patient_data['mrn'] = mrn
                            control.add_log(f"MRN: {mrn}")
                            break
                    except:
                        pass

                # Extract Sex at Birth
                try:
                    radios = demo_win.descendants(control_type='RadioButton')
                    for rb in radios:
                        txt = rb.window_text()
                        if txt in ['Male', 'Female', 'Unknown']:
                            rect = rb.rectangle()
                            # Use window-relative threshold (top 60% of window)
                            _win_rect = demo_win.rectangle()
                            _threshold = _win_rect.top + int((_win_rect.bottom - _win_rect.top) * 0.6)
                            if rect.top < _threshold:
                                try:
                                    legacy = rb.legacy_properties()
                                    state = legacy.get('State', 0)
                                    is_checked = (state & 16) != 0
                                    if is_checked:
                                        patient_data['gender'] = txt
                                        control.add_log(f"Sex at Birth: {txt}")
                                        break
                                except:
                                    pass
                except:
                    pass

                # NOTE: Race, Ethnicity, Language, and Insurance are ALL extracted from the Chart
                # (Demographics tab), not from this Demographics popup. We just get basic info here.

                # Close Demographics and wait for it to disappear
                close_demographics_window(demo_win)
                wait_for_window_close(f'.*{re.escape(name_part)}.*', timeout=2)
                control.add_log("Closed demographics")

            except Exception as demo_err:
                control.add_log(f"Demographics error: {str(demo_err)[:30]}")
                if demo_win:
                    close_demographics_window(demo_win)
                wait_for_window_close(f'.*{re.escape(name_part)}.*', timeout=2)

            # Click chart icon and WAIT for chart window (no fixed sleep)
            control.set_step("Opening chart...")
            try:
                chart_icon_rect = roomed_patient.get('chart_icon_rect')
                if chart_icon_rect:
                    chart_cx = (chart_icon_rect.left + chart_icon_rect.right) // 2
                    chart_cy = (chart_icon_rect.top + chart_icon_rect.bottom) // 2
                else:
                    # Fallback: re-scan for chart icon Image element in the row
                    chart_cx, chart_cy = None, None
                    try:
                        roomed_group = tb_win.child_window(title_re='.*Roomed Patients.*', control_type='Group')
                        data_items = roomed_group.descendants(control_type='DataItem')
                        patient_name_upper = name.upper()
                        for di in data_items:
                            try:
                                di_text = di.window_text()
                                if patient_name_upper in di_text.upper():
                                    images = di.descendants(control_type='Image')
                                    if images:
                                        img_rect = images[0].rectangle()
                                        chart_cx = (img_rect.left + img_rect.right) // 2
                                        chart_cy = (img_rect.top + img_rect.bottom) // 2
                                    break
                            except:
                                continue
                    except:
                        pass
                    # Last resort: use row rect left edge (chart icon is typically left of name)
                    if chart_cx is None and row_rect:
                        chart_cx = row_rect.left + 20
                        chart_cy = cy
                    elif chart_cx is None:
                        control.add_log("Could not determine chart icon position")
                        raise Exception("No chart icon position available")

                pyautogui.click(chart_cx, chart_cy)
                control.add_log(f"Clicked chart icon at ({chart_cx}, {chart_cy})")

                chart_title_re = f'.*{re.escape(name.split(",")[0])}.*'
                chart_win = wait_for_window(chart_title_re, timeout=8)
                if not chart_win:
                    dismiss_popup_dialogs()
                    chart_win = wait_for_window(chart_title_re, timeout=4)

                if not chart_win:
                    control.add_log("Chart didn't open")
                    raise Exception("Chart failed to open")

                # Wait for Demographics tab to actually exist before clicking
                _roomed_tab_clicked = False
                _roomed_deadline = time.time() + 10  # 10 second timeout
                while time.time() < _roomed_deadline:
                    try:
                        tab_items = chart_win.descendants(control_type='TabItem')
                        for elem in tab_items:
                            tab_name = elem.element_info.name or ''
                            if 'demograph' in tab_name.lower():
                                # Verify the element is visible and has a valid rect
                                rect = elem.rectangle()
                                if rect.right > rect.left and rect.bottom > rect.top:
                                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                    _roomed_tab_clicked = True
                                    control.add_log("Demographics tab found and clicked")
                                    break
                    except Exception:
                        pass
                    if _roomed_tab_clicked:
                        break
                    time.sleep(0.2)
                if not _roomed_tab_clicked:
                    control.add_log("Demographics tab never appeared after 10s - skipping insurance")
                    raise Exception("Demographics tab not found")

                deadline = time.time() + 3
                texts_with_pos = []
                while time.time() < deadline:
                    try:
                        all_texts = chart_win.descendants(control_type='Text')
                        if len(all_texts) > 10:
                            for t in all_texts:
                                try:
                                    txt = t.element_info.name or ''
                                    rect = t.rectangle()
                                    if txt:
                                        texts_with_pos.append({'txt': txt, 'x': rect.left, 'y': rect.top})
                                except:
                                    pass
                            break
                    except:
                        pass
                    time.sleep(0.05)

                control.add_log(f"Chart has {len(texts_with_pos)} text elements")

                # Extract Race/Ethnicity/Language from Cultural Information section only
                cultural_y = None
                for t in texts_with_pos:
                    if t['txt'] == 'Cultural Information':
                        cultural_y = t['y']
                        break

                if cultural_y is not None:
                    cultural_texts = [e for e in texts_with_pos if e['y'] > cultural_y and e['y'] < cultural_y + 150]
                    for label_key, data_key in [('Race:', 'race'), ('Ethnicity:', 'ethnicity'), ('Language:', 'language')]:
                        for tp in cultural_texts:
                            if tp['txt'] == label_key:
                                ly, lx = tp['y'], tp['x']
                                for vp in cultural_texts:
                                    if abs(vp['y'] - ly) < 20 and vp['x'] > lx + 30 and vp['txt'] != label_key:
                                        patient_data[data_key] = vp['txt']
                                        control.add_log(f"{data_key}: {vp['txt']}")
                                        break
                                break
                else:
                    control.add_log("Cultural Information section not found")

                # Extract Primary Insurance
                for t in texts_with_pos:
                    if t['txt'] == 'Primary Insurance:':
                        iy, ix = t['y'], t['x']
                        for v in texts_with_pos:
                            if abs(v['y'] - iy) < 20 and v['x'] > ix + 50:
                                patient_data['insurance'] = v['txt']
                                control.add_log(f"Insurance: {v['txt']}")
                                break
                        break

                # Close chart and WAIT for it to disappear
                close_chart_window(chart_win)
                wait_for_window_close(chart_title_re, timeout=3)
                control.add_log("Closed chart")

            except Exception as chart_err:
                control.add_log(f"Chart error: {str(chart_err)[:30]}")

            # Mark as processed
            processed_patients.add(name)
            update_patient_log(f"  >> Extracted data for roomed patient: {name}")
            update_extracted_data(patient_data)
            # Create MHT API JSON file for roomed patient
            json_path = create_mht_api_json(patient_data)
            if json_path:
                update_patient_log(f"  >> MHT API JSON created: {json_path.split('/')[-1]}")
            return patient_data

        except Exception as e:
            control.add_log(f"Error processing roomed patient: {str(e)[:40]}")
            return None

    # === MONITORING LOOP ===
    cycle = 0
    _last_heartbeat = 0
    control.add_log("Starting monitoring loop")

    try:
        while not control.is_killed:
            cycle += 1
            control.add_log(f"--- Cycle {cycle} ---")
            # Heartbeat every 30 seconds
            if _claimed and time.time() - _last_heartbeat > 30:
                try:
                    heartbeat_slot(_db_path, _slot_name)
                    _last_heartbeat = time.time()
                except Exception:
                    pass
            # Start new cycle - clears display and starts fresh
            update_patient_log(f"=== Cycle {cycle} ===", new_cycle=cycle)

            # === CLEAR STALE OVERLAYS BEFORE RESCAN ===
            clear_patient_overlays()

            # === VERIFY TRACKING BOARD DATA IS LOADED ===
            # On first cycle or after refresh, the UI elements may not be rendered yet
            _tb_ready_deadline = time.time() + 3
            while time.time() < _tb_ready_deadline:
                try:
                    _rp = win.child_window(title_re='.*Roomed Patients.*', control_type='Group')
                    _rp.descendants(control_type='DataItem')
                    break  # Group exists and is queryable
                except:
                    time.sleep(0.15)

            # === PRIORITIZE: ROOMED PATIENTS FIRST ===
            # Check for unprocessed roomed patients before looking at waiting room
            current_roomed = find_roomed_patients()
            control.add_log(f"Roomed: {len(current_roomed)} patients")

            # Build current roomed dict
            current_roomed_dict = {}
            for p in current_roomed:
                rname = p.get('name', 'Unknown')
                room = p.get('room', '?')
                rage = p.get('age', 0)
                time_in = p.get('time', '')
                current_roomed_dict[rname] = {'room': room, 'age': rage, 'time': time_in}

            # Check for discharged patients
            if cycle > 1:
                for rname, info in previous_roomed_patients.items():
                    if rname not in current_roomed_dict:
                        room = info.get('room', '?')
                        rage = info.get('age', 0)
                        discharged_time = time.strftime('%H:%M')
                        update_discharged(f"[D] {rname} ({rage}y) from {room} @ {discharged_time}")
                        control.add_log(f"DISCHARGED: {rname}")
                        update_patient_log(f"<< DISCHARGED: {rname}")
                        analytics.record_discharged_patient(rname)
                        expire_patient_assessment(rname)

            # Update previous roomed
            previous_roomed_patients.clear()
            previous_roomed_patients.update(current_roomed_dict)

            # === WAITING ROOM PATIENTS ===
            analytics.start_stage(ProcessingStage.WAITING_ROOM_SCAN)
            patients = find_patient_rows()
            control.add_log(f"Waiting Room: {len(patients)} patients")
            update_patient_log(f"Waiting Room: {len(patients)}")

            # Count qualified patients for analytics
            qualified_count = sum(1 for p in patients if check_qualification(p)[0])
            analytics.end_stage(ProcessingStage.WAITING_ROOM_SCAN, success=True)
            analytics.record_waiting_room_scan(len(patients), qualified_count)

            # Show overlays (this also sets qualified/reason on each patient)
            if patients:
                show_patient_overlays(patients)

            # Build current patient dict and log details
            current_patients = {}
            for p in patients:
                name = p.get('name', 'Unknown')
                age = p.get('age', 0)
                qualified = p.get('qualified', False)
                reason = p.get('reason', 'Unknown')

                current_patients[name] = {'age': age, 'qualified': qualified, 'reason': reason}

                icon = "+" if qualified else "X"
                status = "OK" if qualified else "NO"

                if name not in tracked_patients:
                    tracked_patients[name] = {'first_seen': cycle, 'count': 1, 'qualified': qualified}
                else:
                    tracked_patients[name]['count'] += 1

                # Single concise line: [+] NAME (Age) - Reason
                update_patient_log(f"[{icon}] {name} ({age}y) - {reason}")
                control.add_log(f"{name}: {status}")

            # === PROCESS QUALIFIED PATIENTS ONE AT A TIME ===
            # Keep processing until no more new qualified patients
            while not control.is_killed:
                next_patient = find_next_qualified_patient(patients)
                if not next_patient:
                    # No more new qualified patients to process
                    break

                name = next_patient.get('name', 'Unknown')
                control.add_log(f"Processing qualified patient: {name}")
                update_patient_log(f"Processing: {name}...")

                # Process this one patient
                data = process_one_qualified_patient(next_patient)

                if data:
                    control.add_log(f"Successfully extracted data for {name}")

                    # In demo mode, quick check for outbound after each extraction
                    if demo_mode and outbound_worker:
                        processed = outbound_worker.process_pending()
                        if processed > 0:
                            control.add_log(f"Outbound: processed {processed} event(s)")
                            demo_status_overlay.update_status("Extraction complete", f"Processed {processed} event(s)")

                # Quick re-scan (no sleep, just clear and rescan)
                clear_patient_overlays()
                patients = find_patient_rows()
                if patients:
                    show_patient_overlays(patients)

            # Check for patients moved from Waiting Room to Roomed
            # Also track qualified patients that moved before we could extract their data
            qualified_moved_to_roomed = []
            if cycle > 1:
                for name, info in previous_patients.items():
                    if name not in current_patients:
                        age = info['age']
                        qualified = info['qualified']
                        icon = "+" if qualified else "X"
                        roomed_patients.append(info)
                        roomed_time = time.strftime('%H:%M')
                        update_roomed(f"[{icon}] {name} ({age}y) @ {roomed_time}")
                        control.add_log(f"ROOMED: {name}")
                        update_patient_log(f">> ROOMED: {name}")
                        analytics.record_roomed_patient(name)

                        # Track qualified patients that moved before we processed them
                        if qualified and name not in processed_patients:
                            qualified_moved_to_roomed.append({'name': name, 'age': age, 'qualified': qualified})
                            control.add_log(f"EDGE CASE: Qualified patient {name} moved to roomed before extraction")

            # === EDGE CASE: Process qualified patients that moved to roomed before extraction ===
            if qualified_moved_to_roomed and not control.is_killed:
                control.add_log(f"Processing {len(qualified_moved_to_roomed)} qualified patients from roomed section")
                update_patient_log(f">> {len(qualified_moved_to_roomed)} qualified patients moved to roomed - processing...")

                try:
                    # Reconnect to tracking board for roomed patient processing
                    from pywinauto import findwindows
                    elements = _sfind(title_re='.*Tracking Board.*', backend='uia')
                    target_handle = None
                    for elem in elements:
                        if 'Setup' not in elem.name:
                            target_handle = elem.handle
                            break

                    if target_handle:
                        tb_app = Application(backend='uia').connect(handle=target_handle, timeout=5)
                        tb_win = tb_app.window(handle=target_handle)

                        for roomed_patient in qualified_moved_to_roomed:
                            if control.is_killed:
                                break
                            patient_name = roomed_patient.get('name', 'Unknown')
                            control.add_log(f"Processing roomed qualified patient: {patient_name}")

                            # Process this roomed patient
                            data = process_roomed_qualified_patient(roomed_patient, tb_win)

                            if data:
                                control.add_log(f"Successfully extracted data for roomed patient {patient_name}")

                            time.sleep(0.3)  # Brief pause between patients

                except Exception as roomed_proc_err:
                    control.add_log(f"Error processing roomed patients: {str(roomed_proc_err)[:40]}")

            # Update previous patients for next cycle
            previous_patients = current_patients.copy()

            # Countdown to refresh - process outbound during this downtime
            for remaining in range(REFRESH_INTERVAL, 0, -1):
                if control.is_killed:
                    break
                # Process pending outbound events during refresh wait
                if outbound_worker and remaining > 3:  # Leave 3s buffer before refresh
                    processed = outbound_worker.process_pending()
                    if processed > 0:
                        control.add_log(f"Processed {processed} outbound event(s)")
                        break  # Exit countdown after processing to refresh sooner
                control.set_step(f"Next refresh in {remaining}s")
                time.sleep(1)

            if control.is_killed:
                break

            # Click Refresh - dismiss any error popups if they appear
            control.set_step("Refreshing...")
            pyautogui.click(REFRESH_X, REFRESH_Y)

            # Quick check for error popup (no fixed sleep - just poll once)
            refresh_ok = True
            for check in range(2):
                try:
                    from pywinauto import Desktop
                    desktop = Desktop(backend='uia')
                    for w in desktop.windows():
                        title = w.window_text()
                        if 'Application Error' in title or 'Error' == title:
                            control.add_log(f"[REFRESH] Error detected - dismissing")
                            buttons = w.descendants(control_type='Button')
                            for btn in buttons:
                                if btn.window_text() in ['OK', 'Ok']:
                                    btn.click_input()
                                    time.sleep(0.15)
                                    break
                            pyautogui.click(REFRESH_X, REFRESH_Y)
                            refresh_ok = False
                            break
                except:
                    pass
                if refresh_ok:
                    break
                time.sleep(0.1)

            # Wait for tracking board data to actually reload after refresh
            # Poll for Waiting Room group to have DataItem children (patients rendered)
            _reload_deadline = time.time() + 4  # up to 4 seconds
            _data_loaded = False
            while time.time() < _reload_deadline:
                try:
                    _wr = win.child_window(title_re='.*Waiting Room.*', control_type='Group')
                    _items = _wr.descendants(control_type='DataItem')
                    if len(_items) > 0:
                        _data_loaded = True
                        break
                except:
                    pass
                time.sleep(0.15)

            if not _data_loaded:
                # Waiting room might genuinely be empty, or still loading — try one more refresh
                control.add_log("Tracking board data not ready after refresh - retrying...")
                pyautogui.click(REFRESH_X, REFRESH_Y)
                _retry_deadline = time.time() + 3
                while time.time() < _retry_deadline:
                    try:
                        _wr = win.child_window(title_re='.*Waiting Room.*', control_type='Group')
                        _items = _wr.descendants(control_type='DataItem')
                        if len(_items) > 0:
                            break
                    except:
                        pass
                    time.sleep(0.15)

            control.add_log("Refresh complete")

    except Exception as e:
        control.add_log(f"Monitoring error: {str(e)}")
        import traceback
        error_trace = traceback.format_exc()[:200]
        control.add_log(error_trace)
        analytics.record_error("monitoring_loop", str(e), error_trace)

    finally:
        control.set_status("Stopped")
        control.set_step("Monitoring ended")
        control.add_log("=== MONITORING STOPPED ===")
        update_patient_log("\n=== STOPPED ===")

        # End analytics session and save data
        try:
            analytics.end_session()
            control.add_log("Analytics session saved")

            # Export analytics summary to log
            summary = analytics.get_log_summary()
            control.add_log(summary)

            # Export detailed logs
            log_path = analytics.export_logs()
            control.add_log(f"Analytics logs exported: {log_path}")
        except Exception as analytics_err:
            control.add_log(f"Analytics save error: {str(analytics_err)[:40]}")

        # Cleanup all overlays
        if demo_mode:
            if demo_status_overlay:
                demo_status_overlay.stop()
            if demo_extracted_overlay:
                demo_extracted_overlay.stop()
        else:
            clear_patient_overlays()  # Clear patient row overlays first
            for root in [wr_overlay_root, patient_log_root, roomed_overlay_root, discharged_overlay_root, patient_overlay_root, extracted_data_root]:
                if root:
                    try:
                        root.after(0, root.quit)
                    except:
                        pass
        time.sleep(0.5)  # Give overlays time to close


def _change_location(control, target_location="ATTALLA"):
    """Change clinic location via menu navigation.
    Flow: CLICK Clinic menu → HOVER Current Clinic (opens submenu) → CLICK target location.
    Uses pywinauto element detection first, falls back to proportional window offsets.
    """
    import pyautogui
    from pywinauto import Application, findwindows
    try:
        from mhtagentic.desktop.session_guard import session_find_elements as _sfind
    except ImportError:
        _sfind = findwindows.find_elements

    try:
        # Connect to the Tracking Board / Experity window
        elements = _sfind(title_re='.*Tracking Board.*', backend='uia')
        target_handle = None
        for elem in elements:
            if 'Setup' not in elem.name:
                target_handle = elem.handle
                break

        if not target_handle:
            control.add_log("Tracking Board not found for location change")
            return False

        app = Application(backend='uia').connect(handle=target_handle, timeout=3)
        win = app.window(handle=target_handle)
        win_rect = win.rectangle()
        w_left, w_top = win_rect.left, win_rect.top
        w_width = win_rect.right - win_rect.left
        w_height = win_rect.bottom - win_rect.top
        control.add_log(f"Window rect: {w_left},{w_top} size={w_width}x{w_height}")

        # Strategy 1: Try pywinauto element detection for clinic menu
        clinic_found = False
        try:
            for elem in win.descendants():
                try:
                    name = (elem.element_info.name or '').lower()
                    if 'clinic' in name and elem.element_info.control_type in ('MenuItem', 'Hyperlink', 'Button', 'ListItem'):
                        rect = elem.rectangle()
                        cx, cy = (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2
                        control.add_log(f"Found clinic element: '{elem.element_info.name}' at ({cx},{cy})")
                        show_target(cx, cy, "CLINIC MENU", color="#2196F3", width=rect.right - rect.left + 20, height=rect.bottom - rect.top + 10, duration=0.8)
                        time.sleep(0.6)
                        pyautogui.click(cx, cy)
                        clinic_found = True
                        break
                except:
                    continue
        except:
            pass

        # Strategy 2: Proportional window offsets (scales with window size)
        if not clinic_found:
            control.add_log("Clinic element not found via UIA, using proportional offsets")
            clinic_x = w_left + int(w_width * 0.069)
            clinic_y = w_top + int(w_height * 0.047)
            show_target(clinic_x, clinic_y, "CLINIC MENU", color="#2196F3", width=100, height=30, duration=0.8)
            time.sleep(0.6)
            pyautogui.click(clinic_x, clinic_y)

        control.add_log("Clicked Clinic menu")
        time.sleep(1.0)

        # Step 2: HOVER Current Clinic (opens submenu - DO NOT click!)
        control.set_step("Hovering Current Clinic...")

        # Try element detection for Current Clinic menu item
        current_found = False
        try:
            for elem in win.descendants():
                try:
                    name = (elem.element_info.name or '').lower()
                    if 'current' in name and 'clinic' in name:
                        rect = elem.rectangle()
                        cx, cy = (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2
                        control.add_log(f"Found current clinic: '{elem.element_info.name}' at ({cx},{cy})")
                        show_target(cx, cy, "CURRENT CLINIC", color="#FF9800", width=rect.right - rect.left + 20, height=rect.bottom - rect.top + 10, duration=0.8)
                        time.sleep(0.6)
                        pyautogui.moveTo(cx, cy, duration=0.3)
                        current_found = True
                        break
                except:
                    continue
        except:
            pass

        if not current_found:
            current_x = w_left + int(w_width * 0.087)
            current_y = w_top + int(w_height * 0.141)
            show_target(current_x, current_y, "CURRENT CLINIC", color="#FF9800", width=120, height=30, duration=0.8)
            time.sleep(0.6)
            pyautogui.moveTo(current_x, current_y, duration=0.3)

        control.add_log("Hovering over Current Clinic")
        time.sleep(1.0)

        # Step 3: CLICK target location
        control.set_step(f"Selecting {target_location}...")

        # Try element detection for target location
        location_found = False
        try:
            for elem in win.descendants():
                try:
                    name = (elem.element_info.name or '').lower()
                    if target_location.lower() in name:
                        rect = elem.rectangle()
                        cx, cy = (rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2
                        control.add_log(f"Found {target_location}: '{elem.element_info.name}' at ({cx},{cy})")
                        show_target(cx, cy, target_location, color="#4CAF50", width=rect.right - rect.left + 20, height=rect.bottom - rect.top + 10, duration=0.8)
                        time.sleep(0.6)
                        pyautogui.click(cx, cy)
                        location_found = True
                        break
                except:
                    continue
        except:
            pass

        if not location_found:
            loc_x = w_left + int(w_width * 0.156)
            loc_y = w_top + int(w_height * 0.171)
            show_target(loc_x, loc_y, target_location, color="#4CAF50", width=80, height=30, duration=0.8)
            time.sleep(0.6)
            pyautogui.click(loc_x, loc_y)

        control.add_log(f"Clicked {target_location} - location changed!")
        time.sleep(1.5)

        control.add_log("Location change completed successfully")
        return True

    except Exception as e:
        control.add_log(f"Location change error: {str(e)}")
        return False

def run_monitor_only(mode="inbound"):
    """Run in monitor-only mode — skip launch/login, connect to existing Experity window.

    When launched by start_all_clean.py .bat chain, writes status codes 120-200
    to the per-RDP status file for the orchestrator to track progress.

    Args:
        mode: "inbound" to monitor Waiting Room, "outbound" to process assessment results.
    """
    global _automation, _control

    print(f"[monitor-only] starting mode={mode}", flush=True)

    # --- Post-login status reporting setup ---
    rdp_label = os.environ.get("MHT_RDP_LABEL", "")
    bot_location = os.environ.get("BOT_LOCATION", "ATTALLA")
    bot_role = os.environ.get("BOT_ROLE", "inbound")
    status_file = Path(rf"C:\ProgramData\MHTAgentic\clean_status_{rdp_label}.txt") if rdp_label else None

    def _post_status(code, msg=""):
        if status_file:
            status_file.write_text(f"{code} {msg}".strip())
            print(f"[monitor-only] STATUS {code}: {msg}", flush=True)

    # --- Phase 1: Wait for EMR / Tracking Board (status 120-130) ---
    _post_status(120, "Waiting for Tracking Board")

    # Show bottom-middle overlay immediately with role assignment
    demo_overlay = DemoStatusOverlay()
    demo_overlay.start()
    role_label = "Outbound" if mode == "outbound" else "Inbound"
    demo_overlay.update_status(inbound_text=f"{role_label} — Waiting for EMR...")

    try:
        import pyautogui
        from pywinauto import findwindows, Application
        try:
            from mhtagentic.desktop.session_guard import session_find_elements
        except ImportError:
            session_find_elements = findwindows.find_elements

        # Poll for Tracking Board window (up to 120s — EMR can be slow after login)
        tb_found = False
        for attempt in range(120):
            elements = session_find_elements(title_re='.*Tracking Board.*', backend='uia')
            if elements:
                target_handle = None
                for elem in elements:
                    if 'Setup' not in elem.name:
                        target_handle = elem.handle
                        break
                if target_handle:
                    app = Application(backend='uia').connect(handle=target_handle, timeout=2)
                    win = app.window(handle=target_handle)
                    win_rect = win.rectangle()
                    click_x = (win_rect.left + win_rect.right) // 2
                    click_y = win_rect.top + 30
                    pyautogui.click(click_x, click_y)
                    print(f"[monitor-only] Tracking Board found, refocused at ({click_x}, {click_y})", flush=True)
                    tb_found = True
                    break
            if attempt % 10 == 0 and attempt > 0:
                print(f"[monitor-only] Still waiting for Tracking Board ({attempt}s)...", flush=True)
            time.sleep(1)

        if not tb_found:
            _post_status(-1, "Tracking Board not found after 120s")
            demo_overlay.update_status(inbound_text="ERROR: EMR not found")
            time.sleep(5)
            try:
                demo_overlay.stop()
            except Exception:
                pass
            return

        _post_status(130, "EMR found")
        demo_overlay.update_status(inbound_text="EMR found")

    except Exception as emr_err:
        _post_status(-1, f"EMR wait error: {emr_err}")
        try:
            demo_overlay.stop()
        except Exception:
            pass
        raise

    # --- Phase 2: Switch location (status 140-150) ---
    _post_status(140, f"Switching to {bot_location}")
    demo_overlay.update_status(inbound_text=f"Switching to {bot_location}...")

    try:
        import concurrent.futures
        with concurrent.futures.ThreadPoolExecutor() as pool:
            future = pool.submit(_change_location, demo_overlay, bot_location)
            switch_ok = future.result(timeout=60)
    except Exception as loc_err:
        print(f"[monitor-only] Location switch error/timeout: {loc_err}", flush=True)
        switch_ok = False

    if switch_ok:
        _post_status(150, f"Location: {bot_location}")
        demo_overlay.update_status(inbound_text=f"Location: {bot_location}")
    else:
        print(f"[monitor-only] Location switch failed, continuing with current location", flush=True)
        _post_status(150, f"Location switch skipped")

    # --- Phase 3: Claim slot + signal orchestrator (status 160) ---
    _post_status(160, "Location confirmed")
    demo_overlay.update_status(inbound_text=f"Monitoring {bot_location}")

    # Claim the bot slot in the database so dashboard shows active status
    from mhtagentic import OUTPUT_DIR
    from mhtagentic.db import claim_slot, heartbeat_slot, release_slot
    _db_path = str(OUTPUT_DIR / "mht_data.db")
    _slot_name = os.environ.get("MHT_RDP_LABEL", "RDP1").replace("RDP", "experity").lower()
    # Map RDP1->experityb, RDP2->experityc, RDP3->experityd
    _slot_map = {"rdp1": "experityb", "rdp2": "experityc", "rdp3": "experityd"}
    _slot_name = _slot_map.get(os.environ.get("MHT_RDP_LABEL", "RDP1").lower(), "experityb")
    _claimed = claim_slot(_db_path, _slot_name)
    if _claimed:
        print(f"[monitor-only] Claimed slot: {_claimed}", flush=True)
        # Update slot with role and location
        try:
            import sqlite3 as _sql
            _conn = _sql.connect(_db_path, timeout=10)
            _conn.execute(
                "UPDATE bot_slot SET role = ?, location = ? WHERE slot_name = ?",
                (mode, bot_location, _slot_name),
            )
            _conn.commit()
            _conn.close()
        except Exception as _e:
            print(f"[monitor-only] Slot update error: {_e}", flush=True)
    else:
        print(f"[monitor-only] WARNING: Could not claim slot {_slot_name}", flush=True)

    _should_exit = False

    def on_kill():
        nonlocal _should_exit
        _should_exit = True

    # Use DemoStatusOverlay as _control for monitoring
    print(f"[monitor-only] {mode}: keeping DemoStatusOverlay as _control", flush=True)
    _control = demo_overlay

    try:
        print(f"[monitor-only] bot_location={bot_location}", flush=True)
        try:
            _control.set_status("Connecting")
            _control.set_step(f"Finding Experity window ({mode})...")
            _control.add_log(f"Monitor-only mode={mode}, location={bot_location}")
        except Exception as overlay_err:
            print(f"[monitor-only] overlay error (non-fatal): {overlay_err}", flush=True)

        # Find existing Experity window (don't launch it)
        _automation = DesktopAutomation()

        if not _automation.find_experity_window():
            print("[monitor-only] Experity window NOT found!", flush=True)
            _control.set_status("Error")
            _control.set_step("Experity window not found")
            time.sleep(5)
            _control.stop()
            return
        print(f"[monitor-only] Found Experity window: {_automation.window.title}", flush=True)

        _post_status(200, "Monitoring active")

        _control.set_status("Connected")
        _control.set_step("Found Experity window")
        _control.add_log("Experity window found")
        _control.add_log(f"Location: {bot_location}")

        if _should_exit:
            _control.stop()
            return

        if mode == "outbound":
            logging.basicConfig(
                level=logging.INFO,
                format='[%(name)s] %(message)s',
                stream=sys.stdout,
                force=True
            )
            from mhtagentic.outbound.outbound_worker import OutboundWorker
            from mhtagentic.db.mht_simulator import MHTResponseSimulator
            from mhtagentic import OUTPUT_DIR as _OUT_DIR
            db_path = str(_OUT_DIR / "mht_data.db")

            simulator = MHTResponseSimulator(db_path, response_delay_seconds=0)
            simulator.start()
            print("[outbound] MHT Simulator started (instant)", flush=True)
            _control.add_log("Simulator started (instant)")

            worker = OutboundWorker(db_path, poll_interval=5.0, overlay=_control)
            _control.set_status("Outbound")
            _control.set_step("Waiting for outbound events...")
            _control.add_log(f"Outbound polling loop at {bot_location}")
            print("[monitor-only] entering outbound polling loop", flush=True)

            # Find Refresh button on Tracking Board for periodic refresh
            import pyautogui
            from pywinauto.application import Application as _RefApp
            _ob_refresh_x, _ob_refresh_y = None, None
            try:
                try:
                    from mhtagentic.desktop.session_guard import session_find_elements as _ob_sfind
                except ImportError:
                    from pywinauto import findwindows
                    _ob_sfind = findwindows.find_elements
                _ob_elems = _ob_sfind(title_re='.*Tracking Board.*', backend='uia')
                _ob_tb_handle = None
                for _obe in _ob_elems:
                    if 'Setup' not in _obe.name:
                        _ob_tb_handle = _obe.handle
                        break
                if _ob_tb_handle:
                    _tb_app = _RefApp(backend='uia').connect(handle=_ob_tb_handle, timeout=3)
                    _tb_win = _tb_app.window(handle=_ob_tb_handle)
                    _ref_btn = _tb_win.child_window(title='Refresh', control_type='Button')
                    _ref_rect = _ref_btn.rectangle()
                    _ob_refresh_x = (_ref_rect.left + _ref_rect.right) // 2
                    _ob_refresh_y = (_ref_rect.top + _ref_rect.bottom) // 2
                    print(f"[outbound] Refresh button at ({_ob_refresh_x}, {_ob_refresh_y})", flush=True)
            except Exception as _ref_err:
                print(f"[outbound] Could not find Refresh button: {_ref_err}", flush=True)

            _OB_REFRESH_INTERVAL = 15
            _ob_last_refresh = time.time()

            try:
                while True:
                    try:
                        processed = worker.process_pending()
                        if processed > 0:
                            _control.set_step(f"Processed {processed} event(s)")
                            print(f"[outbound] processed {processed} event(s)", flush=True)
                            _ob_last_refresh = time.time()
                        else:
                            # Periodic Tracking Board refresh (same as inbound)
                            _now = time.time()
                            _since = _now - _ob_last_refresh
                            if _since >= _OB_REFRESH_INTERVAL and _ob_refresh_x:
                                _control.set_step("Refreshing Tracking Board...")
                                pyautogui.click(_ob_refresh_x, _ob_refresh_y)
                                _ob_last_refresh = time.time()
                                time.sleep(0.5)
                                _control.set_step("Waiting for outbound events...")
                                print("[outbound] Tracking Board refreshed", flush=True)
                            else:
                                _remaining = max(0, int(_OB_REFRESH_INTERVAL - _since))
                                _control.set_step(f"Waiting for events... refresh in {_remaining}s")
                    except Exception as poll_err:
                        print(f"[outbound] poll error: {poll_err}", flush=True)
                    time.sleep(5)
            finally:
                simulator.stop()
                print("[outbound] simulator stopped", flush=True)
        else:
            # Inbound mode: monitor Waiting Room
            _control.set_status("Monitoring")
            _control.set_step(f"Starting inbound ({bot_location})...")
            _control.add_log(f"Starting inbound monitoring at {bot_location}")
            time.sleep(1.0)
            _start_monitoring(_control, demo_mode=True)

    except BaseException as e:
        print(f"[monitor-only] ERROR: {type(e).__name__}: {str(e)}", flush=True)
        import traceback
        print(traceback.format_exc(), flush=True)
        try:
            _control.set_status("Error")
            _control.set_step(str(e)[:50])
            _control.add_log(f"ERROR: {str(e)}")
            _control.add_log(traceback.format_exc())
        except:
            pass
        time.sleep(5)

    finally:
        print("[monitor-only] cleanup", flush=True)
        try:
            _control.stop()
        except:
            pass


if __name__ == "__main__":
    print(f"[launcher] argv={sys.argv}", flush=True)
    args = sys.argv[1:]

    if "--monitor-only" in args:
        # Monitor-only mode: skip launch/login, connect to existing Experity
        mode = "inbound"
        if "-outbound" in args:
            mode = "outbound"
        print(f"[launcher] monitor-only mode={mode}", flush=True)
        try:
            run_monitor_only(mode=mode)
        except Exception as e:
            import traceback
            print(f"[launcher] CRASH: {e}", flush=True)
            print(traceback.format_exc(), flush=True)
            error_file = SCRIPT_DIR / "output" / "debug" / "crash_log.txt"
            error_file.parent.mkdir(parents=True, exist_ok=True)
            with open(error_file, "w") as f:
                f.write(f"CRASH LOG (monitor-only)\n")
                f.write(f"Error: {str(e)}\n")
                f.write(f"Traceback:\n{traceback.format_exc()}\n")
    else:
        # run_silent() DISABLED — use start_all_clean.py instead
        print("[launcher] ERROR: run_silent() is disabled. Use start_all_clean.py (Start All) instead.", flush=True)
        sys.exit(1)