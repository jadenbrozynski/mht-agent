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
from pathlib import Path
from datetime import datetime

# Add project root to path
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(SCRIPT_DIR))

from mhtagentic.desktop.control_overlay import (
    ControlOverlay,
    reset_control_overlay,
    register_overlay_window,
    unregister_overlay_window
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


def save_debug():
    """Save screenshot and logs for debugging."""
    global _automation, _control

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    debug_file = _debug_dir / f"debug_{timestamp}"

    try:
        import pyautogui
        screenshot = pyautogui.screenshot()
        screenshot.save(f"{debug_file}_screen.png")
    except:
        pass

    if _control:
        logs = _control.get_logs()
        with open(f"{debug_file}_logs.txt", "w") as f:
            f.write(f"MHT Agentic Debug Log\n")
            f.write(f"Timestamp: {timestamp}\n")
            f.write(f"=" * 50 + "\n\n")
            for log in logs:
                f.write(log + "\n")
        _control.set_step(f"Debug saved: debug_{timestamp}")


def toggle_recording():
    """Toggle macro recording on/off."""
    global _recorder, _control

    try:
        from mhtagentic.desktop.macro_recorder import MacroRecorder
    except ImportError:
        _control.set_step("pynput not installed!")
        return

    if _recorder is None:
        _recorder = MacroRecorder(on_action=lambda msg: _control.set_step(msg) if _control else None)

    if _control.is_recording:
        # Start recording
        _recorder.start_recording()
        _control.add_log("Macro recording STARTED")
    else:
        # Stop recording
        _recorder.stop_recording()
        _control.add_log("Macro recording STOPPED and saved")


def run_silent():
    """Run MHT Agentic silently - auto-launches Experity and runs saved macro."""
    global _automation, _control

    reset_control_overlay()

    _should_exit = False

    def on_kill():
        nonlocal _should_exit
        _should_exit = True

    _control = ControlOverlay(
        on_kill=on_kill,
        on_debug=save_debug,
        on_record=toggle_recording
    )
    _control.start()
    time.sleep(0.3)

    try:
        # Launch Experity immediately
        _control.set_status("Launching")
        _control.set_step("Opening Experity EMR...")
        _control.add_log("Launching Experity EMR...")

        _automation = DesktopAutomation()
        launched = _automation.launch_experity(wait_seconds=120)

        if _should_exit:
            _control.stop()
            return

        if not launched:
            _control.add_log("Launch returned False, searching for window...")
            if not _automation.find_experity_window():
                _control.set_status("Error")
                _control.set_step("Could not find Experity window")
                _control.add_log("ERROR: Could not find Experity window")
                time.sleep(3)
                _control.stop()
                return

        _control.set_status("Connected")
        _control.set_step("Found Experity window")
        _control.add_log("Experity window found")

        region = _automation.get_window_region()
        if region:
            _control.add_log(f"Window region: {region}")

        # Auto-run fast login
        try:
            _control.add_log("Starting fast login")
        except:
            pass
        print("Starting fast login")

        time.sleep(0.5)
        _fast_login(_control)

        try:
            _control.set_step("Click Continue to close")
            _control.add_log("Flow completed - click Continue to close")
            _control.enable_proceed()
            _control.wait_for_proceed(timeout=300)  # Wait up to 5 minutes
        except:
            print("Flow completed")
            time.sleep(10)  # Give user time to see result

    except Exception as e:
        print(f"ERROR: {str(e)}")
        try:
            _control.set_status("Error")
            _control.set_step(str(e)[:50])
            _control.add_log(f"ERROR: {str(e)}")
            import traceback
            _control.add_log(traceback.format_exc())
            _control.enable_proceed()
            _control.wait_for_proceed(timeout=300)  # Wait so user can see error
        except:
            import traceback
            print(traceback.format_exc())
            time.sleep(10)

    finally:
        try:
            _control.stop()
        except:
            pass


# Stored credentials
STORED_USERNAME = "RCALLAGHAN@STHRN"
STORED_PASSWORD = "Mental@2026!!"


def _show_element_overlay(element, label, color="#2196F3", duration=1.0):
    """Show an overlay over a pywinauto element."""
    try:
        rect = element.rectangle()
        center_x = (rect.left + rect.right) // 2
        center_y = (rect.top + rect.bottom) // 2
        width = rect.right - rect.left + 20
        height = rect.bottom - rect.top + 10
        show_target(center_x, center_y, label, color=color, width=width, height=height, duration=duration)
    except:
        pass


def _start_monitoring(control, outbound_worker=None):
    """Start monitoring the Waiting Room with periodic refresh and patient tracking."""
    import pyautogui
    import threading
    import tkinter as tk
    from pywinauto import Application, findwindows
    import re

    REFRESH_INTERVAL = 15  # 15 seconds
    tracked_patients = {}  # Store patient data
    previous_roomed_patients = {}  # Track roomed patients for discharge detection

    # Database for event tracking
    db = MHTDatabase(SCRIPT_DIR / "output" / "mht_data.db")
    patient_event_ids = {}  # {patient_name_upper: event_id}

    # === ELEMENT DETECTION HELPERS (replace fixed sleeps) ===
    def wait_for_window(title_re, timeout=5):
        """Wait for a window to appear, return it immediately when found. No fixed sleep."""
        deadline = time.time() + timeout
        while time.time() < deadline:
            try:
                elements = findwindows.find_elements(title_re=title_re, backend='uia')
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
                elements = findwindows.find_elements(title_re=title_re, backend='uia')
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

    # === BACKGROUND ERROR DETECTION THREAD ===
    # Runs continuously to catch Application Error popups without blocking main process
    error_monitor_running = True

    def background_error_monitor():
        """Background thread that continuously monitors for error popups."""
        from pywinauto import Desktop
        while error_monitor_running and not control.is_killed:
            try:
                desktop = Desktop(backend='uia')
                windows = desktop.windows()
                for w in windows:
                    try:
                        title = w.window_text()
                        if 'Application Error' in title or 'Error' == title:
                            # Found error popup - click OK
                            buttons = w.descendants(control_type='Button')
                            for btn in buttons:
                                if btn.window_text() in ['OK', 'Ok', 'Close']:
                                    btn.click_input()
                                    control.add_log(f"[ERROR MONITOR] Auto-dismissed: {title[:30]}")
                                    time.sleep(0.3)
                                    break
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
                            if rect.left > 400 and rect.right < 1900 and rect.top > 150 and rect.bottom < 1050:
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
        mht_api_payload = {
            "clinic_id": 110,  # Southern Immediate Care - Attalla clinic ID (placeholder, get from MHT)
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
            elements = findwindows.find_elements(title_re='.*Tracking Board.*', backend='uia')
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

        # Try Experity window as fallback
        try:
            app = Application(backend='uia').connect(title_re='.*Experity.*', found_index=0, timeout=3)
            win = app.window(title_re='.*Experity.*', found_index=0)
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

        # Position on LEFT side
        log_width = 480
        log_height = 420  # Compact height
        screen_height = patient_log_root.winfo_screenheight()
        y_pos = screen_height - log_height - 120
        log_win.geometry(f"{log_width}x{log_height}+40+{y_pos}")

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
                                                      fg=text_secondary, wraplength=440, justify="left")
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
                # DEBUG: Print analytics to console
                print(f"[ANALYTICS] Patients: {analytics_cache['total_patients']}, Success: {analytics_cache['successful']}, Rate: {analytics_cache['success_rate']}")
        except Exception as e:
            print(f"[ANALYTICS ERROR] {e}")

    def update_analytics_tab():
        """Update the analytics tab display from cache."""
        if not patient_log_root or not analytics_labels:
            print(f"[ANALYTICS TAB] Cannot update - root: {patient_log_root is not None}, labels: {len(analytics_labels) if analytics_labels else 0}")
            return
        try:
            def _do_update():
                updated = 0
                for key, val in analytics_cache.items():
                    if key in analytics_labels:
                        analytics_labels[key].config(text=val)
                        updated += 1
                print(f"[ANALYTICS TAB] Updated {updated} labels")
            patient_log_root.after(0, _do_update)
        except Exception as e:
            print(f"[ANALYTICS TAB ERROR] {e}")

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
                # Start at y=220 to avoid covering Close button at y=185
                roomed_win.geometry(f'400x250+{screen_width - 440}+220')
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
                # Position below the Roomed Patients overlay (220 + 250 + 10 = 480)
                discharged_win.geometry(f'400x200+{screen_width - 440}+480')
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

        # Position in TOP LEFT corner
        data_win.geometry('450x400+40+40')
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

    # Close button coordinates on Demographics screen
    CLOSE_BTN_X = 2199
    CLOSE_BTN_Y = 1496

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
        pyautogui.click(cx, cy)
        patient_data = {'name': name, 'clicked': True}

        try:
            # Wait for Demographics window - proceeds instantly when it appears
            demo_win = wait_for_window('.*Demographics.*', timeout=5)
            if not demo_win:
                dismiss_popup_dialogs()
                demo_win = wait_for_window('.*Demographics.*', timeout=3)

            if not demo_win:
                control.add_log(f"Demographics window didn't open for {name}")
                pyautogui.click(CLOSE_BTN_X, CLOSE_BTN_Y)
                return patient_data

            control.set_step("Extracting...")

            # BATCH EXTRACT: Get all elements at once
            edits = demo_win.descendants(control_type='Edit')
            texts = demo_win.descendants(control_type='Text')
            radios = demo_win.descendants(control_type='RadioButton')

            for edit in edits[:30]:
                try:
                    txt = edit.window_text()
                    if not txt:
                        continue
                    rect = edit.rectangle()
                    if 330 < rect.top < 345 and rect.left < 600:
                        patient_data['first_name'] = txt
                    elif 368 < rect.top < 385 and rect.left < 600:
                        patient_data['last_name'] = txt
                    elif 565 < rect.top < 580 and rect.left < 600:
                        patient_data['dob'] = txt
                    elif 330 < rect.top < 350 and 1100 < rect.left < 1500:
                        patient_data['cell_phone'] = txt
                    elif 640 < rect.top < 660 and 1100 < rect.left < 1500:
                        patient_data['email'] = txt
                except:
                    pass

            for txt_elem in texts[:30]:
                try:
                    txt = txt_elem.window_text()
                    if txt and 'Patient Number:' in txt:
                        patient_data['mrn'] = txt.split(':')[-1].strip()
                        break
                except:
                    pass

            try:
                for rb in radios:
                    txt = rb.window_text()
                    if txt in ['Male', 'Female', 'Unknown']:
                        rect = rb.rectangle()
                        if rect.top < 700:
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

            skip_values = ['Race:', 'Race', 'Ethnicity:', 'Ethnicity', 'Language:', 'Language', 'Employer', 'Name:', 'Phone:', 'Ext:', 'Cultural Information']
            found_race = found_eth = found_lang = False
            for txt_elem in texts[:100]:
                try:
                    txt = txt_elem.window_text()
                    if not txt:
                        continue
                    if 'Race:' in txt or txt == 'Race':
                        found_race = True
                    elif found_race and txt.strip() and txt.strip() not in skip_values:
                        patient_data['race'] = txt.strip()
                        found_race = False
                    if 'Ethnicity:' in txt or txt == 'Ethnicity':
                        found_eth = True
                    elif found_eth and txt.strip() and txt.strip() not in skip_values:
                        patient_data['ethnicity'] = txt.strip()
                        found_eth = False
                    if 'Language:' in txt or txt == 'Language':
                        found_lang = True
                    elif found_lang and txt.strip() and txt.strip() not in skip_values:
                        patient_data['language'] = txt.strip()
                        found_lang = False
                except:
                    pass

            extracted = [f"{k}={v}" for k, v in patient_data.items() if k not in ['name', 'clicked']]
            control.add_log(f"Extracted: {', '.join(extracted)}")

            # Close demographics and WAIT for it to disappear (no fixed sleep)
            pyautogui.click(CLOSE_BTN_X, CLOSE_BTN_Y)
            wait_for_window_close('.*Demographics.*', timeout=3)

        except Exception as e:
            control.add_log(f"Demographics error: {str(e)[:30]}")
            pyautogui.click(CLOSE_BTN_X, CLOSE_BTN_Y)
            wait_for_window_close('.*Demographics.*', timeout=2)

        # === INSURANCE EXTRACTION - USE STORED DATA_ITEM (NO RE-SEARCH) ===
        control.set_step("Opening chart...")

        try:
            # Get chart icon directly from stored data_item reference
            chart_icon = None
            if data_item:
                try:
                    images_in_row = data_item.descendants(control_type='Image')
                    if images_in_row:
                        chart_icon = images_in_row[0]
                except:
                    control.add_log("Stored data_item stale, falling back")

            # Fallback: quick search if stored ref failed
            if not chart_icon:
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
                                break
                        except:
                            continue
                except:
                    pass

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

                # Click Demographics tab and wait for content to load
                pyautogui.click(2462, 472)
                # Wait for tab content by checking for text elements to appear
                deadline = time.time() + 3
                texts = []
                while time.time() < deadline:
                    try:
                        texts = chart_win.descendants(control_type='Text')
                        if len(texts) > 10:  # Demographics tab has many text elements
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

                    # Race/Ethnicity/Language from chart (if not already found in popup)
                    skip_values = ['Race:', 'Race', 'Ethnicity:', 'Ethnicity', 'Language:', 'Language', 'Employer', 'Name:', 'Phone:', 'Ext:', 'Cultural Information']
                    found_race = found_eth = found_lang = False
                    for t in texts[:150]:
                        try:
                            txt = t.window_text()
                            if not txt:
                                continue
                            if 'race' not in patient_data:
                                if 'Race:' in txt or txt == 'Race':
                                    found_race = True
                                elif found_race and txt.strip() and txt.strip() not in skip_values:
                                    patient_data['race'] = txt.strip()
                                    found_race = False
                            if 'ethnicity' not in patient_data:
                                if 'Ethnicity:' in txt or txt == 'Ethnicity':
                                    found_eth = True
                                elif found_eth and txt.strip() and txt.strip() not in skip_values:
                                    patient_data['ethnicity'] = txt.strip()
                                    found_eth = False
                            if 'language' not in patient_data:
                                if 'Language:' in txt or txt == 'Language':
                                    found_lang = True
                                elif found_lang and txt.strip() and txt.strip() not in skip_values:
                                    patient_data['language'] = txt.strip()
                                    found_lang = False
                        except:
                            pass
                except Exception as ins_err:
                    control.add_log(f"Insurance extraction error: {str(ins_err)[:30]}")

                # Close sidebar then chart, WAIT for chart to disappear
                pyautogui.click(2432, 185)
                time.sleep(0.1)  # tiny gap between clicks
                pyautogui.click(1782, 1203)
                wait_for_window_close(chart_title_re, timeout=3)
                control.add_log("Chart closed")

            else:
                control.add_log("Chart icon not found")

        except Exception as chart_err:
            control.add_log(f"Chart error: {str(chart_err)[:30]}")
            try:
                pyautogui.click(2432, 185)
                time.sleep(0.1)
                pyautogui.click(1782, 1203)
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

        # Start analytics
        try:
            analytics.start_patient_processing(name, '')
        except Exception as e:
            control.add_log(f"Analytics start err: {e}")

        # Click and get data
        data = click_qualified_patient(patient)

        if data:
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

        # Failed
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

            if row_rect:
                # Use the provided rect directly - much faster and more reliable
                cy = (row_rect.top + row_rect.bottom) // 2
                cx = 1069  # Name column x-coordinate for roomed patients
                control.add_log(f"Using provided rect for {name}, clicking at ({cx}, {cy})")
            else:
                # Fallback: Find the Roomed Patients group and search for patient
                control.add_log(f"No rect provided, searching for {name} in Roomed Patients...")
                roomed_group = tb_win.child_window(title_re='.*Roomed Patients.*', control_type='Group')
                roomed_data_items = roomed_group.descendants(control_type='DataItem')

                patient_name = name.upper()
                name_rect = None

                # Find the patient row
                for di in roomed_data_items:
                    try:
                        children = di.descendants()
                        found_name = False

                        for child in children:
                            child_text = child.window_text()
                            if child_text and patient_name in child_text.upper():
                                found_name = True
                                break

                        if found_name:
                            name_rect = di.rectangle()
                            control.add_log(f"Found {name} in Roomed section")
                            break

                    except Exception as row_err:
                        continue

                if not name_rect:
                    control.add_log(f"Could not find {name} in Roomed Patients section")
                    return None

                cy = (name_rect.top + name_rect.bottom) // 2
                cx = 1069  # Name column x-coordinate for roomed patients

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
                    pyautogui.click(CLOSE_BTN_X, CLOSE_BTN_Y)
                    return None

                control.add_log(f"Connected to Demographics for roomed patient")

                # Get all Edit fields
                edits = demo_win.descendants(control_type='Edit')

                for edit in edits[:30]:
                    try:
                        txt = edit.window_text()
                        if not txt:
                            continue
                        rect = edit.rectangle()

                        # First Name
                        if 330 < rect.top < 345 and rect.left < 600:
                            patient_data['first_name'] = txt
                            control.add_log(f"First Name: {txt}")
                        # Last Name
                        elif 368 < rect.top < 385 and rect.left < 600:
                            patient_data['last_name'] = txt
                            control.add_log(f"Last Name: {txt}")
                        # DOB
                        elif 565 < rect.top < 580 and rect.left < 600:
                            patient_data['dob'] = txt
                            control.add_log(f"DOB: {txt}")
                        # Cell Phone
                        elif 330 < rect.top < 350 and 1100 < rect.left < 1500:
                            patient_data['cell_phone'] = txt
                            control.add_log(f"Cell Phone: {txt}")
                        # Email
                        elif 640 < rect.top < 660 and 1100 < rect.left < 1500:
                            patient_data['email'] = txt
                            control.add_log(f"Email: {txt}")
                    except:
                        pass

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
                            if rect.top < 700:
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
                pyautogui.click(CLOSE_BTN_X, CLOSE_BTN_Y)
                wait_for_window_close(f'.*{re.escape(name_part)}.*', timeout=2)
                control.add_log("Closed demographics")

            except Exception as demo_err:
                control.add_log(f"Demographics error: {str(demo_err)[:30]}")
                pyautogui.click(CLOSE_BTN_X, CLOSE_BTN_Y)
                wait_for_window_close(f'.*{re.escape(name_part)}.*', timeout=2)

            # Click chart icon and WAIT for chart window (no fixed sleep)
            control.set_step("Opening chart...")
            try:
                chart_cx = 978
                chart_cy = cy
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

                # Click Demographics tab and wait for content
                pyautogui.click(2462, 472)
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

                # Extract Race
                for t in texts_with_pos:
                    if t['txt'] == 'Race:':
                        ry, rx = t['y'], t['x']
                        for v in texts_with_pos:
                            if abs(v['y'] - ry) < 20 and v['x'] > rx + 50:
                                patient_data['race'] = v['txt']
                                control.add_log(f"Race: {v['txt']}")
                                break
                        break

                # Extract Ethnicity
                for t in texts_with_pos:
                    if t['txt'] == 'Ethnicity:':
                        ey, ex = t['y'], t['x']
                        for v in texts_with_pos:
                            if abs(v['y'] - ey) < 20 and v['x'] > ex + 50:
                                patient_data['ethnicity'] = v['txt']
                                control.add_log(f"Ethnicity: {v['txt']}")
                                break
                        break

                # Extract Language
                for t in texts_with_pos:
                    if t['txt'] == 'Language:':
                        ly, lx = t['y'], t['x']
                        for v in texts_with_pos:
                            if abs(v['y'] - ly) < 20 and v['x'] > lx + 50:
                                patient_data['language'] = v['txt']
                                control.add_log(f"Language: {v['txt']}")
                                break
                        break

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
                pyautogui.click(2406, 184)
                time.sleep(0.1)
                pyautogui.click(1792, 1204)
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
    control.add_log("Starting monitoring loop")

    try:
        while not control.is_killed:
            cycle += 1
            control.add_log(f"--- Cycle {cycle} ---")
            # Start new cycle - clears display and starts fresh
            update_patient_log(f"=== Cycle {cycle} ===", new_cycle=cycle)

            # === CLEAR STALE OVERLAYS BEFORE RESCAN ===
            clear_patient_overlays()

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
                    elements = findwindows.find_elements(title_re='.*Tracking Board.*', backend='uia')
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
        clear_patient_overlays()  # Clear patient row overlays first
        for root in [wr_overlay_root, patient_log_root, roomed_overlay_root, discharged_overlay_root, patient_overlay_root, extracted_data_root]:
            if root:
                try:
                    root.after(0, root.quit)
                except:
                    pass
        time.sleep(0.5)  # Give overlays time to close


def _change_location(control):
    """Change clinic location to ATTALLA via menu navigation."""
    import pyautogui

    # Coordinates from tracker.py - exact screen positions
    CLINIC_MENU = {"x": 177, "y": 67}
    CURRENT_CLINIC = {"x": 223, "y": 203}
    ATTALLA = {"x": 399, "y": 246}

    try:
        # Step 1: Click Clinic menu
        control.set_step("Opening Clinic menu...")
        control.add_log(f"Step 1: Clicking Clinic menu at ({CLINIC_MENU['x']}, {CLINIC_MENU['y']})")

        show_target(CLINIC_MENU['x'], CLINIC_MENU['y'], "CLINIC MENU", color="#2196F3", width=120, height=35, duration=0.6)
        time.sleep(0.4)
        pyautogui.click(CLINIC_MENU['x'], CLINIC_MENU['y'])
        control.add_log("Clicked Clinic menu")
        time.sleep(0.5)

        # Step 2: Hover over Current Clinic to open submenu
        control.set_step("Selecting Current Clinic...")

        show_target(CURRENT_CLINIC['x'], CURRENT_CLINIC['y'], "CURRENT CLINIC", color="#FF9800", width=140, height=35, duration=0.6)
        time.sleep(0.4)
        pyautogui.moveTo(CURRENT_CLINIC['x'], CURRENT_CLINIC['y'], duration=0.2)
        control.add_log("Hovering over Current Clinic")
        time.sleep(0.5)

        # Step 3: Click ATTALLA
        control.set_step("Selecting ATTALLA...")

        show_target(ATTALLA['x'], ATTALLA['y'], "ATTALLA", color="#4CAF50", width=100, height=35, duration=0.6)
        time.sleep(0.4)
        pyautogui.click(ATTALLA['x'], ATTALLA['y'])
        control.add_log("Clicked ATTALLA - location changed!")
        time.sleep(0.8)

        control.add_log("Location change completed successfully")
        return True

    except Exception as e:
        control.add_log(f"Location change error: {str(e)}")
        return False


def _fast_login(control):
    """Fast login using pywinauto for Chromium-based Experity app."""
    from pywinauto import Application
    import pyautogui

    control.set_status("Logging In")
    control.set_step("Connecting to Experity...")
    control.add_log("Starting login with pywinauto")

    try:
        # ===== SCREEN 1: Experity Username =====
        control.add_log("Connecting to Experity window...")

        app = Application(backend='uia').connect(title_re='.*Experity.*', found_index=0, timeout=10)
        win = app.window(title_re='.*Experity.*', found_index=0)
        control.add_log(f"Connected to: {win.window_text()}")

        control.set_status("Logging In")
        control.set_step("Entering username...")

        # Find username field
        edits = win.descendants(control_type='Edit')
        control.add_log(f"Found {len(edits)} edit fields")

        if edits:
            username_edit = edits[0]

            # Show overlay on username field
            _show_element_overlay(username_edit, "USERNAME", color="#2196F3", duration=1.2)
            time.sleep(0.8)

            # Use clipboard paste for reliable input
            import pyautogui
            import pyperclip
            username_edit.click_input()
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            pyperclip.copy(STORED_USERNAME)
            pyautogui.hotkey('ctrl', 'v')
            control.add_log(f"Username entered: {STORED_USERNAME}")

        # Click Next button
        control.set_step("Clicking Next...")
        buttons = win.descendants(control_type='Button')
        for btn in buttons:
            if 'Next' in btn.window_text():
                # Show overlay on Next button
                _show_element_overlay(btn, "NEXT", color="#4CAF50", duration=1.0)
                time.sleep(0.6)

                btn.click()
                control.add_log("Clicked Next button")
                break

        # ===== Wait for Screen 2 (Okta) =====
        control.set_status("Authenticating")
        control.set_step("Waiting for sign in page...")
        time.sleep(5)

        # ===== SCREEN 2: Okta Password =====
        control.add_log("Connecting to Okta Sign In window...")

        app2 = Application(backend='uia').connect(title='Sign In', timeout=10)
        win2 = app2.window(title='Sign In')
        control.add_log(f"Connected to: {win2.window_text()}")

        # Find edit fields
        edits2 = win2.descendants(control_type='Edit')
        control.add_log(f"Found {len(edits2)} edit fields on Okta")

        # Check if username field needs to be filled
        if len(edits2) >= 1:
            username_edit2 = edits2[0]
            current_username = ""
            try:
                current_username = username_edit2.get_value()
            except:
                try:
                    current_username = username_edit2.window_text()
                except:
                    pass

            control.add_log(f"Username field value: '{current_username}'")

            # If username is empty or just placeholder, fill it
            if not current_username or current_username.strip() == "" or "username" in current_username.lower():
                control.set_step("Entering username...")
                _show_element_overlay(username_edit2, "USERNAME", color="#2196F3", duration=1.2)
                time.sleep(0.8)

                username_edit2.click_input()
                time.sleep(0.2)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyperclip.copy(STORED_USERNAME)
                pyautogui.hotkey('ctrl', 'v')
                control.add_log(f"Username entered: {STORED_USERNAME}")
                time.sleep(0.3)
            else:
                control.add_log("Username already filled, skipping")

        # Password field
        control.set_step("Entering password...")
        if len(edits2) >= 2:
            password_edit = edits2[1]

            # Show overlay on password field
            _show_element_overlay(password_edit, "PASSWORD", color="#FF9800", duration=1.2)
            time.sleep(0.8)

            # Use clipboard paste for reliable input
            import pyautogui
            import pyperclip
            password_edit.click_input()
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.1)
            pyperclip.copy(STORED_PASSWORD)
            pyautogui.hotkey('ctrl', 'v')
            control.add_log("Password entered")

        # Click Sign In button
        control.set_step("Clicking Sign In...")
        buttons2 = win2.descendants(control_type='Button')
        for btn in buttons2:
            name = btn.window_text()
            if name == 'Sign In':
                # Show overlay on Sign In button
                _show_element_overlay(btn, "SIGN IN", color="#4CAF50", duration=1.0)
                time.sleep(0.6)

                btn.click()
                control.add_log("Clicked Sign In button")
                break

        # ===== SCREEN 3: OTP Verification (DEMO MODE) =====
        control.set_status("Verification")
        control.set_step("Waiting for SMS screen...")
        control.add_log("Waiting 5 seconds for OTP screen to load...")
        time.sleep(5)

        control.set_step("Loading verification...")
        control.add_log("Analyzing OTP verification screen...")

        # Take screenshot for analysis
        import pyautogui
        otp_screenshot = pyautogui.screenshot()
        otp_debug_path = SCRIPT_DIR / "output" / "debug" / "otp_screen.png"
        otp_debug_path.parent.mkdir(parents=True, exist_ok=True)
        otp_screenshot.save(otp_debug_path)
        control.add_log(f"OTP screen captured: {otp_debug_path}")

        # Reconnect to the window - try multiple title patterns
        win3 = None
        app3 = None

        # Try different window title patterns
        title_patterns = ['Sign In', '.*Sign.*', '.*Okta.*', '.*SMS.*', '.*Verify.*', '.*Authentication.*']

        for pattern in title_patterns:
            try:
                control.add_log(f"Trying window pattern: {pattern}")
                if '.*' in pattern:
                    app3 = Application(backend='uia').connect(title_re=pattern, timeout=3)
                    win3 = app3.window(title_re=pattern)
                else:
                    app3 = Application(backend='uia').connect(title=pattern, timeout=3)
                    win3 = app3.window(title=pattern)
                control.add_log(f"Connected to: {win3.window_text()}")
                break
            except Exception as e:
                control.add_log(f"Pattern '{pattern}' failed: {str(e)[:30]}")
                continue

        if not win3:
            control.set_status("OTP Error", "#dc3545")
            control.set_step("Could not find OTP window!")
            control.add_log("ERROR: Could not connect to OTP window with any pattern")
            time.sleep(3)
            return

        try:
            control.add_log("Connected to OTP screen")

            # Log all buttons found for analysis
            buttons3 = win3.descendants(control_type='Button')
            control.add_log(f"Found {len(buttons3)} buttons on OTP screen:")
            for btn in buttons3:
                name = btn.window_text()
                if name.strip():
                    control.add_log(f"  - Button: '{name}'")

            # Log all edit fields found
            edits3 = win3.descendants(control_type='Edit')
            control.add_log(f"Found {len(edits3)} edit fields on OTP screen")

            # ===== Click Send code to request OTP =====
            control.set_step("Found SMS authentication screen")
            send_code_elem = None

            # Send code is a Hyperlink, not a Button
            try:
                send_code_elem = win3.child_window(control_type='Hyperlink', title='Send code')
                control.add_log(f"Found Send code hyperlink")
            except:
                control.add_log("Send code hyperlink not found with exact match, trying search...")
                links = win3.descendants(control_type='Hyperlink')
                for link in links:
                    name = link.window_text()
                    if 'send' in name.lower():
                        send_code_elem = link
                        control.add_log(f"Found Send code as Hyperlink: '{name}'")
                        break

            if send_code_elem:
                _show_element_overlay(send_code_elem, "SEND CODE", color="#FF5722", duration=1.5)
                time.sleep(1.0)
                control.set_step("Requesting SMS code...")
                send_code_elem.click_input()
                control.add_log("Clicked Send code - SMS requested")
                time.sleep(2.0)
            else:
                control.add_log("Send code not found - continuing anyway")
                time.sleep(1)

            # ===== Prompt user for OTP code =====
            control.set_status("Enter Code")
            control.set_step("Check your phone for SMS")
            control.add_log("=== WAITING FOR OTP CODE INPUT ===")

            # Prompt with 120 second timeout
            otp_code = control.prompt_input("Enter OTP Code:", timeout=120)

            if otp_code:
                control.add_log(f"OTP received: {otp_code}")
                control.set_status("Verifying")
                control.set_step(f"Entering code: {otp_code}")

                # Find OTP input field using exact match
                try:
                    otp_edit = win3.child_window(control_type='Edit', title_re='.*Enter Code.*')
                    control.add_log("Found OTP input field")
                except:
                    otp_edit = edits3[0] if edits3 else None
                    control.add_log("Using first edit field for OTP")

                if otp_edit:
                    # Show overlay on OTP input field
                    _show_element_overlay(otp_edit, "ENTERING CODE", color="#2196F3", duration=1.5)
                    time.sleep(1.0)

                    # Use clipboard paste for reliable input
                    import pyperclip
                    otp_edit.click_input()
                    time.sleep(0.2)
                    pyautogui.hotkey('ctrl', 'a')
                    time.sleep(0.1)
                    pyperclip.copy(otp_code)
                    pyautogui.hotkey('ctrl', 'v')
                    control.add_log(f"OTP '{otp_code}' entered into field")
                    time.sleep(0.5)
                else:
                    control.add_log("ERROR: OTP input field not found!")

                # ===== Click Verify button =====
                try:
                    verify_btn = win3.child_window(control_type='Button', title='Verify')
                    control.add_log("Found Verify button")
                except:
                    verify_btn = None
                    for btn in buttons3:
                        if btn.window_text() == 'Verify':
                            verify_btn = btn
                            break

                if verify_btn:
                    try:
                        _show_element_overlay(verify_btn, "VERIFY", color="#4CAF50", duration=1.5)
                        time.sleep(1.0)
                    except:
                        pass
                    # Click FIRST, then update UI - so click always happens
                    verify_rect = verify_btn.rectangle()
                    verify_x = (verify_rect.left + verify_rect.right) // 2
                    verify_y = (verify_rect.top + verify_rect.bottom) // 2
                    pyautogui.click(verify_x, verify_y)
                    print(f"Clicked Verify button at ({verify_x}, {verify_y})")
                    try:
                        control.set_step("Verifying code...")
                        control.add_log(f"Clicked Verify button at ({verify_x}, {verify_y})")
                    except:
                        pass
                    time.sleep(3.0)
                else:
                    try:
                        control.add_log("Verify button not found")
                    except:
                        pass
                    print("Verify button not found")

                try:
                    control.set_status("Loading")
                    control.set_step("Waiting for EMR to load...")
                    control.add_log("Login successful - waiting for EMR to fully load")
                except:
                    print("Login successful - waiting for EMR to fully load")

                # Wait for Experity EMR to fully load after verification (90 seconds with countdown)
                try:
                    control.add_log("Waiting 90 seconds for EMR to fully load...")
                except:
                    pass
                print("Waiting 90 seconds for EMR to fully load...")
                for remaining in range(90, 0, -1):
                    try:
                        control.set_step(f"EMR loading... {remaining}s remaining")
                    except:
                        pass
                    time.sleep(1)

                # ===== CHANGE LOCATION TO ATTALLA =====
                try:
                    control.set_status("Location")
                    control.set_step("Changing to ATTALLA...")
                    control.add_log("Starting location change process")
                except:
                    pass
                print("Starting location change process")

                _change_location(control)

                try:
                    control.set_status("Success")
                    control.set_step("Ready at ATTALLA!")
                    control.add_log("Location changed - starting monitoring")
                except:
                    pass
                print("Location changed - starting monitoring")

                # Start monitoring the Waiting Room
                time.sleep(1.0)
                _start_monitoring(control)
            else:
                try:
                    control.add_log("OTP input cancelled or timed out")
                    control.set_status("Cancelled")
                    control.set_step("No code entered")
                except:
                    pass
                print("OTP input cancelled or timed out")

        except Exception as otp_err:
            print(f"OTP screen error: {str(otp_err)}")
            try:
                control.add_log(f"OTP screen error: {str(otp_err)}")
                control.set_status("Error")
                control.set_step(f"OTP error: {str(otp_err)[:30]}")
            except:
                pass

    except Exception as e:
        print(f"Login error: {str(e)}")
        try:
            control.set_status("Error")
            control.set_step(f"Error: {str(e)[:40]}")
            control.add_log(f"Login error: {str(e)}")
            import traceback
            control.add_log(f"Traceback: {traceback.format_exc()}")
        except:
            import traceback
            print(f"Traceback: {traceback.format_exc()}")


if __name__ == "__main__":
    try:
        run_silent()
    except Exception as e:
        # Last resort error logging
        import traceback
        error_file = SCRIPT_DIR / "output" / "debug" / "crash_log.txt"
        error_file.parent.mkdir(parents=True, exist_ok=True)
        with open(error_file, "w") as f:
            f.write(f"CRASH LOG\n")
            f.write(f"Error: {str(e)}\n")
            f.write(f"Traceback:\n{traceback.format_exc()}\n")

        # Show error in a simple message box
        import tkinter.messagebox as mb
        mb.showerror("MHT Agentic Crash", f"Script crashed!\n\nError: {str(e)}\n\nCheck: {error_file}")