"""
MHT Outbound Worker

Monitors for completed outbound events (status=100) and processes them
by entering assessment results into Experity.
"""

import sqlite3
import json
import time
import threading
import logging
import ctypes
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List, Callable, TYPE_CHECKING

# Enable DPI awareness so pyautogui coordinates match pywinauto coordinates
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    pass

from pywinauto import Application, findwindows, Desktop
import pyautogui
from mhtagentic.desktop.session_guard import session_find_elements, session_connect

if TYPE_CHECKING:
    from mhtagentic.desktop.control_overlay import ControlOverlay

logger = logging.getLogger(__name__)


class OutboundWorker:
    """
    Worker that polls for completed outbound events and processes them.

    Runs in a background thread, checking every poll_interval seconds
    for outbound events with status=100 (ready to process).
    """

    # Legacy status codes (kept for backwards compatibility)
    STATUS_READY = 100      # Ready to process (from MHT)
    STATUS_PROCESSING = 50  # Currently being processed
    STATUS_DONE = 200       # Successfully entered into Experity
    STATUS_ERROR = -100     # Failed to process

    # Granular step statuses for multi-bot coordination
    STEP_FRESH = 10           # Not yet claimed
    STEP_CLAIMED = 11         # Claimed by a bot
    STEP_LOCATION_CHECK = 15  # Checking/switching location
    STEP_FIND_PATIENT = 20    # Finding patient in Experity
    STEP_OPEN_CHART = 30      # Opening chart / procedures
    STEP_SELECT_ASSESSMENT = 40  # Selecting assessment type
    STEP_FILL_FORM = 50       # Filling radio buttons, scores
    STEP_ENTER_NOTES = 60     # Entering clinical notes
    STEP_CLOSE_FORM = 70      # Closing/saving form
    STEP_CLEANUP = 80         # Delete proc, close chart
    STEP_DONE = 100           # Successfully completed
    # Negative = error at that step (e.g. -20 = failed finding patient)

    def __init__(self, db_path: str, poll_interval: float = 5.0, overlay: Optional['ControlOverlay'] = None):
        """
        Initialize the outbound worker.

        Args:
            db_path: Path to SQLite database
            poll_interval: Seconds between polls (default 5)
            overlay: Optional ControlOverlay for UI status updates
        """
        self.db_path = Path(db_path)
        self.poll_interval = poll_interval
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._callback: Optional[Callable] = None
        self._processing = False  # Flag to indicate if currently processing
        self._overlay = overlay  # UI overlay for status updates
        self.demo_mode = False  # Skip hide/show overlays in demo mode
        self._demo_overlays = []  # Demo overlay references for re-lifting
        import os
        self._current_location = os.environ.get("BOT_LOCATION", "ATTALLA").upper()
        self._bot_name = os.environ.get("FORCE_BOT_USER", os.environ.get("USERNAME", "unknown"))

    def set_callback(self, callback: Callable[[Dict], None]):
        """Set callback function called when an event is processed."""
        self._callback = callback

    def set_overlay(self, overlay: 'ControlOverlay'):
        """Set the UI overlay for status updates."""
        self._overlay = overlay

    def is_processing(self) -> bool:
        """Check if currently processing an outbound event."""
        return self._processing

    def _update_overlay_status(self, text: str, color: str = "#4ECDC4"):
        """Update overlay status text (thread-safe)."""
        if self._overlay:
            try:
                self._overlay.set_status(text, color)
            except Exception as e:
                logger.warning(f"Overlay status update failed: {e}")

    def _update_overlay_step(self, step_num: int, total: int, description: str):
        """Update overlay step progress (thread-safe)."""
        if self._overlay:
            try:
                self._overlay.set_step(f"Step {step_num}/{total}: {description}")
            except Exception as e:
                logger.warning(f"Overlay step update failed: {e}")

    def _hide_all_overlays(self):
        """Hide patient overlays during outbound processing but keep control overlay visible."""
        if self.demo_mode:
            return  # No overlays to hide in demo mode
        try:
            from mhtagentic.desktop.control_overlay import hide_all_overlays, _all_overlay_windows
            # Only hide non-control overlays (patient highlights, etc.)
            # Keep the control overlay (bottom-right) visible
            for win in list(_all_overlay_windows):
                try:
                    # Skip the main control overlay (it has specific class name)
                    if hasattr(win, 'title') and 'MHT' in str(win.title()):
                        continue
                    win.withdraw()
                except:
                    pass
            logger.info("Patient overlays hidden (control overlay kept visible)")
        except Exception as e:
            logger.warning(f"Hide overlays failed: {e}")

    def set_demo_overlays(self, overlays: list):
        """Set demo overlay references for re-lifting after outbound processing."""
        self._demo_overlays = overlays

    def _show_all_overlays(self):
        """Show ALL overlays after outbound processing."""
        if self.demo_mode:
            # Re-lift demo overlays that may have gone behind Experity windows
            for overlay in self._demo_overlays:
                try:
                    if overlay and hasattr(overlay, 'panel') and overlay.panel and overlay._running:
                        overlay.root.after(0, lambda o=overlay: self._relift_demo_overlay(o))
                except Exception as e:
                    logger.warning(f"Failed to re-lift demo overlay: {e}")
            return
        try:
            from mhtagentic.desktop.control_overlay import show_all_overlays
            show_all_overlays()

            logger.info("All overlays restored after outbound processing")
        except Exception as e:
            logger.warning(f"Show all overlays failed: {e}")

    @staticmethod
    def _relift_demo_overlay(overlay):
        """Re-lift a demo overlay to topmost."""
        try:
            if overlay.panel and overlay.panel.winfo_exists():
                overlay.panel.deiconify()
                overlay.panel.lift()
                overlay.panel.attributes("-topmost", True)
        except Exception:
            pass

    # --- Background popup monitor (birthday modals, error dialogs) ---
    _popup_monitor_running = False
    _popup_monitor_thread: Optional[threading.Thread] = None

    def _start_popup_monitor(self):
        """Start a background thread that auto-dismisses birthday/error popups during processing."""
        if self._popup_monitor_running:
            return
        self._popup_monitor_running = True
        self._popup_monitor_thread = threading.Thread(target=self._popup_monitor_loop, daemon=True)
        self._popup_monitor_thread.start()
        logger.info("Popup monitor started")

    def _stop_popup_monitor(self):
        """Stop the background popup monitor."""
        self._popup_monitor_running = False
        self._popup_monitor_thread = None
        logger.info("Popup monitor stopped")

    def _popup_monitor_loop(self):
        """Background loop that checks for and dismisses birthday/error popups every 0.8s."""
        _popup_titles = ('Birthday', 'Application Error', 'Error', 'Experity', 'PROD')
        _popup_buttons = ('OK', 'Ok', 'Close', 'Yes', 'Continue', 'Accept')
        while self._popup_monitor_running and self._processing:
            try:
                desktop = Desktop(backend='uia')
                for w in desktop.windows():
                    try:
                        title = w.window_text()
                        if not title:
                            continue
                        if any(kw in title for kw in _popup_titles) or title == 'Error':
                            # Skip main Tracking Board / chart windows (only dismiss popups)
                            if 'Tracking Board' in title and 'Birthday' not in title:
                                continue
                            buttons = w.descendants(control_type='Button')
                            for btn in buttons:
                                try:
                                    bt = btn.window_text()
                                    if bt in _popup_buttons:
                                        btn.click_input()
                                        logger.info(f"[POPUP MONITOR] Auto-dismissed '{title[:40]}' via [{bt}]")
                                        time.sleep(0.3)
                                        break
                                except:
                                    continue
                    except:
                        continue
            except:
                pass
            time.sleep(0.8)

    def _read_current_location(self) -> str:
        """Read the current clinic location from the Experity status bar.

        Inspects the 'BarStaticItemLink2CurrentClinic' element in the bottom-left
        status bar of the Tracking Board window.

        Returns:
            Current location name (e.g. 'ANNISTON', 'ATTALLA') or empty string if unreadable.
        """
        from pywinauto import Application, findwindows
        try:
            from mhtagentic.desktop.session_guard import session_find_elements as _sfind
        except ImportError:
            _sfind = findwindows.find_elements

        try:
            elements = _sfind(title_re='.*Tracking Board.*', backend='uia')
            target_handle = None
            for elem in elements:
                if 'Setup' not in elem.name:
                    target_handle = elem.handle
                    break
            if not target_handle:
                return ''

            app = Application(backend='uia').connect(handle=target_handle, timeout=3)
            win = app.window(handle=target_handle)

            # Try auto_id first (most reliable)
            try:
                clinic_btn = win.child_window(auto_id='BarStaticItemLink2CurrentClinic')
                location = (clinic_btn.element_info.name or '').strip()
                if location:
                    logger.info(f"Status bar shows current location: {location}")
                    return location.upper()
            except Exception:
                pass

            # Fallback: search status bar descendants for a TextBlock near CurrentClinic
            try:
                status_bar = win.child_window(class_name='RibbonStatusBarLeftPartControl')
                for child in status_bar.descendants(control_type='Text'):
                    name = (child.element_info.name or '').strip()
                    if name and name.upper() in ('ATTALLA', 'ANNISTON'):
                        logger.info(f"Status bar text shows location: {name}")
                        return name.upper()
            except Exception:
                pass

            return ''
        except Exception as e:
            logger.warning(f"Could not read current location from status bar: {e}")
            return ''

    def _switch_location(self, target_location: str):
        """Switch Experity clinic location before processing a patient from a different location.

        Navigates: Clinic menu → Current Clinic → target_location.
        Uses pywinauto element detection by auto_id/name for reliable clicking.
        """
        import time as _time
        from pywinauto import Application, findwindows
        try:
            from mhtagentic.desktop.session_guard import session_find_elements as _sfind
        except ImportError:
            _sfind = findwindows.find_elements

        try:
            elements = _sfind(title_re='.*Tracking Board.*', backend='uia')
            target_handle = None
            for elem in elements:
                if 'Setup' not in elem.name:
                    target_handle = elem.handle
                    break

            if not target_handle:
                logger.error("Tracking Board not found for location switch")
                return False

            app = Application(backend='uia').connect(handle=target_handle, timeout=3)
            win = app.window(handle=target_handle)

            # Step 1: Click Clinic menu (auto_id='Clinic', type=MenuItem)
            logger.info("Location switch: clicking Clinic menu...")
            try:
                clinic_menu = win.child_window(auto_id='Clinic', control_type='MenuItem')
                clinic_menu.click_input()
                logger.info("Clicked Clinic menu via auto_id")
            except Exception as e:
                logger.warning(f"Clinic menu auto_id lookup failed: {e}, trying name search...")
                # Fallback: search descendants for the exact 'Clinic' MenuItem
                clicked = False
                for elem in win.descendants(control_type='MenuItem'):
                    try:
                        if elem.element_info.name == 'Clinic':
                            elem.click_input()
                            clicked = True
                            break
                    except:
                        continue
                if not clicked:
                    logger.error("Could not find Clinic menu item")
                    return False

            _time.sleep(0.4)

            # Step 2: Hover "Current Clinic" to open submenu (DO NOT click)
            logger.info("Location switch: hovering Current Clinic...")
            current_found = False
            # Poll for Current Clinic menu item to appear
            for _ in range(20):
                for elem in win.descendants():
                    try:
                        name = (elem.element_info.name or '')
                        if 'Current Clinic' in name or ('current' in name.lower() and 'clinic' in name.lower()):
                            rect = elem.rectangle()
                            cx = (rect.left + rect.right) // 2
                            cy = (rect.top + rect.bottom) // 2
                            pyautogui.moveTo(cx, cy, duration=0.15)
                            current_found = True
                            logger.info(f"Hovering over '{name}' at ({cx},{cy})")
                            break
                    except:
                        continue
                if current_found:
                    break
                _time.sleep(0.1)

            if not current_found:
                logger.warning("Current Clinic not found in descendants, trying child_window...")
                try:
                    current_item = win.child_window(title_re='.*Current Clinic.*')
                    rect = current_item.rectangle()
                    pyautogui.moveTo((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2, duration=0.15)
                    current_found = True
                except:
                    logger.error("Could not find Current Clinic submenu")

            _time.sleep(0.4)

            # Step 3: Click target location in the submenu
            logger.info(f"Location switch: clicking {target_location}...")
            location_found = False
            target_lower = target_location.lower()
            for elem in win.descendants():
                try:
                    name = (elem.element_info.name or '')
                    if target_lower in name.lower() and elem.element_info.control_type in ('MenuItem', 'ListItem', 'Text', 'Button', 'Hyperlink', 'Custom'):
                        rect = elem.rectangle()
                        # Only click elements that are in the menu area (small, not full-width table headers)
                        el_width = rect.right - rect.left
                        if el_width < 300:
                            cx = (rect.left + rect.right) // 2
                            cy = (rect.top + rect.bottom) // 2
                            pyautogui.click(cx, cy)
                            location_found = True
                            logger.info(f"Clicked '{name}' (type={elem.element_info.control_type}) at ({cx},{cy})")
                            break
                except:
                    continue

            if not location_found:
                logger.error(f"Could not find {target_location} in submenu")
                pyautogui.press('escape')
                _time.sleep(0.3)
                pyautogui.press('escape')
                return False

            logger.info(f"Location switched to {target_location}")
            # Poll for Tracking Board to reload after location change
            for _ in range(20):
                try:
                    elems = _sfind(title_re='.*Tracking Board.*', backend='uia')
                    if any('Setup' not in e.name for e in elems):
                        logger.info("Tracking Board available after location switch")
                        break
                except:
                    pass
                _time.sleep(0.5)
            return True

        except Exception as e:
            logger.error(f"Location switch error: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def start(self):
        """Start the worker thread."""
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()
        logger.info("Outbound worker started")

    def stop(self):
        """Stop the worker thread."""
        self._running = False
        if self._thread:
            self._thread.join(timeout=10)
        logger.info("Outbound worker stopped")

    def is_running(self) -> bool:
        """Check if worker is running."""
        return self._running

    def process_pending(self) -> int:
        """
        Process all pending outbound events immediately (on-demand).
        Uses atomic claiming to prevent multi-bot conflicts.
        Returns number of events processed.
        """
        processed = 0
        try:
            while True:
                event = self._claim_next_outbound_event()
                if not event:
                    break
                self._process_event(event)
                processed += 1
        except Exception as e:
            logger.error(f"Error processing pending outbound: {e}")
        return processed

    def _get_connection(self) -> sqlite3.Connection:
        """Get database connection."""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _get_ready_outbound_events(self) -> List[Dict]:
        """Get outbound events ready to process (status=100).
        Legacy method — use _claim_next_outbound_event() for multi-bot safety.
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM common_event
                WHERE direction = 'O' AND status = ?
                ORDER BY id ASC
            """, (self.STATUS_READY,))
            return [dict(row) for row in cursor.fetchall()]
        finally:
            conn.close()

    def _claim_next_outbound_event(self) -> Optional[Dict]:
        """Atomically claim the next ready outbound event.

        Uses BEGIN EXCLUSIVE to prevent two outbound bots from grabbing
        the same event. Sets status to STEP_CLAIMED (11).

        Returns:
            The claimed event dict, or None if no events available.
        """
        conn = self._get_connection()
        try:
            conn.execute("BEGIN EXCLUSIVE")
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM common_event
                WHERE direction = 'O' AND status = ?
                ORDER BY id ASC LIMIT 1
            """, (self.STATUS_READY,))
            row = cursor.fetchone()
            if not row:
                conn.commit()
                return None

            event = dict(row)
            now = datetime.now().isoformat()
            conn.execute("""
                UPDATE common_event SET status = ?, updated_at = ? WHERE id = ?
            """, (self.STEP_CLAIMED, now, event['id']))
            conn.commit()
            logger.info(f"Claimed outbound event {event['id']} (status → {self.STEP_CLAIMED})")
            return event
        except Exception as e:
            try:
                conn.rollback()
            except Exception:
                pass
            logger.error(f"Failed to claim outbound event: {e}")
            return None
        finally:
            conn.close()

    def _update_event_status(self, event_id: int, status: int):
        """Update event status."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE common_event
                SET status = ?, updated_at = ?
                WHERE id = ?
            """, (status, datetime.now().isoformat(), event_id))
            conn.commit()
        finally:
            conn.close()

    def _run_loop(self):
        """Main worker loop — uses atomic claiming for multi-bot safety."""
        while self._running:
            try:
                event = self._claim_next_outbound_event()
                if event:
                    self._process_event(event)
                    continue  # Check for more immediately
            except Exception as e:
                logger.error(f"Error in outbound worker: {e}")

            # Wait for next poll
            for _ in range(int(self.poll_interval * 10)):
                if not self._running:
                    break
                time.sleep(0.1)

    def _process_event(self, event: Dict):
        """Process a single outbound event."""
        event_id = event['id']
        logger.info(f"Processing outbound event {event_id}")

        self._processing = True
        self._update_event_status(event_id, self.STATUS_PROCESSING)

        # Start background popup monitor (birthday modals, error dialogs)
        self._start_popup_monitor()

        # Hide ALL overlays during outbound processing to avoid click interference
        self._hide_all_overlays()

        try:
            # Parse the event data
            raw_data = json.loads(event['raw_data'])
            data = raw_data.get('data', raw_data)

            # Determine patient's location from event data
            event_location = data.get('clinic_location', '') or ''
            if not event_location:
                # Try to get from clinic object
                clinic = data.get('clinic', {})
                event_location = clinic.get('clinic_location', 'ATTALLA')
                # Strip prefix like "SOUTHERN IMMEDIATE CARE - " if present
                if ' - ' in event_location:
                    event_location = event_location.split(' - ')[-1]
            event_location = event_location.upper() if event_location else 'ATTALLA'

            # Read ACTUAL current location from the Experity status bar
            actual_location = self._read_current_location()
            if actual_location:
                self._current_location = actual_location
            logger.info(f"Patient location: {event_location}, Current UI location: {self._current_location}")

            # Only switch if the patient is at a different location
            if event_location != self._current_location:
                logger.info(f"Switching location: {self._current_location} → {event_location}")
                self._switch_location(event_location)
                self._current_location = event_location

            patient = data.get('patient', {})
            last_name = patient.get('patient_last_name', '') or ''
            first_name = patient.get('patient_first_name', '') or ''
            patient_name = f"{last_name}, {first_name}"

            if not last_name.strip() and not first_name.strip():
                logger.error(f"Event {event_id} has no patient name - skipping")
                self._update_event_status(event_id, self.STATUS_ERROR)
                return

            assessments = data.get('assessment', [])
            if not isinstance(assessments, list):
                assessments = [assessments]

            logger.info(f"Processing patient: {patient_name}")
            logger.info("Overlay hidden for outbound processing")

            # Run the full outbound flow
            success = self._run_outbound_flow(patient_name, assessments)

            if success:
                self._update_event_status(event_id, self.STATUS_DONE)
                logger.info(f"Successfully processed event {event_id}")
                if self._callback:
                    self._callback({'event_id': event_id, 'patient': patient_name, 'success': True})
            else:
                self._update_event_status(event_id, self.STATUS_ERROR)
                logger.error(f"Failed to process event {event_id}")

        except Exception as e:
            logger.error(f"Error processing event {event_id}: {e}")
            self._update_event_status(event_id, self.STATUS_ERROR)
        finally:
            self._stop_popup_monitor()
            self._processing = False
            # Show ALL overlays again after outbound processing
            self._show_all_overlays()

    def _run_outbound_flow(self, patient_name: str, assessments: List[Dict]) -> bool:
        """
        Run the full outbound flow for a patient.

        Steps 1-22 of the assessment entry flow.
        """
        try:
            patient_upper = patient_name.upper()
            name_part = patient_name.split(',')[0].strip()

            # Get assessment data
            assess = assessments[0]
            scores = [int(float(item.get('assessment_item_score', 0)))
                     for item in assess.get('assessment_items', [])]
            total_score = int(assess.get('total_score_value', sum(scores)))
            assess_name = assess.get('assessment_name', '')

            keywords = ['phq-9', 'phq9', 'patient health'] if 'phq' in assess_name.lower() else ['gad-7', 'gad7', 'general anxiety']

            logger.info(f"Assessment: {assess_name}, Scores: {scores}, Total: {total_score}")

            # ===== STEP 1: Connect to Tracking Board =====
            logger.info("[Step 1] Connecting to Tracking Board...")
            self._update_overlay_step(1, 21, "Connecting to Tracking Board...")
            elements = session_find_elements(title_re='.*Tracking Board.*', backend='uia')
            target_handle = None
            for elem in elements:
                if 'Setup' not in elem.name:
                    target_handle = elem.handle
                    break

            if not target_handle:
                logger.error("Tracking Board not found")
                return False

            app = Application(backend='uia').connect(handle=target_handle, timeout=3)
            tb_win = app.window(handle=target_handle)

            # ===== STEP 1b: Refresh Tracking Board before searching =====
            logger.info("[Step 1b] Refreshing Tracking Board (F5)...")
            self._update_overlay_step(1, 21, "Refreshing Tracking Board...")
            try:
                tb_win.set_focus()
                import time as _time
                _time.sleep(0.3)
                pyautogui.press('f5')
                _time.sleep(2)  # Wait for tracking board data to reload
                # Re-connect in case the window handle changed after refresh
                elements = session_find_elements(title_re='.*Tracking Board.*', backend='uia')
                for elem in elements:
                    if 'Setup' not in elem.name:
                        target_handle = elem.handle
                        break
                app = Application(backend='uia').connect(handle=target_handle, timeout=3)
                tb_win = app.window(handle=target_handle)
                logger.info("Tracking Board refreshed successfully")
            except Exception as e:
                logger.warning(f"Tracking Board refresh failed (continuing anyway): {e}")

            # ===== STEP 2: Find patient in Waiting Room or Roomed Patients =====
            logger.info("[Step 2] Finding patient...")
            self._update_overlay_step(2, 21, f"Finding {patient_name}...")
            patient_row = None

            # Search in both Waiting Room and Roomed Patients
            for section_name in ['.*Waiting Room.*', '.*Roomed Patients.*']:
                try:
                    section_group = tb_win.child_window(title_re=section_name, control_type='Group')
                    section_items = section_group.descendants(control_type='DataItem')

                    for di in section_items:
                        try:
                            di_text = di.window_text()
                            if patient_upper in di_text.upper():
                                patient_row = di
                                logger.info(f"Found patient in {section_name}")
                                break
                            for child in di.descendants():
                                if patient_upper in child.window_text().upper():
                                    patient_row = di
                                    logger.info(f"Found patient in {section_name}")
                                    break
                            if patient_row:
                                break
                        except:
                            continue
                    if patient_row:
                        break
                except:
                    continue

            if not patient_row:
                logger.error(f"Patient {patient_name} not found in Waiting Room or Roomed Patients")
                return False

            # ===== STEP 3: Click chart icon =====
            logger.info("[Step 3] Clicking chart icon...")
            self._update_overlay_step(3, 21, "Clicking chart icon...")
            images = patient_row.descendants(control_type='Image')
            if not images:
                logger.error("No chart icon found in patient row")
                return False
            chart_icon = images[0]
            rect = chart_icon.rectangle()
            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

            # ===== STEP 4: Wait for chart window =====
            # Dismiss any modal popups (birthday, alerts) that block the chart from appearing
            logger.info(f"[Step 4] Waiting for chart window (name_part='{name_part}')...")
            self._update_overlay_step(4, 21, "Loading chart...")
            import re as _re
            chart_title_re = f'.*{_re.escape(name_part)}.*'
            chart_win = None
            for _poll in range(20):  # Poll for chart window
                # Check for blocking popups (birthday modal, errors, etc.) each poll
                try:
                    from launcher import _check_tracking_board_popups
                    _check_tracking_board_popups()
                except ImportError:
                    try:
                        # Inline popup dismissal — find any OK/Close button on Tracking Board
                        _tb_app = Application(backend='uia').connect(title='Tracking Board', timeout=0.5)
                        _tb_win = _tb_app.window(title='Tracking Board')
                        for _btn in _tb_win.descendants(control_type='Button'):
                            try:
                                _bt = _btn.window_text()
                                if _bt in ('OK', 'Ok', 'Close', 'Yes'):
                                    _br = _btn.rectangle()
                                    _bw = _br.right - _br.left
                                    _bh = _br.bottom - _br.top
                                    if 30 <= _bw <= 300 and 15 <= _bh <= 60:
                                        _btn.click_input()
                                        logger.info(f"[Step 4] Dismissed popup via [{_bt}]")
                                        time.sleep(0.3)
                                        break
                            except:
                                continue
                    except:
                        pass
                try:
                    chart_app, chart_win = session_connect(title_re=chart_title_re, timeout=1)
                    break
                except:
                    pass
            if not chart_win:
                logger.error(f"Chart window not found for pattern '{chart_title_re}'")
                return False
            logger.info(f"Chart window found: {chart_win.window_text()[:60]}")

            # ===== STEP 5: Navigate to Procedures/Supplies =====
            logger.info("[Step 5] Clicking Procedures/Supplies...")
            self._update_overlay_step(5, 21, "Opening Procedures/Supplies...")
            for _ in range(30):  # Poll for element
                found = False
                try:
                    tabs = chart_win.descendants(control_type='TabItem')
                    for elem in tabs:
                        try:
                            if 'procedures' in (elem.element_info.name or '').lower():
                                rect = elem.rectangle()
                                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                found = True
                                break
                        except:
                            pass
                except:
                    pass
                if found:
                    break
                time.sleep(0.1)

            # ===== STEP 6: Navigate to SIC COMMON ORDERS =====
            logger.info("[Step 6] Clicking SIC COMMON ORDERS...")
            self._update_overlay_step(6, 21, "Opening SIC Common Orders...")
            time.sleep(0.1)  # Brief settle after Procedures tab

            # Log ALL element names containing "sic" to a debug file so we know what's there
            debug_log = Path(r"C:\ProgramData\MHTAgentic\debug\step6_elements.log")
            debug_log.parent.mkdir(parents=True, exist_ok=True)
            try:
                sic_names = []
                all_names = []
                for elem in chart_win.descendants():
                    try:
                        ename = (elem.element_info.name or '').strip()
                        if ename:
                            all_names.append(ename)
                        if 'sic' in ename.lower():
                            sic_names.append(ename)
                    except:
                        pass
                with open(debug_log, "w") as df:
                    df.write(f"SIC elements: {sic_names}\n\n")
                    df.write(f"All elements ({len(all_names)}):\n")
                    for n in all_names:
                        df.write(f"  {n}\n")
                logger.info(f"[Step 6] SIC elements: {sic_names}")
            except Exception as dbg_err:
                logger.info(f"[Step 6] Debug scan error: {dbg_err}")

            # Click "SIC COMMON ORDERS" TabItem specifically
            found = False
            for _ in range(30):
                try:
                    # Match exactly "SIC COMMON ORDERS" TabItem
                    sic_elem = chart_win.child_window(title='SIC COMMON ORDERS', control_type='TabItem')
                    if sic_elem.exists(timeout=0):
                        rect = sic_elem.rectangle()
                        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                        found = True
                        logger.info(f"[Step 6] Clicked 'SIC COMMON ORDERS' TabItem")
                        break
                except:
                    pass

                if not found:
                    # Fallback: scan TabItems for anything containing "SIC" and "COMMON"
                    try:
                        for elem in chart_win.descendants(control_type='TabItem'):
                            try:
                                ename = (elem.element_info.name or '').strip().upper()
                                if 'SIC' in ename and 'COMMON' in ename:
                                    rect = elem.rectangle()
                                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                    found = True
                                    logger.info(f"[Step 6] Clicked TabItem: '{ename}'")
                                    break
                            except:
                                pass
                    except:
                        pass
                if found:
                    break
                time.sleep(0.1)

            # ===== STEP 7: Click on assessment in SIC list =====
            logger.info(f"[Step 7] Finding {assess_name} via pywinauto...")
            self._update_overlay_step(7, 21, f"Selecting {assess_name}...")
            time.sleep(0.3)

            # Find the target assessment Text element in UIA tree
            target_elem = None
            for elem in chart_win.descendants(control_type='Text'):
                try:
                    auto_id = elem.element_info.automation_id or ''
                    if auto_id == 'ContentItemTextBlock':
                        ename = (elem.element_info.name or '').lower()
                        if any(kw in ename for kw in keywords):
                            target_elem = elem
                            break
                except:
                    pass

            if not target_elem:
                logger.error(f"Assessment '{assess_name}' not found in UIA tree")
                return False

            target_name = target_elem.element_info.name or assess_name
            logger.info(f"[Step 7] Found: '{target_name}'")

            # Get SIC group visible area
            try:
                sic_group = chart_win.child_window(title='SIC COMMON ORDERS', control_type='Group')
                sic_rect = sic_group.rectangle()
            except:
                sic_rect = chart_win.rectangle()

            # If off-screen, scroll until visible (retry up to 5 times)
            sic_cx = (sic_rect.left + sic_rect.right) // 2
            sic_cy = (sic_rect.top + sic_rect.bottom) // 2
            for scroll_attempt in range(5):
                target_rect = target_elem.rectangle()
                if target_rect.top >= sic_rect.top and target_rect.bottom <= sic_rect.bottom:
                    break  # Visible
                logger.info(f"Off-screen (y={target_rect.top}, visible={sic_rect.top}-{sic_rect.bottom}), scrolling attempt {scroll_attempt+1}...")
                pyautogui.moveTo(sic_cx, sic_cy)
                pyautogui.scroll(-150)
                time.sleep(0.15)
                target_rect = target_elem.rectangle()
                logger.info(f"After scroll: y={target_rect.top}")

            # Click the assessment
            target_rect = target_elem.rectangle()
            pyautogui.click((target_rect.left + target_rect.right) // 2,
                            (target_rect.top + target_rect.bottom) // 2)
            logger.info(f"[Step 7] Clicked '{target_name}'")
            time.sleep(0.2)

            # ===== STEP 8: Double-click description in procedures grid to open form =====
            logger.info("[Step 8] Opening assessment form...")
            self._update_overlay_step(8, 21, "Opening assessment form...")

            desc_elem = None
            for _ in range(30):
                for elem in chart_win.descendants(control_type='Text'):
                    try:
                        ename = (elem.element_info.name or '').lower()
                        auto_id = elem.element_info.automation_id or ''
                        if auto_id != 'ContentItemTextBlock' and any(kw in ename for kw in keywords):
                            desc_elem = elem
                            break
                    except:
                        pass
                if desc_elem:
                    break
                time.sleep(0.1)

            if not desc_elem:
                logger.error("Assessment description not found in procedures grid")
                return False

            rect = desc_elem.rectangle()
            pyautogui.doubleClick((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
            logger.info(f"[Step 8] Double-clicked description to open form")

            # ===== STEP 9: Fill out form =====
            logger.info("[Step 9] Filling out form...")
            self._update_overlay_step(9, 21, "Filling assessment form...")

            # Wait for radio buttons to appear
            all_radios = []
            for _ in range(40):
                all_radios = []
                for elem in chart_win.descendants(control_type='RadioButton'):
                    try:
                        rect = elem.rectangle()
                        name = elem.element_info.name or ''
                        all_radios.append({'name': name, 'elem': elem, 'y': rect.top, 'x': rect.left})
                    except:
                        pass
                if len(all_radios) >= 8:
                    break
                time.sleep(0.1)

            if not all_radios:
                logger.error("No radio buttons found - form may not have opened")
                return False

            # Sort by y then x, group into rows
            all_radios.sort(key=lambda r: (r['y'], r['x']))
            rows = []
            current_row = []
            current_y = -100
            for btn in all_radios:
                if btn['y'] > current_y + 20:
                    if current_row:
                        rows.append(current_row)
                    current_row = [btn]
                    current_y = btn['y']
                else:
                    current_row.append(btn)
            if current_row:
                rows.append(current_row)

            logger.info(f"[Step 9] Found {len(rows)} rows, {len(all_radios)} radio buttons")

            # Row 0 = patient/caregiver (2 buttons), click "patient"
            # Rows 1-7 (GAD-7) or 1-9 (PHQ-9) = scored questions (4 buttons each)
            # Last row = Q10 difficulty (4 buttons, vertical layout)

            # 9a: Click "patient" radio
            if rows and len(rows[0]) == 2:
                patient_btn = rows[0][0]  # First button = "patient"
                rect = patient_btn['elem'].rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                logger.info(f"[Step 9a] Selected 'patient' radio")
                scored_rows = rows[1:]  # Skip patient/caregiver row
            else:
                scored_rows = rows

            # Separate scored question rows (4 buttons horizontal) from Q10 difficulty (4 buttons vertical)
            question_rows = []
            difficulty_row = []
            for row in scored_rows:
                if len(row) == 4:
                    question_rows.append(row)
                elif len(row) == 1:
                    difficulty_row.append(row[0])

            # If difficulty buttons are individual rows (vertical layout), group them
            if not difficulty_row and len(question_rows) > 0:
                # Check if last few rows might be difficulty (single-column vertical)
                # Difficulty buttons: "Not difficult at all", "Somewhat difficult", etc.
                for row in scored_rows:
                    for btn in row:
                        if 'difficult' in btn['name'].lower():
                            difficulty_row.append(btn)

            # 9b: Fill scored questions (Q1-Q7 or Q1-Q9)
            for i, score in enumerate(scores):
                if i < len(question_rows) and score < len(question_rows[i]):
                    btn = question_rows[i][score]
                    rect = btn['elem'].rectangle()
                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

            logger.info(f"[Step 9b] Filled {min(len(scores), len(question_rows))} scored questions")

            # 9c: Scroll DOWN -1000 to snap to bottom of form
            logger.info("[Step 9c] Scrolling to bottom of form...")
            chart_rect = chart_win.rectangle()
            form_cx = (chart_rect.left + chart_rect.right) // 2
            form_cy = (chart_rect.top + chart_rect.bottom) // 2
            pyautogui.moveTo(form_cx, form_cy)
            pyautogui.scroll(-1000)
            time.sleep(0.15)

            # Re-find radio buttons after scroll (positions changed)
            difficulty_btns = []
            for elem in chart_win.descendants(control_type='RadioButton'):
                try:
                    name = elem.element_info.name or ''
                    if 'difficult' in name.lower():
                        rect = elem.rectangle()
                        difficulty_btns.append({'name': name, 'elem': elem, 'y': rect.top})
                except:
                    pass
            difficulty_btns.sort(key=lambda b: b['y'])

            # 9d: Answer Q10 difficulty
            if difficulty_btns:
                if total_score < 10:
                    diff_btn = difficulty_btns[0]  # "Not difficult at all"
                else:
                    diff_btn = difficulty_btns[1] if len(difficulty_btns) > 1 else difficulty_btns[0]
                rect = diff_btn['elem'].rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                logger.info(f"[Step 9d] Selected difficulty: '{diff_btn['name']}'")
            else:
                logger.warning("[Step 9d] No difficulty radio buttons found")

            # ===== STEP 10: Enter total score =====
            # Acquire typing lock to serialize keyboard input across RDP sessions
            from mhtagentic.db.database import typing_lock_acquire, typing_lock_release
            _typing_lock_held = False
            typing_lock_acquire(str(self.db_path), self._bot_name, timeout=60)
            _typing_lock_held = True
            logger.info(f"[Step 10] Typing lock acquired by {self._bot_name}")

            logger.info(f"[Step 10] Entering total score: {total_score}...")
            self._update_overlay_step(10, 21, f"Score: {total_score}...")

            score_entered = False
            for elem in chart_win.descendants(control_type='Edit'):
                try:
                    ename = (elem.element_info.name or '').lower()
                    if 'total' in ename or 'add up' in ename or 'record' in ename:
                        rect = elem.rectangle()
                        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                        pyautogui.hotkey('ctrl', 'a')
                        pyautogui.typewrite(str(total_score), interval=0.02)
                        score_entered = True
                        logger.info(f"[Step 10] Entered total score: {total_score}")
                        break
                except:
                    pass

            if not score_entered:
                logger.warning("Could not find total score field")

            # ===== STEP 10a: Click Enter button after total score =====
            enter_clicked = False
            for elem in chart_win.descendants(control_type='Button'):
                try:
                    btn_name = (elem.element_info.name or '').strip().lower()
                    auto_id = (elem.element_info.automation_id or '').strip().lower()
                    if btn_name == 'enter' or 'enter' in auto_id or 'cmdenter' in auto_id:
                        rect = elem.rectangle()
                        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                        enter_clicked = True
                        logger.info(f"[Step 10a] Clicked Enter button")
                        break
                except:
                    pass
            if not enter_clicked:
                # Log all buttons to help debug
                for elem in chart_win.descendants(control_type='Button'):
                    try:
                        bn = (elem.element_info.name or '').strip()
                        ai = (elem.element_info.automation_id or '').strip()
                        if bn or ai:
                            logger.info(f"[Step 10a] Available button: name='{bn}' auto_id='{ai}'")
                    except:
                        pass
                logger.warning("[Step 10a] Enter button not found")
            time.sleep(0.2)

            # ===== STEP 10b: Enter referral note =====
            referral_text = "Referral – Yes" if total_score >= 10 else "Referral – No"
            logger.info(f"[Step 10b] Entering note: {referral_text}")
            self._update_overlay_step(10, 21, f"Notes: {referral_text}")

            notes_entered = False
            for elem in chart_win.descendants(control_type='Edit'):
                try:
                    auto_id = (elem.element_info.automation_id or '')
                    if auto_id == 'txtNotes':
                        rect = elem.rectangle()
                        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                        pyautogui.hotkey('ctrl', 'a')
                        pyautogui.typewrite(referral_text, interval=0.02)
                        notes_entered = True
                        logger.info(f"[Step 10b] Entered notes: {referral_text}")
                        break
                except:
                    pass

            if not notes_entered:
                logger.warning("Could not find Notes field")

            # Release typing lock — Steps 10-10b typing complete
            if _typing_lock_held:
                typing_lock_release(str(self.db_path), self._bot_name)
                _typing_lock_held = False
                logger.info("[Step 10b] Typing lock released")

            # ===== STEP 11: Scroll back up, cancel form, delete assessment =====
            logger.info("[Step 11] Scrolling back up...")
            self._update_overlay_step(11, 21, "Closing form...")

            pyautogui.moveTo(form_cx, form_cy)
            pyautogui.scroll(1000)
            time.sleep(0.15)

            # Click Cancel to close the form (data is captured, don't save to Experity)
            for elem in chart_win.descendants(control_type='Button'):
                try:
                    auto_id = (elem.element_info.automation_id or '')
                    if auto_id in ('cmdCancel2', 'cmdCancel'):
                        rect = elem.rectangle()
                        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                        logger.info(f"[Step 11] Clicked Cancel ({auto_id})")
                        break
                except:
                    pass

            time.sleep(0.2)

            # ===== STEP 12: Delete the assessment from procedures grid =====
            logger.info("[Step 12] Deleting assessment from grid...")
            self._update_overlay_step(12, 21, "Removing assessment...")

            # Find and click the delete button for the procedure row
            delete_clicked = False
            for elem in chart_win.descendants(control_type='Button'):
                try:
                    btn_name = (elem.element_info.name or '').lower()
                    auto_id = (elem.element_info.automation_id or '').lower()
                    if 'delete' in btn_name or 'delete' in auto_id or 'remove' in btn_name:
                        rect = elem.rectangle()
                        if rect.top > 200 and rect.top < 300:  # In the procedures grid area
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            delete_clicked = True
                            logger.info(f"[Step 12] Clicked delete button")
                            break
                except:
                    pass

            if not delete_clicked:
                # Try right-clicking the procedure row and selecting delete
                for elem in chart_win.descendants(control_type='Text'):
                    try:
                        ename = (elem.element_info.name or '').lower()
                        auto_id = elem.element_info.automation_id or ''
                        if auto_id != 'ContentItemTextBlock' and any(kw in ename for kw in keywords):
                            rect = elem.rectangle()
                            pyautogui.rightClick((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            time.sleep(0.15)
                            # Look for Delete in context menu
                            pyautogui.press('delete')
                            logger.info("[Step 12] Right-clicked and pressed Delete")
                            delete_clicked = True
                            break
                    except:
                        pass

            time.sleep(0.15)

            # Handle any "Are you sure?" confirmation dialog
            if delete_clicked:
                time.sleep(0.15)
                for elem in chart_win.descendants(control_type='Button'):
                    try:
                        btn_name = (elem.element_info.name or '').lower()
                        if btn_name in ('yes', 'ok', 'delete', 'confirm'):
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            logger.info(f"[Step 12] Confirmed delete: '{elem.element_info.name}'")
                            break
                    except:
                        pass

            time.sleep(0.15)

            # ===== STEP 13: Close chart =====
            logger.info("[Step 13] Closing chart...")
            self._update_overlay_step(13, 21, "Closing chart...")

            # Step 13a: Click Close/X button on the chart
            close_clicked = False
            buttons = chart_win.descendants(control_type='Button')
            for btn in buttons:
                try:
                    btn_text = btn.window_text()
                    if btn_text in ['Close', 'X', 'close', '\ue5cd', '\u2715', '\u2716']:
                        rect = btn.rectangle()
                        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                        close_clicked = True
                        logger.info(f"Clicked close button: '{btn_text}'")
                        break
                except:
                    continue
            if not close_clicked:
                win_rect = chart_win.rectangle()
                pyautogui.click(win_rect.right - 30, win_rect.top + 40)
                logger.info("Used fallback close click (top-right)")

            # Step 13b: Handle "Close Chart" confirmation dialog
            time.sleep(0.15)
            confirmed = False
            try:
                buttons2 = chart_win.descendants(control_type='Button')
                for btn in buttons2:
                    try:
                        if btn.window_text() == 'Close Chart':
                            rect = btn.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            logger.info("Clicked 'Close Chart' confirmation")
                            confirmed = True
                            break
                    except:
                        continue
            except:
                pass

            # Step 13c: Try finding dialog as separate window
            if not confirmed:
                try:
                    dialogs = session_find_elements(title_re='.*Close Chart.*', backend='uia')
                    if dialogs:
                        dlg_app = Application(backend='uia').connect(handle=dialogs[0].handle, timeout=2)
                        dlg_win = dlg_app.window(handle=dialogs[0].handle)
                        close_chart_btn = dlg_win.child_window(control_type='Button', title='Close Chart')
                        close_chart_btn.click_input()
                        logger.info("Clicked 'Close Chart' via dialog window")
                except:
                    pass

            time.sleep(0.15)  # Brief settle after chart close

            # ===== STEP 14: Click patient name =====
            logger.info("[Step 14] Clicking patient name...")
            self._update_overlay_step(14, 21, "Opening patient demographics...")
            elements = session_find_elements(title_re='.*Tracking Board.*', backend='uia')
            target_handle = None
            for elem in elements:
                if 'Setup' not in elem.name:
                    target_handle = elem.handle
                    break

            tb_app = Application(backend='uia').connect(handle=target_handle, timeout=3)
            tb_win = tb_app.window(handle=target_handle)

            # Search in both Waiting Room and Roomed Patients
            clicked = False
            for section_name in ['.*Waiting Room.*', '.*Roomed Patients.*']:
                if clicked:
                    break
                try:
                    section_group = tb_win.child_window(title_re=section_name, control_type='Group')
                    for di in section_group.descendants(control_type='DataItem'):
                        try:
                            for child in di.descendants():
                                if patient_upper in child.window_text().upper() and ',' in child.window_text():
                                    rect = child.rectangle()
                                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                    clicked = True
                                    break
                            if clicked:
                                break
                        except:
                            pass
                except:
                    pass

            # ===== STEP 15: Click Documents tab =====
            logger.info("[Step 15] Clicking Documents tab...")
            # Poll for demographics window and Documents tab
            for _ in range(30):
                try:
                    demo_app, demo_win = session_connect(title_re=f'.*{name_part}.*', timeout=2)
                    found = False
                    for elem in demo_win.descendants():
                        try:
                            if elem.element_info.control_type == 'TabItem' and 'documents' in (elem.element_info.name or '').lower():
                                rect = elem.rectangle()
                                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                found = True
                                break
                        except:
                            pass
                    if found:
                        break
                except:
                    pass

            # ===== STEP 16: Click Scan/Upload =====
            logger.info("[Step 16] Waiting for Scan/Upload button...")
            self._update_overlay_step(16, 21, "Waiting for Scan/Upload...")

            # Wait for Scan/Upload button to actually appear (poll up to 10 seconds)
            scan_upload_found = False
            for attempt in range(50):  # 50 * 0.2s = 10 seconds max
                try:
                    demo_app, demo_win = session_connect(title_re=f'.*{name_part}.*', timeout=2)
                    for elem in demo_win.descendants():
                        try:
                            n = elem.element_info.name or ''
                            if 'scan' in n.lower() and 'upload' in n.lower():
                                rect = elem.rectangle()
                                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                scan_upload_found = True
                                logger.info("Found and clicked Scan/Upload button")
                                break
                        except:
                            pass
                    if scan_upload_found:
                        break
                except:
                    pass

            if not scan_upload_found:
                # Fallback: proportional window-relative position for Scan/Upload button
                demo_rect = demo_win.rectangle()
                win_w = demo_rect.right - demo_rect.left
                win_h = demo_rect.bottom - demo_rect.top
                scan_x = demo_rect.right - int(win_w * 0.03)
                scan_y = demo_rect.top + int(win_h * 0.05)
                pyautogui.click(scan_x, scan_y)
                logger.info(f"Used proportional fallback for Scan/Upload at ({scan_x}, {scan_y})")

            # ===== STEP 17: Close TWAIN popup =====
            logger.info("[Step 17] Closing TWAIN popup...")
            self._update_overlay_step(17, 21, "Closing TWAIN popup...")

            # Poll for the TWAIN popup to appear then close it.
            # Strategy: try Cancel button first (TWAIN install dialog), then gray X,
            # then Escape fallback.
            twain_closed = False
            for attempt in range(100):  # Poll up to ~15s at 0.15s intervals
                try:
                    demo_app, demo_win = session_connect(title_re=f'.*{name_part}.*', timeout=1)

                    # Method 1: Click Cancel/Close button in the TWAIN dialog
                    if not twain_closed:
                        for btn in demo_win.descendants(control_type='Button'):
                            try:
                                bt = btn.window_text()
                                if bt in ('Cancel', 'Close', 'cancel', 'close'):
                                    btn.click_input()
                                    twain_closed = True
                                    logger.info(f"Closed TWAIN popup via [{bt}] button")
                                    break
                            except:
                                continue

                    # Method 2: Click the gray X relative to Custom 'scan/upload' element
                    if not twain_closed:
                        for elem in demo_win.descendants():
                            try:
                                ct = elem.element_info.control_type
                                n = elem.element_info.name or ''
                                if ct == 'Custom' and 'scan/upload' in n.lower():
                                    r = elem.rectangle()
                                    cw = r.right - r.left
                                    ch = r.bottom - r.top
                                    # Gray X is at 112.3% width, 15.3% height from Custom top-left
                                    x = r.left + int(cw * 1.123)
                                    y = r.top + int(ch * 0.153)
                                    pyautogui.click(x, y)
                                    twain_closed = True
                                    logger.info(f"Closed TWAIN popup via relative click at ({x}, {y})")
                                    break
                            except:
                                pass

                    if twain_closed:
                        break
                except:
                    pass
                time.sleep(0.15)

            if not twain_closed:
                pyautogui.press('escape')
                logger.info("Closed TWAIN popup via Escape key fallback")
                time.sleep(0.5)

            # ===== STEP 18: Select Description dropdown =====
            logger.info("[Step 18] Selecting Mental/Behavioral...")
            for _ in range(20):  # Poll for dropdown
                found = False
                for elem in demo_win.descendants():
                    try:
                        if elem.element_info.control_type == 'ComboBox' and 'mat-select' in (elem.element_info.automation_id or ''):
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # Move to window center before scrolling
            demo_rect = demo_win.rectangle()
            demo_cx = (demo_rect.left + demo_rect.right) // 2
            demo_cy = (demo_rect.top + demo_rect.bottom) // 2
            pyautogui.moveTo(demo_cx, demo_cy)
            pyautogui.scroll(-1200)

            for _ in range(20):  # Poll for Mental list item
                found = False
                demo_rect_check = demo_win.rectangle()
                min_top = demo_rect_check.top + int((demo_rect_check.bottom - demo_rect_check.top) * 0.08)
                for elem in demo_win.descendants():
                    try:
                        if elem.element_info.control_type == 'ListItem' and 'mental' in (elem.element_info.name or '').lower():
                            rect = elem.rectangle()
                            if rect.top > min_top:
                                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                found = True
                                break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 19: Cancel =====
            logger.info("[Step 19] Clicking Cancel...")
            for _ in range(20):  # Poll for Cancel
                found = False
                for elem in demo_win.descendants():
                    try:
                        if elem.element_info.control_type == 'Button' and (elem.element_info.name or '').lower() == 'cancel':
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 20: Click OK =====
            logger.info("[Step 20] Clicking OK...")
            for _ in range(20):  # Poll for OK
                found = False
                for elem in demo_win.descendants():
                    try:
                        if elem.element_info.control_type == 'Button' and (elem.element_info.name or '').lower() == 'ok':
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 21: Close Demographics =====
            logger.info("[Step 21] Closing Demographics...")
            self._update_overlay_step(21, 21, "Closing demographics tab...")
            close_buttons = []
            demo_rect = demo_win.rectangle()
            # Use window-relative thresholds instead of absolute pixels
            title_zone_top = demo_rect.top + int((demo_rect.bottom - demo_rect.top) * 0.05)
            title_zone_bottom = demo_rect.top + int((demo_rect.bottom - demo_rect.top) * 0.15)
            for elem in demo_win.descendants():
                try:
                    if elem.element_info.control_type == 'Button' and elem.element_info.automation_id == 'PART_CloseButton':
                        rect = elem.rectangle()
                        if rect.top > title_zone_top and rect.top < title_zone_bottom:
                            close_buttons.append(rect)
                except:
                    pass

            if len(close_buttons) >= 2:
                rect = close_buttons[1]
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

            logger.info("Outbound flow completed successfully!")
            return True

        except Exception as e:
            # Safety release typing lock if still held
            try:
                if locals().get('_typing_lock_held'):
                    typing_lock_release(str(self.db_path), self._bot_name)
            except Exception:
                pass
            logger.error(f"Error in outbound flow: {e}")
            return False
