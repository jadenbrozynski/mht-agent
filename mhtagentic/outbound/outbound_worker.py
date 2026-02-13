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
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, List, Callable, TYPE_CHECKING

from pywinauto import Application, findwindows, Desktop
import pyautogui

if TYPE_CHECKING:
    from mhtagentic.desktop.control_overlay import ControlOverlay

logger = logging.getLogger(__name__)


class OutboundWorker:
    """
    Worker that polls for completed outbound events and processes them.

    Runs in a background thread, checking every poll_interval seconds
    for outbound events with status=100 (ready to process).
    """

    # Status codes
    STATUS_READY = 100      # Ready to process (from MHT)
    STATUS_PROCESSING = 50  # Currently being processed
    STATUS_DONE = 200       # Successfully entered into Experity
    STATUS_ERROR = -100     # Failed to process

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

    def _show_all_overlays(self):
        """Show ALL overlays after outbound processing."""
        if self.demo_mode:
            return  # No overlays to restore in demo mode
        try:
            from mhtagentic.desktop.control_overlay import show_all_overlays
            show_all_overlays()
            logger.info("All overlays restored after outbound processing")
        except Exception as e:
            logger.warning(f"Show all overlays failed: {e}")

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
        Call this during downtime (e.g., refresh wait) instead of auto-polling.
        Returns number of events processed.
        """
        processed = 0
        try:
            events = self._get_ready_outbound_events()
            for event in events:
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
        """Get outbound events ready to process (status=100)."""
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
        """Main worker loop."""
        while self._running:
            try:
                events = self._get_ready_outbound_events()
                if events:
                    for event in events:
                        if not self._running:
                            break
                        self._process_event(event)
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

        # Hide ALL overlays during outbound processing to avoid click interference
        self._hide_all_overlays()

        try:
            # Parse the event data
            raw_data = json.loads(event['raw_data'])
            data = raw_data.get('data', raw_data)
            patient = data.get('patient', {})
            patient_name = f"{patient.get('patient_last_name')}, {patient.get('patient_first_name')}"

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
            elements = findwindows.find_elements(title_re='.*Tracking Board.*', backend='uia')
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
            chart_icon = images[0]
            rect = chart_icon.rectangle()
            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

            # ===== STEP 4: Wait for chart window =====
            logger.info("[Step 4] Waiting for chart window...")
            self._update_overlay_step(4, 21, "Loading chart...")
            chart_win = None
            for _ in range(30):  # Poll for chart window
                try:
                    chart_app = Application(backend='uia').connect(title_re=f'.*{name_part}.*', timeout=1)
                    chart_win = chart_app.window(title_re=f'.*{name_part}.*')
                    break
                except:
                    pass
            if not chart_win:
                logger.error("Chart window not found")
                return False

            # ===== STEP 5: Navigate to Procedures/Supplies =====
            logger.info("[Step 5] Clicking Procedures/Supplies...")
            self._update_overlay_step(5, 21, "Opening Procedures/Supplies...")
            for _ in range(20):  # Poll for element
                found = False
                for elem in chart_win.descendants():
                    try:
                        if elem.element_info.control_type == 'TabItem' and 'procedures' in (elem.element_info.name or '').lower():
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 6: Navigate to SIC COMMON ORDERS =====
            logger.info("[Step 6] Clicking SIC COMMON ORDERS...")
            self._update_overlay_step(6, 21, "Opening SIC Common Orders...")
            for _ in range(20):  # Poll for element
                found = False
                for elem in chart_win.descendants():
                    try:
                        name = elem.element_info.name or ''
                        if 'sic common' in name.lower():
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 7: Click on assessment in list =====
            logger.info(f"[Step 7] Clicking {assess_name}...")
            self._update_overlay_step(7, 21, f"Selecting {assess_name}...")
            for _ in range(20):  # Poll for element
                found = False
                for elem in chart_win.descendants():
                    try:
                        name = elem.element_info.name or ''
                        ctrl_type = elem.element_info.control_type
                        auto_id = elem.element_info.automation_id or ''
                        if ctrl_type == 'Text' and auto_id == 'ContentItemTextBlock' and any(kw in name.lower() for kw in keywords):
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 8: Click assessment description header =====
            logger.info("[Step 8] Clicking description header...")
            self._update_overlay_step(8, 21, "Opening assessment form...")
            for _ in range(20):  # Poll for element
                found = False
                for elem in chart_win.descendants():
                    try:
                        name = elem.element_info.name or ''
                        ctrl_type = elem.element_info.control_type
                        auto_id = elem.element_info.automation_id or ''
                        if ctrl_type == 'Text' and auto_id != 'ContentItemTextBlock' and any(kw in name.lower() for kw in keywords):
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 9: Fill out the form =====
            logger.info("[Step 9] Filling out form...")
            self._update_overlay_step(9, 21, "Filling assessment scores...")
            # Poll for radio buttons to appear
            all_radios = []
            for _ in range(20):
                all_radios = []
                for elem in chart_win.descendants():
                    try:
                        if elem.element_info.control_type == 'RadioButton':
                            rect = elem.rectangle()
                            all_radios.append({'name': elem.element_info.name, 'elem': elem, 'y': rect.top, 'x': rect.left})
                    except:
                        pass
                if len(all_radios) > 5:  # Got enough radio buttons
                    break

            all_radios.sort(key=lambda x: (x['y'], x['x']))

            rows = []
            current_row = []
            current_y = -100
            for btn in all_radios:
                if btn['y'] > current_y + 30:
                    if current_row:
                        rows.append(current_row)
                    current_row = [btn]
                    current_y = btn['y']
                else:
                    current_row.append(btn)
            if current_row:
                rows.append(current_row)

            # Click patient (row 0)
            if rows:
                btn = rows[0][0]
                rect = btn['elem'].rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

            # Fill questions based on scores - NO DELAYS
            for i, score in enumerate(scores):
                row_idx = i + 1
                if row_idx < len(rows) and score < len(rows[row_idx]):
                    btn = rows[row_idx][score]
                    rect = btn['elem'].rectangle()
                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

            # Q10 difficulty
            if len(rows) > 10:
                btn = rows[10][0]
                rect = btn['elem'].rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

            # ===== STEP 10: Enter total score =====
            logger.info("[Step 10] Entering total score...")
            self._update_overlay_step(10, 21, f"Entering total score: {total_score}...")
            # Click center of chart window to ensure focus before scrolling
            chart_rect = chart_win.rectangle()
            chart_center_x = (chart_rect.left + chart_rect.right) // 2
            chart_center_y = (chart_rect.top + chart_rect.bottom) // 2
            pyautogui.click(chart_center_x, chart_center_y)
            pyautogui.scroll(-500)

            for _ in range(20):  # Poll for total score field
                found = False
                for elem in chart_win.descendants():
                    try:
                        if elem.element_info.control_type == 'Edit' and 'total score' in (elem.element_info.name or '').lower():
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            pyautogui.hotkey('ctrl', 'a')
                            pyautogui.typewrite(str(total_score))
                            pyautogui.press('enter')
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 11: Cancel and delete =====
            logger.info("[Step 11] Clicking Cancel...")
            self._update_overlay_step(11, 21, "Cancelling form...")
            pyautogui.scroll(500)

            for _ in range(20):  # Poll for cancel button
                found = False
                for elem in chart_win.descendants():
                    try:
                        if elem.element_info.control_type == 'Button' and elem.element_info.automation_id == 'cmdCancel2':
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 12: Delete assessment =====
            logger.info("[Step 12] Deleting assessment...")
            self._update_overlay_step(12, 21, "Deleting assessment...")
            phq_y = None
            for elem in chart_win.descendants():
                try:
                    name = elem.element_info.name or ''
                    if any(kw in name.lower() for kw in keywords):
                        phq_y = elem.rectangle().top
                        break
                except:
                    pass

            for _ in range(20):  # Poll for delete button
                found = False
                for elem in chart_win.descendants():
                    try:
                        if elem.element_info.control_type == 'Button' and 'delete' in (elem.element_info.name or '').lower():
                            rect = elem.rectangle()
                            if phq_y and abs(rect.top - phq_y) < 30:
                                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                                found = True
                                break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # Poll for OK confirmation
            for _ in range(20):
                found = False
                for elem in chart_win.descendants():
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

            # ===== STEP 13: Close chart =====
            logger.info("[Step 13] Closing chart...")
            self._update_overlay_step(13, 21, "Saving and closing chart...")
            for _ in range(20):  # Poll for save button
                found = False
                for elem in chart_win.descendants():
                    try:
                        if elem.element_info.control_type == 'Button' and elem.element_info.automation_id == 'SaveButton':
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            for _ in range(20):  # Poll for btnSave
                found = False
                for elem in chart_win.descendants():
                    try:
                        if elem.element_info.control_type == 'Button' and elem.element_info.automation_id == 'btnSave':
                            rect = elem.rectangle()
                            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                            found = True
                            break
                    except:
                        pass
                if found:
                    break
                time.sleep(0.05)  # Small delay to prevent UI overload

            # ===== STEP 14: Click patient name =====
            logger.info("[Step 14] Clicking patient name...")
            self._update_overlay_step(14, 21, "Opening patient demographics...")
            elements = findwindows.find_elements(title_re='.*Tracking Board.*', backend='uia')
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
                    demo_app = Application(backend='uia').connect(title_re=f'.*{name_part}.*', timeout=1)
                    demo_win = demo_app.window(title_re=f'.*{name_part}.*')
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
                    demo_app = Application(backend='uia').connect(title_re=f'.*{name_part}.*', timeout=2)
                    demo_win = demo_app.window(title_re=f'.*{name_part}.*')
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

            # ===== STEP 17: Wait for TWAIN popup =====
            logger.info("[Step 17] Waiting for TWAIN popup...")
            self._update_overlay_step(17, 21, "Closing TWAIN popup...")

            # Wait for TWAIN popup to appear (poll up to 10 seconds)
            twain_closed = False
            for attempt in range(50):  # 50 * 0.2s = 10 seconds max
                try:
                    desktop = Desktop(backend='uia')
                    demo_win = None
                    for w in desktop.windows():
                        try:
                            if name_part.upper() in w.window_text().upper():
                                demo_win = w
                                break
                        except:
                            pass

                    if demo_win:
                        # Strategy 1: Find close/X button near the TWAIN element
                        twain_elem = None
                        for elem in demo_win.descendants():
                            try:
                                n = elem.element_info.name or ''
                                if 'dynamic web twain' in n.lower():
                                    twain_elem = elem
                                    break
                            except:
                                pass

                        if twain_elem:
                            twain_rect = twain_elem.rectangle()
                            # Look for a close/X button near the TWAIN popup
                            close_found = False
                            for elem in demo_win.descendants():
                                try:
                                    ct = elem.element_info.control_type
                                    n = (elem.element_info.name or '').lower()
                                    if ct == 'Button' and n in ('close', 'x', '\u2715', '\u2716', '\ue5cd'):
                                        btn_rect = elem.rectangle()
                                        # Must be near the TWAIN popup (within reasonable range)
                                        if abs(btn_rect.top - twain_rect.top) < 150:
                                            pyautogui.click((btn_rect.left + btn_rect.right) // 2, (btn_rect.top + btn_rect.bottom) // 2)
                                            close_found = True
                                            logger.info("Closed TWAIN popup via close button")
                                            break
                                except:
                                    pass

                            if not close_found:
                                # Fallback: use proportional offset from TWAIN element
                                # Close button is typically above-right of the popup title
                                popup_width = twain_rect.right - twain_rect.left
                                popup_height = twain_rect.bottom - twain_rect.top
                                close_x = twain_rect.right + int(popup_width * 0.1)
                                close_y = twain_rect.top - int(popup_height * 0.5)
                                pyautogui.click(close_x, close_y)
                                logger.info("Closed TWAIN popup via proportional offset")

                            twain_closed = True

                    if twain_closed:
                        break
                except:
                    pass

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
            logger.error(f"Error in outbound flow: {e}")
            return False
