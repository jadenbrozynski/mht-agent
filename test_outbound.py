"""
MHT Agentic - Test Outbound Flow

Tests the outbound assessment retrieval flow:
1. Starts from monitoring screen (Experity Tracking Board)
2. Queries database for outbound events (status=10)
3. Finds the patient in Roomed Patients
4. Clicks on their chart
5. Navigates: Procedures/Supplies -> Common -> Behavioral Health Evals
6. Clicks on the assessment (PHQ-9 or GAD-7)

Usage:
    python test_outbound.py                    # Run full test
    python test_outbound.py --inspect          # Just inspect UI elements
    python test_outbound.py --patient "SMITH"  # Test with specific patient name
"""

import sys
import os
import time
import json
import argparse
import sqlite3
import logging
from pathlib import Path
from datetime import datetime

# Setup paths
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(SCRIPT_DIR))
os.chdir(str(SCRIPT_DIR))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger("test_outbound")

# Imports
import pyautogui
from pywinauto import Desktop, Application
from pywinauto import findwindows

# Database path
DB_PATH = SCRIPT_DIR / "output" / "mht_data.db"

# Configure pyautogui
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1


class OutboundFlowTester:
    """Tests the outbound assessment retrieval flow."""

    def __init__(self):
        self.experity_window = None
        self.chart_window = None
        self.current_patient_info = None  # Stores assessment data for current patient

    def find_experity_window(self):
        """Find the Experity/Tracking Board window using same method as launcher."""
        try:
            # Use findwindows to get the Tracking Board handle (same as launcher.pyw)
            elements = findwindows.find_elements(title_re='.*Tracking Board.*', backend='uia')

            # Find the one that's NOT "Setup"
            target_handle = None
            for elem in elements:
                if 'Setup' not in elem.name:
                    target_handle = elem.handle
                    break

            if target_handle:
                app = Application(backend='uia').connect(handle=target_handle, timeout=3)
                self.experity_window = app.window(handle=target_handle)
                logger.info(f"Connected to Tracking Board: {self.experity_window.window_text()}")
                return self.experity_window

            # Fallback: try Experity window
            app = Application(backend='uia').connect(title_re='.*Experity.*', found_index=0, timeout=3)
            self.experity_window = app.window(title_re='.*Experity.*', found_index=0)
            logger.info(f"Connected to Experity: {self.experity_window.window_text()}")
            return self.experity_window

        except Exception as e:
            logger.error(f"Experity window not found: {e}")
            return None

    def get_completed_outbound_events(self):
        """Query database for completed outbound events (status=100) that have assessment data."""
        if not DB_PATH.exists():
            logger.warning(f"Database not found: {DB_PATH}")
            return []

        try:
            conn = sqlite3.connect(str(DB_PATH))
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()

            # Get completed outbound events with assessment results
            cursor.execute("""
                SELECT id, received_at, raw_data, converted_data, response_data, status, kind
                FROM common_event
                WHERE direction = 'O' AND status = 100
                ORDER BY received_at DESC
                LIMIT 10
            """)

            events = [dict(row) for row in cursor.fetchall()]
            conn.close()

            logger.info(f"Found {len(events)} completed outbound events")
            return events
        except Exception as e:
            logger.error(f"Database error: {e}")
            return []

    def get_patient_info_from_event(self, event):
        """Extract patient name and assessment data from outbound event."""
        result = {
            'patient_name': None,
            'patient_first_name': None,
            'patient_last_name': None,
            'assessments': []
        }

        try:
            # Try raw_data first (has the full assessment data from MHT)
            if event.get('raw_data'):
                raw = event['raw_data']
                if isinstance(raw, str):
                    raw = json.loads(raw)

                data = raw.get('data', raw)
                patient = data.get('patient', {})

                result['patient_first_name'] = patient.get('patient_first_name', '')
                result['patient_last_name'] = patient.get('patient_last_name', '')
                result['patient_name'] = f"{result['patient_last_name']}, {result['patient_first_name']}".strip(', ')

                # Get assessment data
                assessments = data.get('assessment', [])
                if not isinstance(assessments, list):
                    assessments = [assessments]

                for assessment in assessments:
                    assess_info = {
                        'name': assessment.get('assessment_name', ''),
                        'type': assessment.get('assessment_type', ''),
                        'total_score': assessment.get('total_score_value'),
                        'severity': assessment.get('total_score_legend') or assessment.get('patient_score_legend_value'),
                        'items': []
                    }

                    # Get individual item scores (PHQ9_1, PHQ9_2, etc.)
                    items = assessment.get('assessment_items', [])
                    for item in items:
                        assess_info['items'].append({
                            'name': item.get('assessment_item_name', ''),
                            'score': item.get('assessment_item_score'),
                            'legend': item.get('assessment_item_legend_value', '')
                        })

                    result['assessments'].append(assess_info)

            # Also check converted_data for patient name
            if not result['patient_name'] and event.get('converted_data'):
                conv = event['converted_data']
                if isinstance(conv, str):
                    conv = json.loads(conv)
                summary = conv.get('summary', {})
                if summary.get('patient_name'):
                    result['patient_name'] = summary['patient_name']

            return result
        except Exception as e:
            logger.error(f"Error parsing event data: {e}")
            return result

    def get_patient_name_from_event(self, event):
        """Extract just patient name from event (for backwards compat)."""
        info = self.get_patient_info_from_event(event)
        return info.get('patient_name')

    def find_patient_in_roomed(self, patient_name):
        """Find a patient in the Roomed Patients section (same approach as launcher.pyw)."""
        if not self.experity_window:
            logger.error("No Experity window")
            return None

        logger.info(f"Searching for patient: {patient_name}")
        patient_upper = patient_name.upper()

        try:
            # Find the Roomed Patients group (same as launcher.pyw line 1884)
            roomed_group = self.experity_window.child_window(
                title_re='.*Roomed Patients.*',
                control_type='Group'
            )
            logger.info("Found Roomed Patients group")

            # Get all DataItem elements in the Roomed Patients section
            roomed_data_items = roomed_group.descendants(control_type='DataItem')
            logger.info(f"Found {len(roomed_data_items)} patient rows in Roomed Patients")

            for di in roomed_data_items:
                try:
                    # Quick text check on the DataItem itself
                    di_text = di.window_text()
                    if patient_upper in di_text.upper():
                        logger.info(f"Found patient row: {di_text[:50]}...")
                        return di

                    # Also check children/descendants for name
                    for child in di.descendants():
                        child_text = child.window_text()
                        if patient_upper in child_text.upper():
                            logger.info(f"Found patient via child: {child_text[:50]}...")
                            return di
                except:
                    continue

            logger.warning(f"Patient '{patient_name}' not found in Roomed Patients")
            return None

        except Exception as e:
            logger.error(f"Error searching Roomed Patients: {e}")

            # Fallback: try searching all descendants
            logger.info("Trying fallback search...")
            try:
                descendants = self.experity_window.descendants()
                for elem in descendants:
                    try:
                        ctrl_type = elem.element_info.control_type
                        name = elem.element_info.name or ""

                        if ctrl_type == "DataItem" and patient_upper in name.upper():
                            logger.info(f"Found patient row (fallback): {name[:50]}...")
                            return elem
                    except:
                        continue
            except:
                pass

            return None

    def click_chart_icon(self, patient_row):
        """Click the chart icon for a patient row (same as launcher.pyw line 1894-1921)."""
        try:
            # Find Image elements (icons) in the row
            images = patient_row.descendants(control_type='Image')

            if images:
                # First image is usually the chart icon
                chart_icon = images[0]
                rect = chart_icon.rectangle()
                cx = (rect.left + rect.right) // 2
                cy = (rect.top + rect.bottom) // 2

                logger.info(f"Clicking chart icon at ({cx}, {cy})")
                self._show_target(cx, cy, "Chart Icon")
                pyautogui.click(cx, cy)
                time.sleep(1.5)

                # Dismiss any popups that appear
                self._dismiss_popup_dialogs()

                return True
            else:
                logger.warning("No chart icon found in row")
                return False
        except Exception as e:
            logger.error(f"Error clicking chart icon: {e}")
            return False

    def _dismiss_popup_dialogs(self):
        """Dismiss common popup dialogs."""
        try:
            # Try pressing Enter to dismiss OK dialogs
            pyautogui.press('enter')
            time.sleep(0.3)
        except:
            pass

    def wait_for_chart_window(self, patient_name, timeout=5):
        """Wait for chart window to open."""
        start_time = time.time()
        name_part = patient_name.split(',')[0].strip()

        while time.time() - start_time < timeout:
            try:
                app = Application(backend='uia').connect(
                    title_re=f'.*{name_part}.*',
                    timeout=1
                )
                self.chart_window = app.window(title_re=f'.*{name_part}.*')
                logger.info("Chart window opened")
                return True
            except:
                time.sleep(0.5)

        logger.warning("Chart window did not open")
        return False

    def navigate_to_procedures_supplies(self):
        """Click on Procedures/Supplies tab in the chart."""
        if not self.chart_window:
            logger.error("No chart window")
            return False

        logger.info("Looking for Procedures/Supplies tab...")

        try:
            descendants = self.chart_window.descendants()

            # Look for Procedures/Supplies tab/button
            for elem in descendants:
                try:
                    name = elem.element_info.name or ""
                    ctrl_type = elem.element_info.control_type

                    if "procedure" in name.lower() and "suppl" in name.lower():
                        logger.info(f"Found: {name} ({ctrl_type})")
                        rect = elem.rectangle()
                        cx = (rect.left + rect.right) // 2
                        cy = (rect.top + rect.bottom) // 2

                        self._show_target(cx, cy, "Procedures/Supplies")
                        elem.click_input()
                        time.sleep(1.0)
                        return True
                except:
                    continue

            # Try searching for just "Procedures" or "Supplies" separately
            for elem in descendants:
                try:
                    name = elem.element_info.name or ""
                    ctrl_type = elem.element_info.control_type

                    if ctrl_type in ["TabItem", "Button", "TreeItem"] and (
                        "procedures" in name.lower() or "supplies" in name.lower()
                    ):
                        logger.info(f"Found: {name} ({ctrl_type})")
                        rect = elem.rectangle()
                        cx = (rect.left + rect.right) // 2
                        cy = (rect.top + rect.bottom) // 2

                        self._show_target(cx, cy, "Procedures/Supplies")
                        elem.click_input()
                        time.sleep(1.0)
                        return True
                except:
                    continue

            logger.warning("Procedures/Supplies tab not found")
            return False
        except Exception as e:
            logger.error(f"Error navigating to Procedures/Supplies: {e}")
            return False

    def navigate_to_sic_common(self):
        """Click on SIC COMMON ORDERS tab."""
        if not self.chart_window:
            return False

        logger.info("Looking for SIC COMMON ORDERS tab...")

        try:
            # Refresh descendants after previous click
            time.sleep(0.5)
            descendants = self.chart_window.descendants()

            # Look for SIC COMMON ORDERS tab specifically
            for elem in descendants:
                try:
                    name = elem.element_info.name or ""
                    ctrl_type = elem.element_info.control_type
                    auto_id = elem.element_info.automation_id or ""

                    # Match "SIC COMMON ORDERS" or similar
                    if "sic common" in name.lower() or "sic common" in auto_id.lower():
                        logger.info(f"Found: {name} ({ctrl_type})")
                        rect = elem.rectangle()
                        cx = (rect.left + rect.right) // 2
                        cy = (rect.top + rect.bottom) // 2

                        self._show_target(cx, cy, "SIC COMMON ORDERS")
                        elem.click_input()
                        time.sleep(1.0)
                        return True
                except:
                    continue

            # Fallback: try "Common Orders"
            for elem in descendants:
                try:
                    name = elem.element_info.name or ""
                    ctrl_type = elem.element_info.control_type

                    if ctrl_type == "TabItem" and "common order" in name.lower():
                        logger.info(f"Found fallback: {name} ({ctrl_type})")
                        rect = elem.rectangle()
                        cx = (rect.left + rect.right) // 2
                        cy = (rect.top + rect.bottom) // 2

                        self._show_target(cx, cy, "Common Orders")
                        elem.click_input()
                        time.sleep(1.0)
                        return True
                except:
                    continue

            logger.warning("SIC COMMON ORDERS tab not found")
            return False
        except Exception as e:
            logger.error(f"Error navigating to SIC COMMON ORDERS: {e}")
            return False

    def navigate_to_behavioral_health_evals(self):
        """Click on Behavioral Health Evals section."""
        if not self.chart_window:
            return False

        logger.info("Looking for Behavioral Health Evals...")

        try:
            time.sleep(0.5)
            descendants = self.chart_window.descendants()

            for elem in descendants:
                try:
                    name = elem.element_info.name or ""
                    ctrl_type = elem.element_info.control_type

                    if "behavioral" in name.lower() and "health" in name.lower():
                        logger.info(f"Found: {name} ({ctrl_type})")
                        rect = elem.rectangle()
                        cx = (rect.left + rect.right) // 2
                        cy = (rect.top + rect.bottom) // 2

                        self._show_target(cx, cy, "Behavioral Health Evals")
                        elem.click_input()
                        time.sleep(1.0)
                        return True

                    # Also try "BH Evals" or similar abbreviations
                    if "bh eval" in name.lower() or "bh screen" in name.lower():
                        logger.info(f"Found: {name} ({ctrl_type})")
                        rect = elem.rectangle()
                        cx = (rect.left + rect.right) // 2
                        cy = (rect.top + rect.bottom) // 2

                        self._show_target(cx, cy, "BH Evals")
                        elem.click_input()
                        time.sleep(1.0)
                        return True
                except:
                    continue

            logger.warning("Behavioral Health Evals not found")
            return False
        except Exception as e:
            logger.error(f"Error navigating to Behavioral Health Evals: {e}")
            return False

    def find_and_click_assessment(self, assessment_type=None):
        """Find and click on the specific assessment (PHQ-9 or GAD-7)."""
        if not self.chart_window:
            return False

        # Determine which assessment to look for from the event data
        if not assessment_type and self.current_patient_info:
            assessments = self.current_patient_info.get('assessments', [])
            if assessments:
                assessment_type = assessments[0].get('name', '')
                logger.info(f"Looking for specific assessment: {assessment_type}")

        if not assessment_type:
            logger.warning("No assessment type specified")
            return False

        logger.info(f"Looking for assessment: {assessment_type}")

        try:
            time.sleep(0.5)
            descendants = self.chart_window.descendants()

            # Build specific keywords based on assessment type
            assessment_lower = assessment_type.lower()

            if "phq" in assessment_lower or "patient health" in assessment_lower:
                keywords = ["phq-9", "phq9", "patient health questionnaire"]
            elif "gad" in assessment_lower or "anxiety" in assessment_lower:
                keywords = ["gad-7", "gad7", "generalized anxiety"]
            else:
                keywords = [assessment_lower]

            for elem in descendants:
                try:
                    name = elem.element_info.name or ""
                    ctrl_type = elem.element_info.control_type
                    name_lower = name.lower()

                    if any(kw in name_lower for kw in keywords):
                        logger.info(f"Found assessment: {name} ({ctrl_type})")
                        rect = elem.rectangle()
                        cx = (rect.left + rect.right) // 2
                        cy = (rect.top + rect.bottom) // 2

                        self._show_target(cx, cy, assessment_type, color="#4CAF50")
                        elem.click_input()
                        time.sleep(1.0)
                        logger.info(f"Clicked on {assessment_type}!")
                        return True
                except:
                    continue

            logger.warning(f"Assessment '{assessment_type}' not found")
            return False
        except Exception as e:
            logger.error(f"Error finding assessment: {e}")
            return False

    def find_and_click_assessment_description(self, assessment_type=None):
        """Find and click on the assessment description header (e.g., 'Patient Health Questionnaire (PHQ-9)')."""
        if not self.chart_window:
            return False

        # Determine which assessment to look for from the event data
        if not assessment_type and self.current_patient_info:
            assessments = self.current_patient_info.get('assessments', [])
            if assessments:
                assessment_type = assessments[0].get('name', '')

        if not assessment_type:
            logger.warning("No assessment type specified for description")
            return False

        logger.info(f"Looking for assessment description: {assessment_type}")

        try:
            time.sleep(0.5)
            descendants = self.chart_window.descendants()

            # Build specific keywords based on assessment type
            assessment_lower = assessment_type.lower()

            if "phq" in assessment_lower:
                keywords = ["patient health questionnaire"]
            elif "gad" in assessment_lower:
                keywords = ["general anxiety disorder", "generalized anxiety"]
            else:
                keywords = [assessment_lower]

            for elem in descendants:
                try:
                    name = elem.element_info.name or ""
                    ctrl_type = elem.element_info.control_type
                    auto_id = elem.element_info.automation_id or ""
                    name_lower = name.lower()

                    # Look for the description header (Text element without ContentItemTextBlock AutoID)
                    # The header is at the top, list items have AutoID
                    if ctrl_type == "Text" and any(kw in name_lower for kw in keywords):
                        # Prefer the one without AutoID (header) over list items
                        if not auto_id or "ContentItemTextBlock" not in auto_id:
                            rect = elem.rectangle()
                            cx = (rect.left + rect.right) // 2
                            cy = (rect.top + rect.bottom) // 2

                            logger.info(f"Found assessment description: {name}")

                            # Show highlight around the description
                            self._show_highlight_box(
                                rect.left - 5,
                                rect.top - 5,
                                rect.right - rect.left + 10,
                                rect.bottom - rect.top + 10,
                                f"{assessment_type} Description",
                                color="#4CAF50",
                                duration=2.0
                            )

                            # Click on it
                            elem.click_input()
                            time.sleep(1.0)
                            logger.info(f"Clicked on assessment description!")
                            return True
                except:
                    continue

            logger.warning(f"Assessment description for '{assessment_type}' not found")
            return False
        except Exception as e:
            logger.error(f"Error finding assessment description: {e}")
            return False

    def _show_highlight_box(self, x, y, width, height, label, color="#4CAF50", duration=2.0):
        """Show a highlight box around an element."""
        try:
            import tkinter as tk
            import threading

            def show():
                root = tk.Tk()
                root.withdraw()

                # Main highlight box
                win = tk.Toplevel(root)
                win.overrideredirect(True)
                win.attributes("-topmost", True)
                win.attributes("-alpha", 0.8)
                win.geometry(f"{width}x{height}+{x}+{y}")

                canvas = tk.Canvas(win, width=width, height=height, highlightthickness=0, bg=color)
                canvas.pack()

                # Draw border only (hollow rectangle)
                border = 4
                canvas.create_rectangle(border, border, width - border, height - border, fill="black", outline="")

                try:
                    win.attributes("-transparentcolor", "black")
                except:
                    pass

                # Label above
                lbl_win = tk.Toplevel(root)
                lbl_win.overrideredirect(True)
                lbl_win.attributes("-topmost", True)
                lbl = tk.Label(lbl_win, text=f">>> {label}", font=("Segoe UI", 11, "bold"),
                              bg=color, fg="white", padx=10, pady=5)
                lbl.pack()
                lbl_win.geometry(f"+{x}+{y - 35}")

                root.after(int(duration * 1000), root.quit)
                root.mainloop()

                try:
                    lbl_win.destroy()
                    win.destroy()
                    root.destroy()
                except:
                    pass

            thread = threading.Thread(target=show, daemon=True)
            thread.start()
            time.sleep(0.3)
        except Exception as e:
            logger.debug(f"Highlight box error: {e}")

    def _show_target(self, x, y, label, color="#FF6B6B", duration=1.0):
        """Show a visual target overlay."""
        try:
            import tkinter as tk
            import threading

            def show():
                root = tk.Tk()
                root.withdraw()

                win = tk.Toplevel(root)
                win.overrideredirect(True)
                win.attributes("-topmost", True)
                win.attributes("-alpha", 0.8)

                width, height = 200, 40
                win.geometry(f"{width}x{height}+{x - width//2}+{y - height//2}")

                canvas = tk.Canvas(win, width=width, height=height, highlightthickness=0, bg=color)
                canvas.pack()
                canvas.create_rectangle(3, 3, width-3, height-3, fill="black", outline="")

                try:
                    win.attributes("-transparentcolor", "black")
                except:
                    pass

                # Label
                lbl_win = tk.Toplevel(root)
                lbl_win.overrideredirect(True)
                lbl_win.attributes("-topmost", True)
                lbl = tk.Label(lbl_win, text=f">>> {label}", font=("Segoe UI", 10, "bold"),
                              bg=color, fg="white", padx=8, pady=4)
                lbl.pack()
                lbl_win.geometry(f"+{x - 50}+{y - 50}")

                root.after(int(duration * 1000), root.quit)
                root.mainloop()

                try:
                    lbl_win.destroy()
                    win.destroy()
                    root.destroy()
                except:
                    pass

            thread = threading.Thread(target=show, daemon=True)
            thread.start()
            time.sleep(0.3)
        except Exception as e:
            logger.debug(f"Overlay error: {e}")

    def inspect_ui_elements(self):
        """Inspect and print all UI elements (for debugging)."""
        if not self.find_experity_window():
            return

        print("\n" + "=" * 80)
        print("INSPECTING EXPERITY UI ELEMENTS")
        print("=" * 80 + "\n")

        try:
            descendants = self.experity_window.descendants()
            print(f"Total elements: {len(descendants)}\n")

            # Filter for interesting elements
            keywords = ["procedure", "supply", "common", "behavioral", "health",
                       "eval", "phq", "gad", "assessment", "patient"]

            matches = []
            for elem in descendants:
                try:
                    name = elem.element_info.name or ""
                    ctrl_type = elem.element_info.control_type
                    auto_id = elem.element_info.automation_id or ""

                    name_lower = name.lower()
                    auto_id_lower = auto_id.lower()

                    if any(kw in name_lower or kw in auto_id_lower for kw in keywords):
                        matches.append({
                            'name': name,
                            'control_type': ctrl_type,
                            'automation_id': auto_id,
                            'element': elem
                        })
                except:
                    continue

            print(f"Found {len(matches)} relevant elements:\n")
            for i, m in enumerate(matches):
                print(f"[{i}] Type: {m['control_type']}")
                print(f"    Name: {m['name'][:80]}..." if len(m['name']) > 80 else f"    Name: {m['name']}")
                if m['automation_id']:
                    print(f"    AutoID: {m['automation_id']}")
                print()

            return matches
        except Exception as e:
            print(f"Error: {e}")
            return []

    def run_full_test(self, patient_name=None):
        """Run the full outbound test flow."""
        print("\n" + "=" * 60)
        print("MHT OUTBOUND FLOW TEST")
        print("=" * 60 + "\n")

        # Step 1: Find Experity
        print("Step 1: Finding Experity window...")
        if not self.find_experity_window():
            print("FAILED: Experity window not found")
            return False
        print("OK: Found Experity\n")

        # Step 2: Get patient name and assessment data
        patient_info = None
        if not patient_name:
            print("Step 2: Querying for completed outbound events (status=100)...")
            events = self.get_completed_outbound_events()

            if events:
                event = events[0]
                patient_info = self.get_patient_info_from_event(event)
                patient_name = patient_info.get('patient_name')
                print(f"OK: Found outbound event ID {event['id']} for patient: {patient_name}")

                # Show assessment data
                if patient_info.get('assessments'):
                    print("\nAssessment data to enter:")
                    for assess in patient_info['assessments']:
                        print(f"  - {assess['name']}: Score {assess['total_score']} ({assess['severity']})")
                        if assess.get('items'):
                            for item in assess['items'][:3]:  # Show first 3 items
                                print(f"      {item['name']}: {item['score']}")
                            if len(assess['items']) > 3:
                                print(f"      ... and {len(assess['items']) - 3} more items")
                print()

                # Store for later use
                self.current_patient_info = patient_info
            else:
                print("WARNING: No completed outbound events found")
                print("Using test search - looking for any patient in Roomed section...\n")

        if not patient_name:
            print("FAILED: No patient name to test with")
            print("Use --patient 'NAME' to specify a patient")
            return False

        # Step 3: Find patient
        print(f"Step 3: Finding patient '{patient_name}' in Roomed Patients...")
        patient_row = self.find_patient_in_roomed(patient_name)
        if not patient_row:
            print(f"FAILED: Could not find patient '{patient_name}'")
            return False
        print("OK: Found patient\n")

        # Step 4: Click chart
        print("Step 4: Clicking chart icon...")
        if not self.click_chart_icon(patient_row):
            print("FAILED: Could not click chart icon")
            return False
        print("OK: Clicked chart\n")

        # Step 5: Wait for chart window
        print("Step 5: Waiting for chart window...")
        if not self.wait_for_chart_window(patient_name):
            print("FAILED: Chart window did not open")
            return False
        print("OK: Chart window opened\n")

        # Step 6: Navigate to Procedures/Supplies
        print("Step 6: Navigating to Procedures/Supplies...")
        if not self.navigate_to_procedures_supplies():
            print("WARNING: Could not find Procedures/Supplies - trying to continue\n")
        else:
            print("OK: Clicked Procedures/Supplies\n")

        # Step 7: Navigate to SIC COMMON ORDERS
        print("Step 7: Navigating to SIC COMMON ORDERS...")
        if not self.navigate_to_sic_common():
            print("WARNING: Could not find SIC COMMON ORDERS - trying to continue\n")
        else:
            print("OK: Clicked SIC COMMON ORDERS\n")

        # Step 8: Navigate to Behavioral Health Evals
        print("Step 8: Navigating to Behavioral Health Evals...")
        if not self.navigate_to_behavioral_health_evals():
            print("WARNING: Could not find Behavioral Health Evals - trying to continue\n")
        else:
            print("OK: Clicked Behavioral Health Evals\n")

        # Step 9: Find and click the specific assessment from the event
        assessment_name = "Unknown"
        if self.current_patient_info and self.current_patient_info.get('assessments'):
            assessment_name = self.current_patient_info['assessments'][0].get('name', 'Unknown')
        print(f"Step 9: Finding and clicking assessment: {assessment_name}...")
        if not self.find_and_click_assessment():
            print(f"WARNING: Could not find assessment '{assessment_name}'")
        else:
            print(f"OK: Clicked {assessment_name}\n")

        # Step 10: Find and click the assessment description header
        print(f"Step 10: Finding and clicking assessment description...")
        if not self.find_and_click_assessment_description():
            print(f"WARNING: Could not find assessment description")
        else:
            print(f"OK: Clicked assessment description\n")

        print("=" * 60)
        print("TEST COMPLETE")
        print("=" * 60)
        return True


def main():
    parser = argparse.ArgumentParser(description="Test MHT outbound assessment flow")
    parser.add_argument("--inspect", action="store_true", help="Just inspect UI elements")
    parser.add_argument("--patient", type=str, help="Patient name to test with")
    args = parser.parse_args()

    tester = OutboundFlowTester()

    if args.inspect:
        tester.inspect_ui_elements()
    else:
        tester.run_full_test(patient_name=args.patient)


if __name__ == "__main__":
    main()
