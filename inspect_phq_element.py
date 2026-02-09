"""
Inspect the Patient Health Questionnaire description element using pywinauto.
Run this after clicking on an assessment to find the PHQ description element.
"""

from pywinauto import Desktop, Application
from pywinauto.findwindows import find_elements
import time

def find_experity_window():
    """Find the Experity window."""
    desktop = Desktop(backend="uia")
    windows = desktop.windows()

    for win in windows:
        try:
            title = win.window_text()
            if "experity" in title.lower():
                print(f"Found Experity window: {title}")
                return win
        except:
            continue
    return None

def print_all_children(control, indent=0):
    """Recursively print all children of a control."""
    prefix = "  " * indent
    try:
        ctrl_type = control.element_info.control_type
        name = control.element_info.name or ""
        auto_id = control.element_info.automation_id or ""
        class_name = control.element_info.class_name or ""

        # Look for potential PHQ/assessment description elements
        name_lower = name.lower()
        if any(kw in name_lower for kw in ["patient health", "phq", "questionnaire", "assessment", "depression", "anxiety", "gad"]):
            print(f"{prefix}*** MATCH *** Type: {ctrl_type}, Name: '{name}', AutoID: '{auto_id}', Class: '{class_name}'")
        else:
            print(f"{prefix}Type: {ctrl_type}, Name: '{name[:50]}...'" if len(name) > 50 else f"{prefix}Type: {ctrl_type}, Name: '{name}'")

    except Exception as e:
        print(f"{prefix}Error: {e}")
        return

    try:
        children = control.children()
        for child in children:
            print_all_children(child, indent + 1)
    except:
        pass

def inspect_top_area(window):
    """Inspect the top area of the window where PHQ description appears."""
    print("\n" + "="*80)
    print("INSPECTING TOP AREA FOR PHQ DESCRIPTION ELEMENT")
    print("="*80 + "\n")

    try:
        # Get all descendants and look for PHQ-related elements
        descendants = window.descendants()

        print(f"Total elements found: {len(descendants)}\n")
        print("Looking for PHQ/Assessment related elements...\n")

        matches = []
        for elem in descendants:
            try:
                name = elem.element_info.name or ""
                ctrl_type = elem.element_info.control_type
                auto_id = elem.element_info.automation_id or ""
                class_name = elem.element_info.class_name or ""

                name_lower = name.lower()
                auto_id_lower = auto_id.lower()

                # Keywords to search for
                keywords = ["patient health", "phq", "questionnaire", "assessment",
                           "depression", "anxiety", "gad", "procedure", "supply"]

                if any(kw in name_lower or kw in auto_id_lower for kw in keywords):
                    matches.append({
                        'name': name,
                        'control_type': ctrl_type,
                        'automation_id': auto_id,
                        'class_name': class_name,
                        'element': elem
                    })

            except:
                continue

        print(f"Found {len(matches)} potential matches:\n")
        for i, match in enumerate(matches):
            print(f"[{i}] Control Type: {match['control_type']}")
            print(f"    Name: {match['name']}")
            print(f"    AutomationID: {match['automation_id']}")
            print(f"    ClassName: {match['class_name']}")
            print()

        return matches

    except Exception as e:
        print(f"Error inspecting: {e}")
        return []

def main():
    print("Looking for Experity window...")
    window = find_experity_window()

    if not window:
        print("Experity window not found!")
        print("\nAvailable windows:")
        desktop = Desktop(backend="uia")
        for win in desktop.windows():
            try:
                print(f"  - {win.window_text()}")
            except:
                pass
        return

    print(f"\nFound window, waiting 2 seconds...")
    time.sleep(2)

    matches = inspect_top_area(window)

    if matches:
        print("\n" + "="*80)
        print("COPY THESE IDENTIFIERS FOR YOUR AUTOMATION CODE")
        print("="*80)
        for m in matches:
            if m['automation_id']:
                print(f"  automation_id='{m['automation_id']}'")
            if m['name']:
                print(f"  title='{m['name']}'")

if __name__ == "__main__":
    main()
