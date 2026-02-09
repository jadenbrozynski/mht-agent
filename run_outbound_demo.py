"""
MHT Outbound Flow - Full Demo (FAST)
Runs the complete outbound assessment entry flow.
"""

from pywinauto import Application, findwindows
import pyautogui
import sqlite3
import json
import time
from pathlib import Path

DB_PATH = Path(__file__).parent / 'output' / 'mht_data.db'

def run_demo():
    print('=' * 60)
    print('MHT OUTBOUND FLOW - FULL DEMO')
    print('=' * 60)

    # ===== STEP 1: Connect to Tracking Board =====
    print('\n[Step 1] Connecting to Tracking Board...')
    elements = findwindows.find_elements(title_re='.*Tracking Board.*', backend='uia')
    target_handle = None
    for elem in elements:
        if 'Setup' not in elem.name:
            target_handle = elem.handle
            break

    app = Application(backend='uia').connect(handle=target_handle, timeout=3)
    tb_win = app.window(handle=target_handle)
    print(f'Connected: {tb_win.window_text()}')

    # ===== STEP 2: Get outbound event data =====
    print('\n[Step 2] Getting outbound event data...')
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute('''
        SELECT id, raw_data FROM common_event
        WHERE direction = 'O' AND status = 100
        ORDER BY id DESC LIMIT 1
    ''')
    row = cursor.fetchone()
    raw = json.loads(row['raw_data'])
    data = raw.get('data', raw)
    patient = data.get('patient', {})
    patient_name = f"{patient.get('patient_last_name')}, {patient.get('patient_first_name')}"
    assessments = data.get('assessment', [])
    if not isinstance(assessments, list):
        assessments = [assessments]
    assess = assessments[0]
    scores = [int(float(item.get('assessment_item_score', 0))) for item in assess.get('assessment_items', [])]
    total_score = int(assess.get('total_score_value', sum(scores)))
    assess_name = assess.get('assessment_name', '')

    print(f'Patient: {patient_name}')
    print(f'Assessment: {assess_name}')
    print(f'Scores: {scores}')
    print(f'Total: {total_score}')
    conn.close()

    patient_upper = patient_name.upper()

    # ===== STEP 3: Find patient in Roomed Patients =====
    print(f'\n[Step 3] Finding patient in Roomed Patients...')
    roomed_group = tb_win.child_window(title_re='.*Roomed Patients.*', control_type='Group')
    roomed_items = roomed_group.descendants(control_type='DataItem')
    patient_row = None

    for di in roomed_items:
        try:
            di_text = di.window_text()
            if patient_upper in di_text.upper():
                patient_row = di
                break
            for child in di.descendants():
                if patient_upper in child.window_text().upper():
                    patient_row = di
                    break
            if patient_row:
                break
        except:
            continue

    if not patient_row:
        print(f'ERROR: Patient {patient_name} not found!')
        return
    print('Found patient')

    # ===== STEP 4: Click chart icon =====
    print('\n[Step 4] Clicking chart icon...')
    images = patient_row.descendants(control_type='Image')
    chart_icon = images[0]
    rect = chart_icon.rectangle()
    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
    time.sleep(0.4)
    print('Clicked chart')

    # ===== STEP 5: Wait for chart window =====
    print('\n[Step 5] Waiting for chart window...')
    name_part = patient_name.split(',')[0]
    chart_app = Application(backend='uia').connect(title_re=f'.*{name_part}.*', timeout=5)
    chart_win = chart_app.window(title_re=f'.*{name_part}.*')
    print('Chart opened')

    # ===== STEP 6: Navigate to Procedures/Supplies =====
    print('\n[Step 6] Navigating to Procedures/Supplies...')
    time.sleep(0.2)
    for elem in chart_win.descendants():
        try:
            if elem.element_info.control_type == 'TabItem' and 'procedures' in (elem.element_info.name or '').lower():
                elem.click_input()
                print('Clicked Procedures/Supplies')
                break
        except:
            pass
    time.sleep(0.2)

    # ===== STEP 7: Navigate to SIC COMMON ORDERS =====
    print('\n[Step 7] Navigating to SIC COMMON ORDERS...')
    for elem in chart_win.descendants():
        try:
            name = elem.element_info.name or ''
            if 'sic common' in name.lower():
                elem.click_input()
                print('Clicked SIC COMMON ORDERS')
                break
        except:
            pass
    time.sleep(0.2)

    # ===== STEP 8: Click on assessment in the list (skip Behavioral Health Evals) =====
    print(f'\n[Step 8] Clicking on {assess_name} in the list...')
    keywords = ['phq-9', 'phq9', 'patient health'] if 'phq' in assess_name.lower() else ['gad-7', 'gad7', 'general anxiety']

    for elem in chart_win.descendants():
        try:
            name = elem.element_info.name or ''
            ctrl_type = elem.element_info.control_type
            auto_id = elem.element_info.automation_id or ''

            if ctrl_type == 'Text' and auto_id == 'ContentItemTextBlock' and any(kw in name.lower() for kw in keywords):
                rect = elem.rectangle()
                print(f'Found: {name}')
                elem.click_input()
                print(f'Clicked {assess_name}')
                break
        except:
            pass
    time.sleep(0.2)

    # ===== STEP 9: Click on assessment description header =====
    print('\n[Step 9] Clicking on assessment description header...')
    for elem in chart_win.descendants():
        try:
            name = elem.element_info.name or ''
            ctrl_type = elem.element_info.control_type
            auto_id = elem.element_info.automation_id or ''

            if ctrl_type == 'Text' and auto_id != 'ContentItemTextBlock' and any(kw in name.lower() for kw in keywords):
                rect = elem.rectangle()
                print(f'Found header: {name}')
                elem.click_input()
                print('Clicked description header')
                break
        except:
            pass
    time.sleep(0.2)

    # ===== STEP 10: Fill out the form =====
    print('\n[Step 10] Filling out assessment form...')

    all_radios = []
    for elem in chart_win.descendants():
        try:
            if elem.element_info.control_type == 'RadioButton':
                rect = elem.rectangle()
                all_radios.append({'name': elem.element_info.name, 'elem': elem, 'y': rect.top, 'x': rect.left})
        except:
            pass

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

    print(f'Found {len(rows)} rows')

    # Click patient (row 0)
    btn = rows[0][0]
    rect = btn['elem'].rectangle()
    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

    # Fill Q1-Q9 based on scores
    for i, score in enumerate(scores):
        row_idx = i + 1
        if row_idx < len(rows) and score < len(rows[row_idx]):
            btn = rows[row_idx][score]
            rect = btn['elem'].rectangle()
            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

    # Click Q10 (difficulty - first option)
    if len(rows) > 10:
        btn = rows[10][0]
        rect = btn['elem'].rectangle()
        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)

    print('Filled all questions')

    # ===== STEP 11: Scroll down and enter total score =====
    print('\n[Step 11] Entering total score...')
    pyautogui.click(1200, 800)
    pyautogui.scroll(-500)  # Instant scroll to bottom
    time.sleep(0.2)

    for elem in chart_win.descendants():
        try:
            if elem.element_info.control_type == 'Edit' and 'total score' in (elem.element_info.name or '').lower():
                rect = elem.rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                pyautogui.hotkey('ctrl', 'a')
                pyautogui.typewrite(str(total_score))
                pyautogui.press('enter')
                print(f'Entered total: {total_score}')
                break
        except:
            pass

    # ===== STEP 12: Scroll to top and click Cancel =====
    print('\n[Step 12] Clicking Cancel...')
    pyautogui.scroll(500)  # Instant scroll to top
    time.sleep(0.2)

    for elem in chart_win.descendants():
        try:
            if elem.element_info.control_type == 'Button' and elem.element_info.automation_id == 'cmdCancel2':
                rect = elem.rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                print('Clicked Cancel')
                break
        except:
            pass
    time.sleep(0.3)

    # ===== STEP 13: Delete assessment =====
    print('\n[Step 13] Deleting assessment...')
    phq_y = None
    for elem in chart_win.descendants():
        try:
            name = elem.element_info.name or ''
            if any(kw in name.lower() for kw in keywords):
                phq_y = elem.rectangle().top
                break
        except:
            pass

    for elem in chart_win.descendants():
        try:
            if elem.element_info.control_type == 'Button' and 'delete' in (elem.element_info.name or '').lower():
                rect = elem.rectangle()
                if phq_y and abs(rect.top - phq_y) < 30:
                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                    print('Clicked Delete')
                    break
        except:
            pass

    time.sleep(0.3)

    for elem in chart_win.descendants():
        try:
            if elem.element_info.control_type == 'Button' and (elem.element_info.name or '').lower() == 'ok':
                rect = elem.rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                print('Clicked OK')
                break
        except:
            pass

    time.sleep(0.3)

    # ===== STEP 14: Close chart =====
    print('\n[Step 14] Closing chart...')

    for elem in chart_win.descendants():
        try:
            if elem.element_info.control_type == 'Button' and elem.element_info.automation_id == 'SaveButton':
                elem.click_input()
                print('Clicked Close')
                break
        except:
            pass

    time.sleep(0.3)

    for elem in chart_win.descendants():
        try:
            if elem.element_info.control_type == 'Button' and elem.element_info.automation_id == 'btnSave':
                rect = elem.rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                print('Clicked Close Chart')
                break
        except:
            pass

    time.sleep(0.2)

    # ===== STEP 15: Click on patient name in Roomed Patients =====
    print('\n[Step 15] Clicking on patient name...')

    elements = findwindows.find_elements(title_re='.*Tracking Board.*', backend='uia')
    target_handle = None
    for elem in elements:
        if 'Setup' not in elem.name:
            target_handle = elem.handle
            break

    tb_app = Application(backend='uia').connect(handle=target_handle, timeout=3)
    tb_win = tb_app.window(handle=target_handle)

    roomed_group = tb_win.child_window(title_re='.*Roomed Patients.*', control_type='Group')
    roomed_items = roomed_group.descendants(control_type='DataItem')

    for di in roomed_items:
        try:
            for child in di.descendants():
                child_text = child.window_text()
                if patient_upper in child_text.upper() and ',' in child_text:
                    rect = child.rectangle()
                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                    print(f'Clicked on {child_text}')
                    break
        except:
            continue

    time.sleep(0.5)

    # ===== STEP 16: Click Documents tab =====
    print('\n[Step 16] Clicking Documents tab...')

    demo_app = Application(backend='uia').connect(title_re=f'.*{name_part}.*', timeout=5)
    demo_win = demo_app.window(title_re=f'.*{name_part}.*')
    print(f'Connected to: {demo_win.window_text()[:50]}')

    for elem in demo_win.descendants():
        try:
            ctrl_type = elem.element_info.control_type
            name = elem.element_info.name or ''
            if ctrl_type == 'TabItem' and 'documents' in name.lower():
                elem.click_input()
                print('Clicked Documents tab')
                break
        except:
            pass

    time.sleep(2.0)  # Wait for Documents tab to fully load

    # ===== STEP 17: Click Scan/Upload button =====
    print('\n[Step 17] Clicking Scan/Upload button...')

    # Click Scan/Upload at known position
    pyautogui.click(2240, 243)
    print('Clicked Scan/Upload at (2240, 243)')

    # Re-fetch window for next steps
    from pywinauto import Desktop
    desktop = Desktop(backend='uia')
    demo_win = None
    for w in desktop.windows():
        try:
            if name_part.upper() in w.window_text().upper():
                demo_win = w
                break
        except:
            pass

    # ===== STEP 18: Wait and close TWAIN popup =====
    print('\n[Step 18] Waiting for TWAIN popup...')
    time.sleep(5)

    for elem in demo_win.descendants():
        try:
            name = elem.element_info.name or ''
            if 'dynamic web twain' in name.lower():
                rect = elem.rectangle()
                twain_x = rect.right + 29
                twain_y = rect.top - 67
                pyautogui.click(twain_x, twain_y)
                print(f'Clicked TWAIN X')
                break
        except:
            pass

    time.sleep(0.3)

    # ===== STEP 19: Select Description dropdown =====
    print('\n[Step 19] Selecting Mental/Behavioral Health Assessment...')

    for elem in demo_win.descendants():
        try:
            ctrl_type = elem.element_info.control_type
            auto_id = elem.element_info.automation_id or ''
            if ctrl_type == 'ComboBox' and 'mat-select' in auto_id:
                rect = elem.rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                print('Clicked dropdown')
                break
        except:
            pass

    time.sleep(0.3)

    pyautogui.moveTo(1117, 850)
    pyautogui.scroll(-1200)  # 60% down
    time.sleep(0.2)

    for elem in demo_win.descendants():
        try:
            ctrl_type = elem.element_info.control_type
            name = elem.element_info.name or ''
            if ctrl_type == 'ListItem' and 'mental' in name.lower() and 'health' in name.lower():
                rect = elem.rectangle()
                if rect.top > 100:
                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                    print(f'Selected: {name}')
                    break
        except:
            pass

    time.sleep(0.3)

    # ===== STEP 20: Click Cancel =====
    print('\n[Step 20] Clicking Cancel...')
    for elem in demo_win.descendants():
        try:
            ctrl_type = elem.element_info.control_type
            name = elem.element_info.name or ''
            if ctrl_type == 'Button' and name.lower() == 'cancel':
                rect = elem.rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                print('Clicked Cancel')
                break
        except:
            pass

    time.sleep(0.3)

    # ===== STEP 21: Click OK on unsaved changes =====
    print('\n[Step 21] Clicking OK...')
    for elem in demo_win.descendants():
        try:
            ctrl_type = elem.element_info.control_type
            name = elem.element_info.name or ''
            if ctrl_type == 'Button' and name.lower() == 'ok':
                rect = elem.rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                print('Clicked OK')
                break
        except:
            pass

    time.sleep(0.2)

    # ===== STEP 22: Close Demographics tab =====
    print('\n[Step 22] Closing Demographics tab...')
    close_buttons = []
    for elem in demo_win.descendants():
        try:
            ctrl_type = elem.element_info.control_type
            auto_id = elem.element_info.automation_id or ''
            if ctrl_type == 'Button' and auto_id == 'PART_CloseButton':
                rect = elem.rectangle()
                if rect.top > 100 and rect.top < 200:
                    close_buttons.append(rect)
        except:
            pass

    if len(close_buttons) >= 2:
        rect = close_buttons[1]
        pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
        print('Closed Demographics tab')

    print('\n' + '=' * 60)
    print('DEMO COMPLETE!')
    print('=' * 60)


if __name__ == '__main__':
    run_demo()
