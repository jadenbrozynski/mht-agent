"""Run steps 16-23 for patient MCFARLAND, SHELIA"""
from pywinauto import Application, findwindows
import pyautogui
import time

patient_name = 'MCFARLAND, SHELIA'
patient_upper = patient_name.upper()

print('=' * 60)
print('RUNNING STEPS 16-23 FOR', patient_name)
print('=' * 60)

# ===== STEP 16: Click on patient name in Roomed Patients =====
print('\n[Step 16] Clicking on patient name to open details...')

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

clicked = False
for di in roomed_items:
    try:
        for child in di.descendants():
            child_text = child.window_text()
            if patient_upper in child_text.upper() and ',' in child_text:
                rect = child.rectangle()
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                print(f'Clicked on {child_text}')
                clicked = True
                break
        if clicked:
            break
    except:
        continue

time.sleep(1.0)

# ===== STEP 17: Click Documents tab =====
print('\n[Step 17] Clicking Documents tab...')

name_part = patient_name.split(',')[0]
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

time.sleep(1.0)

# ===== STEP 18: Click Scan/Upload button =====
print('\n[Step 18] Clicking Scan/Upload button...')

scan_clicked = False
for elem in demo_win.descendants():
    try:
        ctrl_type = elem.element_info.control_type
        name = elem.element_info.name or ''
        if ctrl_type == 'Button' and 'Scan/Upload' in name:
            rect = elem.rectangle()
            if rect.left > 0:  # Valid position
                pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                print(f'Clicked: {name} at ({rect.left}, {rect.top})')
                scan_clicked = True
                break
    except:
        pass

if not scan_clicked:
    print('ERROR: Scan/Upload button not found!')

# ===== STEP 19: Wait and close TWAIN popup =====
print('\n[Step 19] Waiting 5 seconds then closing TWAIN popup...')
time.sleep(5)

for elem in demo_win.descendants():
    try:
        name = elem.element_info.name or ''
        if 'dynamic web twain' in name.lower():
            rect = elem.rectangle()
            twain_x = rect.right + 29
            twain_y = rect.top - 67
            pyautogui.click(twain_x, twain_y)
            print(f'Clicked TWAIN popup X at ({twain_x}, {twain_y})')
            break
    except:
        pass

time.sleep(0.5)

# ===== STEP 20: Select Description dropdown =====
print('\n[Step 20] Selecting Mental/Behavioral Health Assessment...')

for elem in demo_win.descendants():
    try:
        ctrl_type = elem.element_info.control_type
        auto_id = elem.element_info.automation_id or ''
        if ctrl_type == 'ComboBox' and 'mat-select' in auto_id:
            rect = elem.rectangle()
            pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
            print('Clicked Description dropdown')
            break
    except:
        pass

time.sleep(0.4)

# Scroll 60% down in dropdown
pyautogui.moveTo(1117, 850)
pyautogui.scroll(-1200)
time.sleep(0.3)

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

time.sleep(0.5)

# ===== STEP 21: Click Cancel =====
print('\n[Step 21] Clicking Cancel...')
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

time.sleep(0.5)

# ===== STEP 22: Click OK on unsaved changes modal =====
print('\n[Step 22] Clicking OK on unsaved changes modal...')
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

time.sleep(0.3)

# ===== STEP 23: Close Demographics tab =====
print('\n[Step 23] Closing Demographics tab...')
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
    print('Clicked Demographics close')

print('\n' + '=' * 60)
print('STEPS 16-23 COMPLETE!')
print('=' * 60)
