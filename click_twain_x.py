"""Click the X button on the Dynamic Web TWAIN popup."""
from pywinauto import Desktop
import pyautogui

desktop = Desktop(backend='uia')

# Find the unnamed button at top of popup (X button)
print('Looking for X button on TWAIN popup...')
for w in desktop.windows():
    try:
        for elem in w.descendants():
            try:
                name = elem.element_info.name or ''
                ctrl = elem.element_info.control_type
                rect = elem.rectangle()
                # Look for unnamed button around X=920, Y=638 (top-right of popup)
                if ctrl == 'Button' and name == '' and rect.left > 900 and rect.top > 620 and rect.top < 660:
                    print(f'Found X button at ({rect.left}, {rect.top})')
                    pyautogui.click((rect.left + rect.right) // 2, (rect.top + rect.bottom) // 2)
                    print('Clicked X!')
                    break
            except:
                pass
    except:
        pass
