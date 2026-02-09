"""Find the X button on the TWAIN popup."""
from pywinauto import Desktop
import pyautogui

desktop = Desktop(backend='uia')

# TWAIN text is at (842, 700) - find clickable elements above it (header area)
print('Looking for clickable elements in popup header area...')

for w in desktop.windows():
    try:
        title = w.window_text()
        if 'MCFARLAND' in title.upper():
            print(f'Window: {title}')
            for elem in w.descendants():
                try:
                    ctrl = elem.element_info.control_type
                    name = elem.element_info.name or ''
                    rect = elem.rectangle()
                    # Look for elements above the popup text (Y < 700) and in the right area (X > 800)
                    # X button is usually top-right, so look for elements with Y around 620-680, X around 1050-1150
                    if ctrl in ['Button', 'Image', 'Hyperlink'] and rect.top > 600 and rect.top < 720 and rect.left > 800:
                        print(f'  [{ctrl}] "{name}" at ({rect.left},{rect.top}) size={rect.width()}x{rect.height()}')
                except:
                    pass
    except:
        pass

print('\n\nTrying to click at estimated X button position...')
# The popup is likely around 400px wide, centered around X=1000
# If popup header is around Y=640, X button would be at top-right: approx (1180, 660)
# But let me try clicking where the earlier empty buttons were found
pyautogui.click(920, 650)
print('Clicked at (920, 650)')
