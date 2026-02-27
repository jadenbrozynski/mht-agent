"""
Win32 PrintWindow screenshot capture for mstsc (RDP) windows.

Uses the PrintWindow API via ctypes to capture window contents without
activating or focusing the window, so the running bot is not disturbed.
"""

import ctypes
import ctypes.wintypes
import io
import logging
from typing import Optional

from PIL import Image

logger = logging.getLogger("mht_dashboard.screenshot")

# Win32 constants
SRCCOPY = 0x00CC0020
PW_RENDERFULLCONTENT = 0x00000002
DIB_RGB_COLORS = 0

# Win32 API
user32 = ctypes.windll.user32
gdi32 = ctypes.windll.gdi32

# Enable DPI awareness so GetWindowRect returns real pixel coordinates.
# Without this, 150% scaling causes 960x540 to report as 640x360,
# and the screenshot only captures the top-left corner.
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor DPI aware
except Exception:
    user32.SetProcessDPIAware()  # Fallback for older Windows


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", ctypes.wintypes.DWORD),
        ("biWidth", ctypes.wintypes.LONG),
        ("biHeight", ctypes.wintypes.LONG),
        ("biPlanes", ctypes.wintypes.WORD),
        ("biBitCount", ctypes.wintypes.WORD),
        ("biCompression", ctypes.wintypes.DWORD),
        ("biSizeImage", ctypes.wintypes.DWORD),
        ("biXPelsPerMeter", ctypes.wintypes.LONG),
        ("biYPelsPerMeter", ctypes.wintypes.LONG),
        ("biClrUsed", ctypes.wintypes.DWORD),
        ("biClrImportant", ctypes.wintypes.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [
        ("bmiHeader", BITMAPINFOHEADER),
        ("bmiColors", ctypes.wintypes.DWORD * 3),
    ]


def capture_window(hwnd: int) -> Optional[Image.Image]:
    """
    Capture a window screenshot using PrintWindow (no focus change).

    Args:
        hwnd: Window handle (HWND)

    Returns:
        PIL Image or None on failure
    """
    try:
        # Get full window dimensions (including title bar and borders)
        # This matches what PW_RENDERFULLCONTENT renders
        rect = ctypes.wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        width = rect.right - rect.left
        height = rect.bottom - rect.top

        if width <= 0 or height <= 0:
            return None

        # Use screen DC for compatible bitmap (not window DC)
        screen_dc = user32.GetDC(0)
        if not screen_dc:
            return None

        cdc = gdi32.CreateCompatibleDC(screen_dc)
        bitmap = gdi32.CreateCompatibleBitmap(screen_dc, width, height)
        gdi32.SelectObject(cdc, bitmap)

        # PrintWindow captures the full window without activating it
        result = user32.PrintWindow(hwnd, cdc, PW_RENDERFULLCONTENT)

        if not result:
            # Fallback to BitBlt from screen DC
            wdc = user32.GetDC(hwnd)
            gdi32.BitBlt(cdc, 0, 0, width, height, wdc, 0, 0, SRCCOPY)
            user32.ReleaseDC(hwnd, wdc)

        # Extract pixel data
        bmi = BITMAPINFO()
        bmi.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
        bmi.bmiHeader.biWidth = width
        bmi.bmiHeader.biHeight = -height  # Negative = top-down
        bmi.bmiHeader.biPlanes = 1
        bmi.bmiHeader.biBitCount = 32
        bmi.bmiHeader.biCompression = 0

        buf_size = width * height * 4
        buf = ctypes.create_string_buffer(buf_size)
        gdi32.GetDIBits(cdc, bitmap, 0, height, buf, ctypes.byref(bmi), DIB_RGB_COLORS)

        # Cleanup GDI objects
        gdi32.DeleteObject(bitmap)
        gdi32.DeleteDC(cdc)
        user32.ReleaseDC(0, screen_dc)

        # Convert BGRA buffer to PIL Image
        img = Image.frombuffer("RGBA", (width, height), buf, "raw", "BGRA", 0, 1)
        return img.convert("RGB")

    except Exception as e:
        logger.error(f"Screenshot capture failed for hwnd {hwnd}: {e}")
        return None


def capture_window_jpeg(hwnd: int, quality: int = 70) -> Optional[bytes]:
    """
    Capture a window and return JPEG bytes (for serving via HTTP).

    Args:
        hwnd: Window handle
        quality: JPEG quality (1-100)

    Returns:
        JPEG bytes or None on failure
    """
    img = capture_window(hwnd)
    if img is None:
        return None

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()
