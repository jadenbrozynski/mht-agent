"""Desktop automation for Experity application."""

from __future__ import annotations

import subprocess
import time
import logging
from pathlib import Path
from typing import Optional, Tuple
import pyautogui
import pygetwindow as gw
from PIL import Image

logger = logging.getLogger("mhtagentic.desktop")

# Configure pyautogui safety
pyautogui.FAILSAFE = True  # Move mouse to corner to abort
pyautogui.PAUSE = 0.1  # Small pause between actions


class DesktopAutomation:
    """
    Desktop automation controller for Experity or browser-based testing.

    Handles launching the app, finding windows, and interacting with UI elements.
    """

    EXPERITY_SHORTCUT = r"C:\Users\Public\Desktop\Experity EMR - PROD.lnk"
    WINDOW_TITLE_PATTERN = "Experity"
    BROWSER_PATTERNS = ["Chrome", "Edge", "Firefox", "Brave", "Opera", "Login Test"]

    def __init__(self, shortcut_path: Optional[str] = None, window_pattern: Optional[str] = None):
        """
        Initialize desktop automation.

        Args:
            shortcut_path: Path to Experity shortcut (uses default if not provided)
            window_pattern: Custom window title pattern to search for
        """
        self.shortcut_path = shortcut_path or self.EXPERITY_SHORTCUT
        self.custom_pattern = window_pattern
        self.window: Optional[gw.Window] = None
        self.screenshot_dir = Path("output/screenshots")
        self.screenshot_dir.mkdir(parents=True, exist_ok=True)

    def launch_experity(self, wait_seconds: int = 10) -> bool:
        """
        Launch Experity application via shortcut.

        Args:
            wait_seconds: Seconds to wait for app to load

        Returns:
            True if app launched and window found
        """
        if not Path(self.shortcut_path).exists():
            logger.error(f"Shortcut not found: {self.shortcut_path}")
            return False

        logger.info(f"Launching Experity from: {self.shortcut_path}")

        try:
            # Launch via shortcut
            subprocess.Popen(
                ["cmd", "/c", "start", "", self.shortcut_path],
                shell=True
            )

            # Wait for window to appear
            logger.info(f"Waiting {wait_seconds}s for Experity to load...")

            for i in range(wait_seconds * 2):
                time.sleep(0.5)
                if self.find_experity_window():
                    logger.info("Experity window found!")
                    return True

            logger.warning("Experity window not found after timeout")
            return False

        except Exception as e:
            logger.error(f"Failed to launch Experity: {e}")
            return False

    def find_experity_window(self) -> bool:
        """
        Find and focus the Experity window (does NOT fall back to browser).

        Returns:
            True if Experity window found
        """
        try:
            # First try custom pattern if set
            if self.custom_pattern:
                windows = gw.getWindowsWithTitle(self.custom_pattern)
                if windows:
                    self.window = windows[0]
                    logger.info(f"Found custom window: {self.window.title}")
                    return True

            # Try Experity - must contain "Experity" in title
            all_windows = gw.getAllWindows()
            for w in all_windows:
                if "experity" in w.title.lower():
                    self.window = w
                    logger.info(f"Found Experity window: {self.window.title}")
                    return True

            return False

        except Exception as e:
            logger.error(f"Error finding window: {e}")
            return False

    def find_browser_window(self) -> bool:
        """
        Find a browser window.

        Returns:
            True if browser window found
        """
        try:
            all_browser_windows = []
            for pattern in self.BROWSER_PATTERNS:
                windows = gw.getWindowsWithTitle(pattern)
                all_browser_windows.extend(windows)

            if all_browser_windows:
                # Prefer windows with login-related titles
                for w in all_browser_windows:
                    title_lower = w.title.lower()
                    if any(k in title_lower for k in ["login test", "mhtagentic", "login", "sign in"]):
                        self.window = w
                        logger.info(f"Found login browser window")
                        return True
                # Otherwise use first match
                self.window = all_browser_windows[0]
                logger.info(f"Found browser window")
                return True
            return False
        except Exception as e:
            logger.error(f"Error finding browser: {e}")
            return False

    def launch_test_page(self, html_path: str, wait_seconds: int = 5) -> bool:
        """
        Launch a test HTML page in the default browser.

        Args:
            html_path: Path to HTML file
            wait_seconds: Seconds to wait for browser

        Returns:
            True if browser window found after launch
        """
        import webbrowser

        logger.info(f"Opening test page: {html_path}")
        webbrowser.open(f"file://{html_path}")

        # Wait for browser window
        for i in range(wait_seconds * 2):
            time.sleep(0.5)
            if self.find_browser_window():
                return True

        return False

    def focus_window(self) -> bool:
        """Bring Experity window to foreground."""
        if not self.window:
            if not self.find_experity_window():
                return False

        try:
            self.window.activate()
            time.sleep(0.3)
            return True
        except Exception as e:
            logger.error(f"Failed to focus window: {e}")
            return False

    def get_window_region(self) -> Optional[Tuple[int, int, int, int]]:
        """
        Get the window region (left, top, width, height).

        Returns:
            Tuple of (left, top, width, height) or None
        """
        if not self.window:
            if not self.find_experity_window():
                return None

        try:
            return (
                self.window.left,
                self.window.top,
                self.window.width,
                self.window.height
            )
        except Exception:
            return None

    def take_screenshot(self, filename: str = None) -> Optional[Path]:
        """
        Take a screenshot of the Experity window.

        Args:
            filename: Optional filename (auto-generated if not provided)

        Returns:
            Path to saved screenshot
        """
        if not self.focus_window():
            logger.error("Cannot take screenshot - window not found")
            return None

        region = self.get_window_region()
        if not region:
            # Full screen fallback
            screenshot = pyautogui.screenshot()
        else:
            screenshot = pyautogui.screenshot(region=region)

        if filename is None:
            filename = f"experity_{int(time.time())}.png"

        filepath = self.screenshot_dir / filename
        screenshot.save(filepath)
        logger.info(f"Screenshot saved: {filepath}")

        return filepath

    def click_at(self, x: int, y: int, clicks: int = 1) -> None:
        """Click at screen coordinates."""
        pyautogui.click(x, y, clicks=clicks)
        logger.debug(f"Clicked at ({x}, {y})")

    def type_text(self, text: str, interval: float = 0.02) -> None:
        """Type text with optional interval between keystrokes."""
        pyautogui.write(text, interval=interval)
        logger.debug(f"Typed text (length: {len(text)})")

    def press_key(self, key: str) -> None:
        """Press a single key."""
        pyautogui.press(key)
        logger.debug(f"Pressed key: {key}")

    def hotkey(self, *keys: str) -> None:
        """Press a key combination."""
        pyautogui.hotkey(*keys)
        logger.debug(f"Hotkey: {'+'.join(keys)}")

    def move_to(self, x: int, y: int, duration: float = 0.2) -> None:
        """Move mouse to coordinates."""
        pyautogui.moveTo(x, y, duration=duration)

    def get_mouse_position(self) -> Tuple[int, int]:
        """Get current mouse position."""
        return pyautogui.position()

    def wait_for_image(
        self,
        image_path: str,
        timeout: int = 10,
        confidence: float = 0.9
    ) -> Optional[Tuple[int, int]]:
        """
        Wait for an image to appear on screen.

        Args:
            image_path: Path to reference image
            timeout: Seconds to wait
            confidence: Match confidence (0-1)

        Returns:
            Center coordinates of found image, or None
        """
        start_time = time.time()

        while time.time() - start_time < timeout:
            try:
                location = pyautogui.locateCenterOnScreen(
                    image_path,
                    confidence=confidence
                )
                if location:
                    return location
            except Exception:
                pass
            time.sleep(0.5)

        return None

    def input_field(
        self,
        x: int,
        y: int,
        text: str,
        clear_first: bool = True
    ) -> None:
        """
        Click on a field and input text.

        Args:
            x, y: Field coordinates
            text: Text to input
            clear_first: Whether to clear existing text
        """
        self.click_at(x, y)
        time.sleep(0.1)

        if clear_first:
            self.hotkey("ctrl", "a")
            time.sleep(0.05)

        self.type_text(text)
