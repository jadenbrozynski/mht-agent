"""
Macro Recorder - Records mouse clicks and keyboard input for playback.
"""

import json
import time
import threading
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import List, Optional, Callable
import pyautogui
from pynput import mouse, keyboard

MACRO_FILE = Path(__file__).parent.parent.parent / "config" / "login_macro.json"


@dataclass
class MacroAction:
    """A single recorded action."""
    action_type: str  # "click", "type", "key", "wait"
    x: Optional[int] = None
    y: Optional[int] = None
    text: Optional[str] = None
    key: Optional[str] = None
    delay: float = 0.0  # Delay before this action


class MacroRecorder:
    """Records mouse and keyboard actions."""

    def __init__(self, on_action: Optional[Callable[[str], None]] = None):
        self.actions: List[MacroAction] = []
        self.recording = False
        self.last_action_time = 0
        self.on_action = on_action  # Callback for UI updates
        self._mouse_listener = None
        self._keyboard_listener = None
        self._current_text = ""  # Buffer for typed text
        self._text_start_time = 0

    def start_recording(self):
        """Start recording actions."""
        self.actions = []
        self.recording = True
        self.last_action_time = time.time()
        self._current_text = ""

        # Start mouse listener
        self._mouse_listener = mouse.Listener(on_click=self._on_click)
        self._mouse_listener.start()

        # Start keyboard listener
        self._keyboard_listener = keyboard.Listener(
            on_press=self._on_key_press,
            on_release=self._on_key_release
        )
        self._keyboard_listener.start()

        if self.on_action:
            self.on_action("Recording started...")

    def stop_recording(self):
        """Stop recording and save."""
        self.recording = False

        # Flush any pending text
        self._flush_text()

        # Stop listeners
        if self._mouse_listener:
            self._mouse_listener.stop()
        if self._keyboard_listener:
            self._keyboard_listener.stop()

        # Save to file
        self.save()

        if self.on_action:
            self.on_action(f"Recording stopped. {len(self.actions)} actions saved.")

    def _on_click(self, x, y, button, pressed):
        """Handle mouse click."""
        if not self.recording or not pressed:
            return

        # Only record left clicks
        if button != mouse.Button.left:
            return

        # Flush any pending text first
        self._flush_text()

        # Calculate delay
        now = time.time()
        delay = now - self.last_action_time
        self.last_action_time = now

        # Record click
        action = MacroAction(
            action_type="click",
            x=int(x),
            y=int(y),
            delay=round(delay, 2)
        )
        self.actions.append(action)

        if self.on_action:
            self.on_action(f"Click at ({x}, {y})")

    def _on_key_press(self, key):
        """Handle key press."""
        if not self.recording:
            return

        try:
            # Regular character
            char = key.char
            if char:
                if not self._current_text:
                    self._text_start_time = time.time()
                self._current_text += char
        except AttributeError:
            # Special key
            self._flush_text()

            now = time.time()
            delay = now - self.last_action_time
            self.last_action_time = now

            key_name = str(key).replace("Key.", "")

            # Skip modifier keys alone
            if key_name in ["shift", "shift_r", "ctrl_l", "ctrl_r", "alt_l", "alt_r"]:
                return

            action = MacroAction(
                action_type="key",
                key=key_name,
                delay=round(delay, 2)
            )
            self.actions.append(action)

            if self.on_action:
                self.on_action(f"Key: {key_name}")

    def _on_key_release(self, key):
        """Handle key release."""
        pass

    def _flush_text(self):
        """Flush buffered text as a type action."""
        if self._current_text:
            delay = self._text_start_time - self.last_action_time if self._text_start_time > self.last_action_time else 0
            self.last_action_time = time.time()

            action = MacroAction(
                action_type="type",
                text=self._current_text,
                delay=round(delay, 2)
            )
            self.actions.append(action)

            if self.on_action:
                self.on_action(f"Type: {self._current_text[:20]}...")

            self._current_text = ""

    def save(self, path: Path = None):
        """Save recorded actions to file."""
        path = path or MACRO_FILE
        path.parent.mkdir(parents=True, exist_ok=True)

        data = {
            "recorded_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "actions": [asdict(a) for a in self.actions]
        }

        with open(path, "w") as f:
            json.dump(data, f, indent=2)

    def load(self, path: Path = None) -> bool:
        """Load recorded actions from file."""
        path = path or MACRO_FILE

        if not path.exists():
            return False

        with open(path) as f:
            data = json.load(f)

        self.actions = [
            MacroAction(**a) for a in data.get("actions", [])
        ]
        return True


class MacroPlayer:
    """Plays back recorded macro actions."""

    def __init__(self, on_action: Optional[Callable[[str], None]] = None):
        self.on_action = on_action
        self.playing = False
        self.actions: List[MacroAction] = []

    def load(self, path: Path = None) -> bool:
        """Load macro from file."""
        path = path or MACRO_FILE

        if not path.exists():
            return False

        with open(path) as f:
            data = json.load(f)

        self.actions = [
            MacroAction(**a) for a in data.get("actions", [])
        ]
        return True

    def play(self, speed: float = 1.0):
        """Play back the recorded macro."""
        self.playing = True

        for i, action in enumerate(self.actions):
            if not self.playing:
                break

            # Wait for delay (adjusted by speed)
            if action.delay > 0.1:
                wait_time = min(action.delay / speed, 2.0)  # Cap at 2 seconds
                time.sleep(wait_time)

            # Execute action
            if action.action_type == "click":
                if self.on_action:
                    self.on_action(f"Click ({action.x}, {action.y})")
                pyautogui.click(action.x, action.y)

            elif action.action_type == "type":
                if self.on_action:
                    self.on_action(f"Type: {action.text[:20]}...")
                pyautogui.write(action.text, interval=0.02)

            elif action.action_type == "key":
                if self.on_action:
                    self.on_action(f"Key: {action.key}")
                pyautogui.press(action.key)

            time.sleep(0.1)  # Small delay between actions

        self.playing = False

    def stop(self):
        """Stop playback."""
        self.playing = False
