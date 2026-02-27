"""
Control overlay with pause/play and kill buttons.

Provides a small floating control panel in the bottom-right corner
with buttons to control the agentic script execution.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk
import threading
import time
from typing import Optional, Callable, List
from pathlib import Path
import logging
import sys
import os

logger = logging.getLogger("mhtagentic.control_overlay")


def _compute_scale(root):
    """Compute UI scale factor relative to 1920x1080 baseline."""
    try:
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        return min(sw / 1920, sh / 1080)
    except:
        return 1.0


def _sf(base_size, scale):
    """Scale a font size, minimum 7."""
    return max(7, round(base_size * scale))


class ModeSelectionOverlay:
    """
    Initial mode selection overlay.
    Lets user choose between Agentic (learning) and Saved (fast) modes.
    """

    def __init__(self, on_select: Optional[Callable[[str], None]] = None):
        self.on_select = on_select
        self.root: Optional[tk.Tk] = None
        self.panel: Optional[tk.Toplevel] = None
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._selected_mode: Optional[str] = None
        self._selection_event = threading.Event()

    def start(self):
        """Start the mode selection overlay."""
        if self._running:
            return
        self._running = True
        self._selected_mode = None
        self._selection_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        time.sleep(0.3)

    def _run(self):
        """Run the overlay."""
        try:
            self.root = tk.Tk()
            self.root.withdraw()
            self._create_panel()
            self.root.mainloop()
        except Exception as e:
            logger.error(f"Mode selection error: {e}")
        finally:
            self._running = False

    def _create_panel(self):
        """Create the mode selection panel."""
        s = _compute_scale(self.root)
        self.panel = tk.Toplevel(self.root)
        self.panel.title("MHT Agentic")
        self.panel.attributes("-topmost", True)
        self.panel.overrideredirect(True)

        panel_width = int(260 * s)
        panel_height = int(300 * s)

        # Center on screen
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_pos = (screen_width - panel_width) // 2
        y_pos = (screen_height - panel_height) // 2

        self.panel.geometry(f"{panel_width}x{panel_height}+{x_pos}+{y_pos}")

        # Colors
        bg_color = "#1a1a2e"
        accent_color = "#4ECDC4"

        self.panel.configure(bg=bg_color)

        # Main frame
        main_frame = tk.Frame(self.panel, bg=bg_color, highlightbackground=accent_color, highlightthickness=2)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)

        # Title
        title_frame = tk.Frame(main_frame, bg=accent_color, height=int(28 * s))
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        title_label = tk.Label(
            title_frame,
            text="MHT Agentic",
            font=("Segoe UI", _sf(10, s), "bold"),
            bg=accent_color,
            fg="#1a1a2e"
        )
        title_label.pack(pady=int(4 * s))

        # Subtitle
        subtitle = tk.Label(
            main_frame,
            text="Select Mode",
            font=("Segoe UI", _sf(9, s)),
            bg=bg_color,
            fg="#888888"
        )
        subtitle.pack(pady=(int(12 * s), int(8 * s)))

        # Button frame
        btn_frame = tk.Frame(main_frame, bg=bg_color)
        btn_frame.pack(fill=tk.X, padx=int(18 * s), pady=int(4 * s))

        # Saved mode button (recommended)
        saved_btn = tk.Button(
            btn_frame,
            text="Saved (Fast)",
            font=("Segoe UI", _sf(10, s), "bold"),
            bg="#28a745",
            fg="white",
            activebackground="#218838",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self._select_mode("saved")
        )
        saved_btn.pack(fill=tk.X, pady=int(4 * s), ipady=int(6 * s))

        saved_desc = tk.Label(
            btn_frame,
            text="Replay recorded login",
            font=("Segoe UI", _sf(8, s)),
            bg=bg_color,
            fg="#666666"
        )
        saved_desc.pack()

        # Record mode button
        record_btn = tk.Button(
            btn_frame,
            text="Record",
            font=("Segoe UI", _sf(10, s)),
            bg="#ff9800",
            fg="white",
            activebackground="#f57c00",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self._select_mode("record")
        )
        record_btn.pack(fill=tk.X, pady=(int(10 * s), int(4 * s)), ipady=int(6 * s))

        record_desc = tk.Label(
            btn_frame,
            text="Record your login actions",
            font=("Segoe UI", _sf(8, s)),
            bg=bg_color,
            fg="#666666"
        )
        record_desc.pack()

        # Agentic mode button
        agentic_btn = tk.Button(
            btn_frame,
            text="Agentic (AI)",
            font=("Segoe UI", _sf(10, s)),
            bg="#0097d6",
            fg="white",
            activebackground="#0077a8",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=lambda: self._select_mode("agentic")
        )
        agentic_btn.pack(fill=tk.X, pady=(int(10 * s), int(4 * s)), ipady=int(6 * s))

        agentic_desc = tk.Label(
            btn_frame,
            text="AI learns new EMR (slower)",
            font=("Segoe UI", _sf(8, s)),
            bg=bg_color,
            fg="#666666"
        )
        agentic_desc.pack()

        # Hover effects
        self._add_hover(saved_btn, "#28a745", "#218838")
        self._add_hover(record_btn, "#ff9800", "#f57c00")
        self._add_hover(agentic_btn, "#0097d6", "#0077a8")

    def _add_hover(self, btn, normal, hover):
        btn.bind("<Enter>", lambda e: btn.config(bg=hover))
        btn.bind("<Leave>", lambda e: btn.config(bg=normal))

    def _select_mode(self, mode: str):
        """Handle mode selection."""
        self._selected_mode = mode
        self._selection_event.set()
        if self.on_select:
            self.on_select(mode)
        self.stop()

    def wait_for_selection(self, timeout: float = None) -> Optional[str]:
        """Wait for user to select a mode."""
        self._selection_event.wait(timeout)
        return self._selected_mode

    def stop(self):
        """Stop the overlay."""
        self._running = False
        self._selection_event.set()
        if self.root:
            try:
                self.root.after(0, self._close)
            except:
                pass

    def _close(self):
        try:
            if self.panel:
                self.panel.destroy()
            if self.root:
                self.root.quit()
                self.root.destroy()
        except:
            pass


class ControlOverlay:
    """
    Floating control panel with pause/play and kill buttons.

    Positioned in the bottom-right corner of the screen.
    """

    def __init__(
        self,
        on_proceed: Optional[Callable[[], None]] = None,
        on_kill: Optional[Callable[[], None]] = None,
        on_debug: Optional[Callable[[], None]] = None,
        on_record: Optional[Callable[[], None]] = None,
        logo_path: Optional[str] = None
    ):
        """
        Initialize the control overlay.

        Args:
            on_proceed: Callback when proceed/play button is clicked
            on_kill: Callback when kill button is clicked
            on_debug: Callback when debug button is clicked
            on_record: Callback when record button is clicked
            logo_path: Path to logo image (optional)
        """
        self.on_proceed = on_proceed
        self.on_kill = on_kill
        self.on_debug = on_debug
        self.on_record = on_record
        self.logo_path = logo_path

        self.root: Optional[tk.Tk] = None
        self.panel: Optional[tk.Toplevel] = None
        self._running = False
        self._paused = True  # Start paused, waiting for user to proceed
        self._killed = False
        self._recording = False
        self._thread: Optional[threading.Thread] = None
        self._proceed_event = threading.Event()
        self._input_event = threading.Event()
        self._input_value: Optional[str] = None
        self._status_text = "Ready"
        self._step_text = "Waiting to start..."
        self._logs: list = []  # Store logs for debug export
        self._input_frame: Optional[tk.Frame] = None
        self._input_entry: Optional[tk.Entry] = None
        self._input_label: Optional[tk.Label] = None
        self._button_frame: Optional[tk.Frame] = None
        self._refresh_counter = 0  # Counter for periodic panel refresh

    @property
    def is_paused(self) -> bool:
        """Check if execution is paused."""
        return self._paused

    @property
    def is_killed(self) -> bool:
        """Check if kill was requested."""
        return self._killed

    def start(self):
        """Start the control overlay."""
        if self._running:
            return

        self._running = True
        self._killed = False
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        time.sleep(0.3)  # Wait for window to initialize

    def _run(self):
        """Run the overlay in a separate thread."""
        while self._running and not self._killed:
            try:
                self.root = tk.Tk()
                self.root.withdraw()

                self._create_panel()
                self.root.mainloop()

            except Exception as e:
                logger.error(f"Control overlay error: {e}")

            # If we're still supposed to be running, restart the overlay
            # This handles both exceptions AND unexpected mainloop exits
            if self._running and not self._killed:
                logger.warning("Overlay mainloop exited - regenerating panel...")
                import time
                time.sleep(0.3)
                continue
            break  # Only exit loop if explicitly stopped
        self._running = False

    def _create_panel(self):
        """Create the control panel window."""
        s = _compute_scale(self.root)
        self._scale = s  # Store for prompt_input resizing
        self.panel = tk.Toplevel(self.root)
        self.panel.title("MHT Agentic")
        self.panel.attributes("-topmost", True)
        self.panel.overrideredirect(True)  # No window decorations

        # Register for global hide/show
        register_overlay_window(self.panel)

        # Dynamic sizing based on screen resolution
        panel_width = int(260 * s)
        panel_height = int(240 * s)
        self._panel_width = panel_width
        self._panel_height = panel_height

        # Position in bottom-right corner (percentage-based)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_pos = screen_width - panel_width - int(screen_width * 0.02)
        y_pos = int(screen_height * 0.63)

        self.panel.geometry(f"{panel_width}x{panel_height}+{x_pos}+{y_pos}")

        # Elegant dark theme - softer colors
        bg_color = "#1a1a24"
        card_color = "#252532"
        accent_color = "#8b9dc3"
        text_primary = "#f0f0f0"
        text_secondary = "#9a9a9a"
        success_color = "#5cb85c"
        danger_color = "#d9534f"

        self.panel.configure(bg=bg_color)

        # Main frame with subtle border
        main_frame = tk.Frame(self.panel, bg=bg_color, highlightbackground="#333344", highlightthickness=1)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # Title bar (draggable)
        title_frame = tk.Frame(main_frame, bg=card_color)
        title_frame.pack(fill=tk.X)

        title_label = tk.Label(
            title_frame,
            text="MHT Agentic",
            font=("Segoe UI", _sf(11, s), "bold"),
            bg=card_color,
            fg=text_primary
        )
        title_label.pack(side=tk.LEFT, padx=int(15 * s), pady=int(8 * s))

        # X button to close/kill (top right)
        close_btn = tk.Label(
            title_frame,
            text="X",
            font=("Segoe UI", _sf(10, s), "bold"),
            bg=card_color,
            fg="#888888",
            cursor="hand2"
        )
        close_btn.pack(side=tk.RIGHT, padx=int(12 * s), pady=int(6 * s))
        close_btn.bind("<Enter>", lambda e: close_btn.configure(fg=danger_color))
        close_btn.bind("<Leave>", lambda e: close_btn.configure(fg="#888888"))
        close_btn.bind("<Button-1>", lambda e: self._on_kill_click())

        # Small accent bar under title
        accent_bar = tk.Frame(main_frame, bg=accent_color, height=2)
        accent_bar.pack(fill=tk.X)

        # Make title bar draggable
        title_frame.bind("<Button-1>", self._start_drag)
        title_frame.bind("<B1-Motion>", self._drag)
        title_label.bind("<Button-1>", self._start_drag)
        title_label.bind("<B1-Motion>", self._drag)

        # Status area
        pad_x = int(15 * s)
        status_frame = tk.Frame(main_frame, bg=bg_color)
        status_frame.pack(fill=tk.X, padx=pad_x, pady=(int(12 * s), int(10 * s)))

        self.status_label = tk.Label(
            status_frame,
            text=self._status_text,
            font=("Segoe UI", _sf(11, s)),
            bg=bg_color,
            fg=accent_color,
            anchor="w"
        )
        self.status_label.pack(fill=tk.X)

        self.step_label = tk.Label(
            status_frame,
            text=self._step_text,
            font=("Segoe UI", _sf(9, s)),
            bg=bg_color,
            fg=text_secondary,
            anchor="w",
            wraplength=panel_width - int(40 * s),
            justify="left"
        )
        self.step_label.pack(fill=tk.X, pady=(int(6 * s), 0))

        # Button frame
        button_frame = tk.Frame(main_frame, bg=bg_color)
        button_frame.pack(fill=tk.X, padx=pad_x, pady=(int(10 * s), int(15 * s)))

        # Continue button
        self.proceed_btn = tk.Button(
            button_frame,
            text="Continue",
            font=("Segoe UI", _sf(10, s)),
            bg=success_color,
            fg="white",
            activebackground="#4a9d4a",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=self._on_proceed_click
        )
        self.proceed_btn.pack(fill=tk.X, pady=(0, int(8 * s)), ipady=int(6 * s))

        # Stop button
        self.kill_btn = tk.Button(
            button_frame,
            text="Stop",
            font=("Segoe UI", _sf(9, s)),
            bg=card_color,
            fg=danger_color,
            activebackground="#333344",
            activeforeground=danger_color,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._on_kill_click
        )
        self.kill_btn.pack(fill=tk.X, ipady=int(4 * s))

        # Hidden debug/record buttons (keep for functionality but don't show)
        self.debug_btn = tk.Button(button_frame, command=self._on_debug_click)
        self.record_btn = tk.Button(button_frame, command=self._on_record_click)

        # Hover effects - subtle
        self._add_hover_effect(self.proceed_btn, success_color, "#5daf61")
        self._add_hover_effect(self.kill_btn, card_color, "#3a3a4e")

        # Store colors for input frame
        self._colors = {
            'bg': bg_color,
            'card': card_color,
            'accent': accent_color,
            'text_primary': text_primary,
            'text_secondary': text_secondary,
            'success': success_color
        }

        # Input frame (hidden by default) - store reference to button_frame for packing
        self._button_frame = button_frame

        self._input_frame = tk.Frame(main_frame, bg=card_color)
        # Don't pack yet - will be shown when needed

        self._input_label = tk.Label(
            self._input_frame,
            text="Enter OTP Code",
            font=("Segoe UI", _sf(9, s)),
            bg=card_color,
            fg=text_primary
        )
        self._input_label.pack(fill=tk.X, padx=int(12 * s), pady=(int(10 * s), int(4 * s)))

        self._input_entry = tk.Entry(
            self._input_frame,
            font=("Segoe UI", _sf(12, s)),
            bg="#3a3a4e",
            fg=text_primary,
            insertbackground=text_primary,
            relief=tk.FLAT,
            justify="center"
        )
        self._input_entry.pack(fill=tk.X, padx=int(12 * s), pady=int(4 * s), ipady=int(6 * s))
        self._input_entry.bind("<Return>", lambda e: self._on_input_submit())

        self._input_submit_btn = tk.Button(
            self._input_frame,
            text="Submit",
            font=("Segoe UI", _sf(9, s)),
            bg=success_color,
            fg="white",
            activebackground="#4a9d4a",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            command=self._on_input_submit
        )
        self._input_submit_btn.pack(fill=tk.X, padx=int(12 * s), pady=(int(4 * s), int(10 * s)), ipady=int(4 * s))
        self._add_hover_effect(self._input_submit_btn, success_color, "#4a9d4a")

    def _add_hover_effect(self, button, normal_color, hover_color):
        """Add hover effect to a button."""
        def on_enter(e):
            button.config(bg=hover_color)
        def on_leave(e):
            button.config(bg=normal_color)
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def _start_drag(self, event):
        """Start dragging the window."""
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def _drag(self, event):
        """Handle window dragging."""
        if self.panel:
            x = self.panel.winfo_x() + (event.x - self._drag_start_x)
            y = self.panel.winfo_y() + (event.y - self._drag_start_y)
            self.panel.geometry(f"+{x}+{y}")

    def _on_proceed_click(self):
        """Handle proceed button click."""
        self._paused = False
        self._proceed_event.set()

        if self.on_proceed:
            self.on_proceed()

        # Update button text to show waiting state
        self.proceed_btn.config(text="WAITING...", bg="#6c757d")

    def _on_kill_click(self):
        """Handle kill button click."""
        self._killed = True
        self._paused = False  # Unblock any waiting
        self._proceed_event.set()

        if self.on_kill:
            self.on_kill()

        # Update UI to show killed state
        self.kill_btn.config(text="KILLED", bg="#6c757d")
        self.proceed_btn.config(state=tk.DISABLED)
        self.set_status("Killed by user", "#dc3545")

    def _on_debug_click(self):
        """Handle debug button click - takes screenshot and saves logs."""
        self.debug_btn.config(text="SAVING...", bg="#ffc107")

        if self.on_debug:
            self.on_debug()

        # Reset button after a moment
        if self.root:
            try:
                self.root.after(1500, lambda: self.debug_btn.config(text="DEBUG", bg="#6c757d"))
            except:
                pass

    def _on_record_click(self):
        """Handle record button click - toggle recording."""
        self._recording = not self._recording

        if self._recording:
            self.record_btn.config(text="STOP REC", bg="#f44336")
            self.set_status("Recording...", "#f44336")
        else:
            self.record_btn.config(text="RECORD", bg="#ff9800")
            self.set_status("Recording saved")

        if self.on_record:
            self.on_record()

    @property
    def is_recording(self) -> bool:
        """Check if recording is active."""
        return self._recording

    def add_log(self, message: str):
        """Add a log message to the internal log buffer."""
        import time
        timestamp = time.strftime("%H:%M:%S")
        self._logs.append(f"[{timestamp}] {message}")

    def get_logs(self) -> list:
        """Get all logged messages."""
        return self._logs.copy()

    def clear_logs(self):
        """Clear the log buffer."""
        self._logs.clear()

    def wait_for_proceed(self, timeout: Optional[float] = None) -> bool:
        """
        Wait for user to click proceed.

        Args:
            timeout: Max seconds to wait (None = wait forever)

        Returns:
            True if proceed was clicked, False if killed or timeout
        """
        self._paused = True
        self._proceed_event.clear()

        # Reset button to ready state
        if self.proceed_btn:
            self._safe_after(lambda: self.proceed_btn.config(text="PROCEED", bg="#28a745"))

        result = self._proceed_event.wait(timeout)

        if self._killed:
            return False

        return result

    def _safe_after(self, func):
        """Safely schedule a function on the Tkinter thread."""
        try:
            if self.root and self._running:
                # Check if root still exists
                try:
                    self.root.winfo_exists()
                    self.root.after(0, func)
                except:
                    pass  # Root is gone, ignore
        except Exception as e:
            pass  # Tkinter in bad state, ignore

    def set_status(self, text: str, color: str = "#4ECDC4"):
        """Update status text (thread-safe)."""
        self._status_text = text
        if self.status_label and self.root:
            self._safe_after(lambda: self.status_label.config(text=text, fg=color))

    def set_step(self, text: str):
        """Update step description (thread-safe)."""
        self._step_text = text
        if self.step_label and self.root:
            self._safe_after(lambda: self.step_label.config(text=text))
        # Periodically refresh panel visibility every 5 step updates
        self._refresh_counter += 1
        if self._refresh_counter >= 5:
            self._refresh_counter = 0
            self.refresh_panel()

    def refresh_panel(self):
        """Explicitly refresh and bring panel to foreground (thread-safe)."""
        if self.panel and self.root and self._running:
            def do_refresh():
                try:
                    if self.panel and self.panel.winfo_exists():
                        self.panel.update_idletasks()
                        self.panel.lift()
                        self.panel.attributes("-topmost", True)
                        self.panel.focus_force()
                except Exception as e:
                    logger.warning(f"Panel refresh error: {e}")
            self._safe_after(do_refresh)

    def hide(self):
        """Hide the overlay panel (thread-safe)."""
        if self.panel and self.root and self._running:
            def do_hide():
                try:
                    if self.panel and self.panel.winfo_exists():
                        self.panel.withdraw()
                except Exception as e:
                    logger.warning(f"Panel hide error: {e}")
            self._safe_after(do_hide)

    def show(self):
        """Show the overlay panel (thread-safe)."""
        if self.panel and self.root and self._running:
            def do_show():
                try:
                    if self.panel and self.panel.winfo_exists():
                        self.panel.deiconify()
                        self.panel.lift()
                        self.panel.attributes("-topmost", True)
                except Exception as e:
                    logger.warning(f"Panel show error: {e}")
            self._safe_after(do_show)

    def enable_proceed(self):
        """Enable and reset the proceed button (thread-safe)."""
        if self.proceed_btn and self.root:
            self._safe_after(lambda: self.proceed_btn.config(
                text="PROCEED",
                bg="#28a745",
                state=tk.NORMAL
            ))

    def _on_input_submit(self):
        """Handle input submission."""
        if self._input_entry:
            self._input_value = self._input_entry.get()
            self._input_event.set()

    def prompt_input(self, prompt_text: str = "Enter OTP Code:", timeout: Optional[float] = None) -> Optional[str]:
        """
        Show input prompt and wait for user to enter value.

        Args:
            prompt_text: Text to show above input field
            timeout: Max seconds to wait (None = wait forever)

        Returns:
            The entered value, or None if timeout/cancelled
        """
        self._input_value = None
        self._input_event.clear()

        # Show input frame and update label
        def show_input():
            try:
                if self._input_label:
                    self._input_label.config(text=prompt_text)
                if self._input_entry:
                    self._input_entry.delete(0, tk.END)
                if self._input_frame and self._button_frame:
                    # Pack input frame BEFORE the button frame so it's visible
                    self._input_frame.pack(fill=tk.X, padx=20, pady=(0, 15), before=self._button_frame)
                # Resize panel to fit input - keep same position by adjusting y
                if self.panel:
                    s = getattr(self, '_scale', 1.0)
                    screen_height = self.root.winfo_screenheight()
                    new_height = int(360 * s)
                    pw = getattr(self, '_panel_width', int(260 * s))
                    x_pos = self.panel.winfo_x()
                    y_pos = int(screen_height * 0.63)
                    self.panel.geometry(f"{pw}x{new_height}+{x_pos}+{y_pos}")
                    self.panel.update()
                # Focus the entry after a short delay
                if self._input_entry and self.root:
                    self.root.after(100, lambda: self._input_entry.focus_set())
            except Exception as e:
                print(f"Error showing input: {e}")

        self._safe_after(show_input)
        # Also try immediate update
        time.sleep(0.2)

        # Wait for input
        result = self._input_event.wait(timeout)

        # Hide input frame and ensure panel stays visible
        def hide_input():
            try:
                if self._input_frame:
                    self._input_frame.pack_forget()
                if self.panel and self.root:
                    s = getattr(self, '_scale', 1.0)
                    screen_height = self.root.winfo_screenheight()
                    ph = getattr(self, '_panel_height', int(240 * s))
                    pw = getattr(self, '_panel_width', int(260 * s))
                    x_pos = self.panel.winfo_x()
                    y_pos = int(screen_height * 0.63)
                    self.panel.geometry(f"{pw}x{ph}+{x_pos}+{y_pos}")
                    # CRITICAL: Ensure panel remains visible after hiding input
                    self.panel.update_idletasks()
                    self.panel.lift()
                    self.panel.attributes("-topmost", True)
                    # Schedule another lift after a short delay to ensure visibility
                    self.root.after(100, lambda: self.panel.lift() if self.panel else None)
                    self.root.after(500, lambda: self.panel.lift() if self.panel else None)
            except Exception as e:
                logger.warning(f"Error hiding input: {e}")

        self._safe_after(hide_input)
        # Also refresh panel after a short delay to ensure it stays visible
        time.sleep(0.3)
        self.refresh_panel()

        if self._killed or not result:
            return None

        return self._input_value

    def hide_input(self):
        """Hide the input frame (thread-safe)."""
        if self._input_frame and self.root:
            self._safe_after(lambda: self._input_frame.pack_forget())

    def stop(self):
        """Stop the control overlay."""
        self._running = False
        self._proceed_event.set()  # Unblock any waiting

        if self.root:
            try:
                self.root.after(0, self._close_internal)
            except:
                pass

    def _close_internal(self):
        """Internal close method."""
        try:
            if self.panel:
                self.panel.destroy()
            if self.root:
                self.root.quit()
                self.root.destroy()
        except:
            pass
        self._running = False


class AnalyticsDashboardOverlay:
    """
    Floating analytics dashboard showing real-time processing metrics.

    Displays:
    - Patients processed today
    - Success rate
    - Average time per patient
    - Patients per hour
    - Weekly summary
    - Value metrics
    """

    def __init__(self):
        self.root: Optional[tk.Tk] = None
        self.panel: Optional[tk.Toplevel] = None
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._stats: dict = {}
        self._value_metrics: dict = {}

        # Labels for updating
        self._labels: dict = {}

    def start(self):
        """Start the analytics dashboard overlay."""
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        time.sleep(0.3)

    def _run(self):
        """Run the overlay."""
        try:
            self.root = tk.Tk()
            self.root.withdraw()
            self._create_panel()
            self.root.mainloop()
        except Exception as e:
            logger.error(f"Analytics dashboard error: {e}")
        finally:
            self._running = False

    def _create_panel(self):
        """Create the analytics dashboard panel."""
        s = _compute_scale(self.root)
        self._scale = s
        self.panel = tk.Toplevel(self.root)
        self.panel.title("MHT Analytics")
        self.panel.attributes("-topmost", True)
        self.panel.overrideredirect(True)

        panel_width = int(300 * s)
        panel_height = int(400 * s)

        # Position in top-right, offset from other overlays
        screen_width = self.root.winfo_screenwidth()
        x_pos = screen_width - panel_width - int(360 * s)
        y_pos = int(30 * s)

        self.panel.geometry(f"{panel_width}x{panel_height}+{x_pos}+{y_pos}")

        # Colors
        bg_color = "#1a1a24"
        card_color = "#252532"
        accent_color = "#4ECDC4"
        text_primary = "#f0f0f0"
        text_secondary = "#9a9a9a"
        success_color = "#5cb85c"
        warning_color = "#f0ad4e"

        self.panel.configure(bg=bg_color)

        # Main frame
        main_frame = tk.Frame(self.panel, bg=bg_color, highlightbackground="#333344", highlightthickness=1)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        # Title bar (draggable)
        title_frame = tk.Frame(main_frame, bg=card_color)
        title_frame.pack(fill=tk.X)

        title_label = tk.Label(
            title_frame,
            text="Analytics Dashboard",
            font=("Segoe UI", _sf(10, s), "bold"),
            bg=card_color,
            fg=text_primary
        )
        title_label.pack(side=tk.LEFT, padx=int(10 * s), pady=int(7 * s))

        # Live indicator
        self._labels['live'] = tk.Label(
            title_frame,
            text="● LIVE",
            font=("Segoe UI", _sf(8, s), "bold"),
            bg=card_color,
            fg=success_color
        )
        self._labels['live'].pack(side=tk.RIGHT, padx=int(10 * s), pady=int(7 * s))

        # Accent bar
        tk.Frame(main_frame, bg=accent_color, height=2).pack(fill=tk.X)

        # Make title bar draggable
        title_frame.bind("<Button-1>", self._start_drag)
        title_frame.bind("<B1-Motion>", self._drag)
        title_label.bind("<Button-1>", self._start_drag)
        title_label.bind("<B1-Motion>", self._drag)

        # Content frame with padding
        pad = int(10 * s)
        content_frame = tk.Frame(main_frame, bg=bg_color)
        content_frame.pack(fill=tk.BOTH, expand=True, padx=pad, pady=int(6 * s))

        # === SESSION STATS SECTION ===
        self._create_section(content_frame, "Current Session", bg_color, text_primary, s)

        stats_grid = tk.Frame(content_frame, bg=bg_color)
        stats_grid.pack(fill=tk.X, pady=(int(3 * s), int(8 * s)))

        # Row 1: Patients processed & Success rate
        row1 = tk.Frame(stats_grid, bg=bg_color)
        row1.pack(fill=tk.X)

        self._create_stat_box(row1, "patients_processed", "Patients", "0", success_color, bg_color, card_color, text_primary, s)
        self._create_stat_box(row1, "success_rate", "Success", "0%", accent_color, bg_color, card_color, text_primary, s)

        # Row 2: Avg time & Patients/hour
        row2 = tk.Frame(stats_grid, bg=bg_color)
        row2.pack(fill=tk.X, pady=(int(3 * s), 0))

        self._create_stat_box(row2, "avg_time", "Avg Time", "0s", warning_color, bg_color, card_color, text_primary, s)
        self._create_stat_box(row2, "patients_per_hour", "Per Hour", "0", "#9c27b0", bg_color, card_color, text_primary, s)

        # === DETAILED METRICS ===
        self._create_section(content_frame, "Processing Breakdown", bg_color, text_primary, s)

        details_frame = tk.Frame(content_frame, bg=card_color)
        details_frame.pack(fill=tk.X, pady=(int(3 * s), int(8 * s)))

        self._create_detail_row(details_frame, "successful", "Successful:", "0", success_color, card_color, text_primary, text_secondary, s)
        self._create_detail_row(details_frame, "partial", "Partial:", "0", warning_color, card_color, text_primary, text_secondary, s)
        self._create_detail_row(details_frame, "failed", "Failed:", "0", "#d9534f", card_color, text_primary, text_secondary, s)
        self._create_detail_row(details_frame, "skipped", "Skipped:", "0", text_secondary, card_color, text_primary, text_secondary, s)
        self._create_detail_row(details_frame, "scans", "WR Scans:", "0", accent_color, card_color, text_primary, text_secondary, s)
        self._create_detail_row(details_frame, "popups", "Popups:", "0", "#ff9800", card_color, text_primary, text_secondary, s)

        # === VALUE METRICS ===
        self._create_section(content_frame, "Value Assessment", bg_color, text_primary, s)

        value_frame = tk.Frame(content_frame, bg=card_color)
        value_frame.pack(fill=tk.X, pady=(int(3 * s), int(6 * s)))

        self._create_detail_row(value_frame, "efficiency", "Efficiency:", "0x faster", success_color, card_color, text_primary, text_secondary, s)
        self._create_detail_row(value_frame, "time_saved", "Time Saved:", "0s/patient", accent_color, card_color, text_primary, text_secondary, s)
        self._create_detail_row(value_frame, "weekly_hours", "Weekly Hours:", "0h saved", "#9c27b0", card_color, text_primary, text_secondary, s)

        # Recommendation label
        self._labels['recommendation'] = tk.Label(
            value_frame,
            text="Gathering data...",
            font=("Segoe UI", _sf(8, s)),
            bg=card_color,
            fg=text_secondary,
            wraplength=panel_width - int(40 * s),
            justify="left"
        )
        self._labels['recommendation'].pack(fill=tk.X, padx=int(8 * s), pady=(int(3 * s), int(6 * s)))

        # === DURATION ===
        duration_frame = tk.Frame(main_frame, bg=card_color)
        duration_frame.pack(fill=tk.X, side=tk.BOTTOM)

        self._labels['duration'] = tk.Label(
            duration_frame,
            text="Session: 0m | Monitoring...",
            font=("Segoe UI", _sf(8, s)),
            bg=card_color,
            fg=text_secondary
        )
        self._labels['duration'].pack(pady=int(6 * s))

    def _create_section(self, parent, title, bg_color, text_color, scale=1.0):
        """Create a section header."""
        tk.Label(
            parent,
            text=title,
            font=("Segoe UI", _sf(9, scale), "bold"),
            bg=bg_color,
            fg=text_color,
            anchor="w"
        ).pack(fill=tk.X)

    def _create_stat_box(self, parent, key, label, value, accent, bg_color, card_color, text_color, scale=1.0):
        """Create a stat box widget."""
        frame = tk.Frame(parent, bg=card_color)
        frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, int(3 * scale)))

        # Accent bar on left
        tk.Frame(frame, bg=accent, width=max(2, int(3 * scale))).pack(side=tk.LEFT, fill=tk.Y)

        content = tk.Frame(frame, bg=card_color)
        content.pack(fill=tk.BOTH, expand=True, padx=int(6 * scale), pady=int(5 * scale))

        self._labels[f"{key}_value"] = tk.Label(
            content,
            text=value,
            font=("Segoe UI", _sf(14, scale), "bold"),
            bg=card_color,
            fg=text_color
        )
        self._labels[f"{key}_value"].pack(anchor="w")

        tk.Label(
            content,
            text=label,
            font=("Segoe UI", _sf(8, scale)),
            bg=card_color,
            fg="#888888"
        ).pack(anchor="w")

    def _create_detail_row(self, parent, key, label, value, value_color, bg_color, text_color, label_color, scale=1.0):
        """Create a detail row."""
        row = tk.Frame(parent, bg=bg_color)
        row.pack(fill=tk.X, padx=int(8 * scale), pady=1)

        tk.Label(
            row,
            text=label,
            font=("Segoe UI", _sf(8, scale)),
            bg=bg_color,
            fg=label_color,
            width=12,
            anchor="w"
        ).pack(side=tk.LEFT)

        self._labels[key] = tk.Label(
            row,
            text=value,
            font=("Segoe UI", _sf(8, scale), "bold"),
            bg=bg_color,
            fg=value_color
        )
        self._labels[key].pack(side=tk.RIGHT)

    def _start_drag(self, event):
        """Start dragging the window."""
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def _drag(self, event):
        """Handle window dragging."""
        if self.panel:
            x = self.panel.winfo_x() + (event.x - self._drag_start_x)
            y = self.panel.winfo_y() + (event.y - self._drag_start_y)
            self.panel.geometry(f"+{x}+{y}")

    def update_stats(self, stats: dict, value_metrics: dict = None):
        """Update displayed statistics (thread-safe)."""
        self._stats = stats
        if value_metrics:
            self._value_metrics = value_metrics

        if self.root and self._running:
            try:
                self.root.after(0, self._do_update)
            except:
                pass

    def _do_update(self):
        """Actually update the labels."""
        try:
            # Session stats
            if 'patients_processed_value' in self._labels:
                self._labels['patients_processed_value'].config(
                    text=str(self._stats.get('total_patients_processed', 0))
                )

            if 'success_rate_value' in self._labels:
                self._labels['success_rate_value'].config(
                    text=f"{self._stats.get('success_rate', 0):.0f}%"
                )

            if 'avg_time_value' in self._labels:
                self._labels['avg_time_value'].config(
                    text=f"{self._stats.get('avg_time_per_patient_seconds', 0):.1f}s"
                )

            if 'patients_per_hour_value' in self._labels:
                self._labels['patients_per_hour_value'].config(
                    text=f"{self._stats.get('patients_per_hour', 0):.1f}"
                )

            # Detailed metrics
            if 'successful' in self._labels:
                self._labels['successful'].config(text=str(self._stats.get('successful_extractions', 0)))
            if 'partial' in self._labels:
                self._labels['partial'].config(text=str(self._stats.get('partial_extractions', 0)))
            if 'failed' in self._labels:
                self._labels['failed'].config(text=str(self._stats.get('failed_extractions', 0)))
            if 'skipped' in self._labels:
                self._labels['skipped'].config(text=str(self._stats.get('skipped_patients', 0)))
            if 'scans' in self._labels:
                self._labels['scans'].config(text=str(self._stats.get('waiting_room_scans', 0)))
            if 'popups' in self._labels:
                self._labels['popups'].config(text=str(self._stats.get('popups_dismissed', 0)))

            # Duration
            if 'duration' in self._labels:
                duration = self._stats.get('duration_minutes', 0)
                self._labels['duration'].config(text=f"Session: {duration:.0f}m | Monitoring...")

            # Value metrics
            if self._value_metrics:
                time_analysis = self._value_metrics.get('time_analysis', {})
                weekly = self._value_metrics.get('weekly_metrics', {})

                if 'efficiency' in self._labels:
                    eff = time_analysis.get('efficiency_multiplier', 0)
                    self._labels['efficiency'].config(text=f"{eff}x faster")

                if 'time_saved' in self._labels:
                    saved = time_analysis.get('time_saved_per_patient_seconds', 0)
                    self._labels['time_saved'].config(text=f"{saved:.0f}s/patient")

                if 'weekly_hours' in self._labels:
                    hours = weekly.get('time_saved_hours', 0)
                    self._labels['weekly_hours'].config(text=f"{hours:.1f}h saved")

                if 'recommendation' in self._labels:
                    rec = self._value_metrics.get('recommendation', 'Gathering data...')
                    # Shorten recommendation for display
                    if ':' in rec:
                        rec = rec.split(':')[0] + ": " + rec.split(':')[1][:60] + "..."
                    self._labels['recommendation'].config(text=rec)
        except Exception as e:
            logger.warning(f"Analytics update error: {e}")

    def hide(self):
        """Hide the analytics dashboard (thread-safe)."""
        if self.panel and self.root and self._running:
            try:
                self.root.after(0, lambda: self.panel.withdraw() if self.panel else None)
            except:
                pass

    def show(self):
        """Show the analytics dashboard (thread-safe)."""
        if self.panel and self.root and self._running:
            try:
                def do_show():
                    if self.panel:
                        self.panel.deiconify()
                        self.panel.lift()
                        self.panel.attributes("-topmost", True)
                self.root.after(0, do_show)
            except:
                pass

    def stop(self):
        """Stop the analytics dashboard."""
        self._running = False
        if self.root:
            try:
                self.root.after(0, self._close)
            except:
                pass

    def _close(self):
        try:
            if self.panel:
                self.panel.destroy()
            if self.root:
                self.root.quit()
                self.root.destroy()
        except:
            pass


class DemoStatusOverlay:
    """
    Sleek bottom-center status bar for demo mode.

    Shows inbound/outbound status in a minimal bar.
    Size: ~600x36, centered horizontally, 60px from bottom.
    Style: #3a3a3a bg, 0.85 alpha, no title bar, no borders.
    """

    def __init__(self):
        self.root: Optional[tk.Tk] = None
        self.panel: Optional[tk.Toplevel] = None
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._inbound_text = "Idle"
        self._outbound_text = "Idle"
        self._status_label: Optional[tk.Label] = None

    def start(self):
        """Start the demo status overlay."""
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        time.sleep(0.3)

    def _run(self):
        """Run the overlay."""
        try:
            self.root = tk.Tk()
            self.root.withdraw()
            self._create_panel()
            self.root.mainloop()
        except Exception as e:
            logger.error(f"Demo status overlay error: {e}")
        finally:
            self._running = False

    def _create_panel(self):
        """Create the status bar panel — dynamically resizes to fit text."""
        s = _compute_scale(self.root)
        self._s = s
        self.panel = tk.Toplevel(self.root)
        self.panel.overrideredirect(True)
        self.panel.attributes("-topmost", True)
        self.panel.attributes("-alpha", 0.90)

        self._panel_height = int(34 * s)
        self._radius = int(10 * s)
        self._bg = "#2d2d2d"
        self._transparent = "#f0f0f0"
        self._screen_width = self.root.winfo_screenwidth()
        self._screen_height = self.root.winfo_screenheight()
        self._y_pos = self._screen_height - self._panel_height - int(160 * s)

        self.panel.configure(bg=self._transparent)
        try:
            self.panel.attributes("-transparentcolor", self._transparent)
        except:
            self.panel.configure(bg=self._bg)

        self._canvas = tk.Canvas(
            self.panel, height=self._panel_height,
            bg=self._transparent, highlightthickness=0, bd=0
        )
        self._canvas.pack(fill=tk.BOTH, expand=True)

        self._status_label = tk.Label(
            self.panel,
            text=f"Inbound: {self._inbound_text}  |  Outbound: {self._outbound_text}",
            font=("Segoe UI", _sf(9, self._s)),
            bg=self._bg,
            fg="#e0e0e0"
        )

        # Initial size
        self._resize_to_fit()

    def _resize_to_fit(self):
        """Resize the panel to fit the current label text."""
        self._status_label.config(wraplength=0)  # No wrapping
        self._status_label.update_idletasks()
        text_width = self._status_label.winfo_reqwidth()
        text_height = self._status_label.winfo_reqheight()
        panel_width = text_width + 60  # generous padding
        panel_width = max(panel_width, 300)  # minimum

        x_pos = (self._screen_width - panel_width) // 2
        self.panel.geometry(f"{panel_width}x{self._panel_height}+{x_pos}+{self._y_pos}")

        self._canvas.config(width=panel_width)
        self._canvas.delete("all")
        self._draw_rounded_rect(self._canvas, 0, 0, panel_width, self._panel_height, self._radius, self._bg)
        self._canvas.create_window(panel_width // 2, self._panel_height // 2, window=self._status_label, width=text_width + 10)

    @staticmethod
    def _draw_rounded_rect(canvas, x1, y1, x2, y2, r, fill):
        """Draw a rounded rectangle on a canvas."""
        canvas.create_arc(x1, y1, x1 + 2*r, y1 + 2*r, start=90, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x2 - 2*r, y1, x2, y1 + 2*r, start=0, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x1, y2 - 2*r, x1 + 2*r, y2, start=180, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x2 - 2*r, y2 - 2*r, x2, y2, start=270, extent=90, fill=fill, outline=fill)
        canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline=fill)
        canvas.create_rectangle(x1, y1 + r, x2, y2 - r, fill=fill, outline=fill)

    def update_status(self, inbound_text=None, outbound_text=None):
        """Update the status bar text (thread-safe). Pass None to leave a side unchanged."""
        if inbound_text is not None:
            self._inbound_text = inbound_text
        if outbound_text is not None:
            self._outbound_text = outbound_text
        if self._status_label and self.root and self._running:
            try:
                self.root.after(0, self._do_update_status)
            except:
                pass

    def _do_update_status(self):
        """Update label text and resize panel to fit."""
        if self._status_label:
            self._status_label.config(
                text=f"Inbound: {self._inbound_text}  |  Outbound: {self._outbound_text}"
            )
            self._resize_to_fit()

    def stop(self):
        """Stop the overlay."""
        self._running = False
        if self.root:
            try:
                self.root.after(0, self._close)
            except:
                pass

    def _close(self):
        try:
            if self.panel:
                self.panel.destroy()
            if self.root:
                self.root.quit()
                self.root.destroy()
        except:
            pass


class DemoExtractedDataOverlay:
    """
    Sleek top-left extracted data panel for demo mode.

    Shows extracted patient data without pharmacy field.
    Only displays a patient once all key fields have been extracted.
    Rounded corners via transparent background + Canvas.
    Draggable via header area.
    """

    # Minimum fields required before showing a patient
    _REQUIRED_FIELDS = ('first_name', 'last_name', 'dob')

    def __init__(self):
        self.root: Optional[tk.Tk] = None
        self.panel: Optional[tk.Toplevel] = None
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._data_text: Optional[tk.Text] = None
        self._count_label: Optional[tk.Label] = None
        self._title_label: Optional[tk.Label] = None
        self._patients: list = []
        self._pending: dict = {}  # name -> patient_data (waiting for completion)

    def start(self):
        """Start the demo extracted data overlay."""
        if self._running:
            return
        self._running = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        time.sleep(0.3)

    def _run(self):
        """Run the overlay."""
        try:
            self.root = tk.Tk()
            self.root.withdraw()
            self._create_panel()
            self.root.mainloop()
        except Exception as e:
            logger.error(f"Demo extracted data overlay error: {e}")
        finally:
            self._running = False

    def _create_panel(self):
        """Create the extracted data panel with rounded corners."""
        s = _compute_scale(self.root)
        self._s = s
        self.panel = tk.Toplevel(self.root)
        self.panel.overrideredirect(True)
        self.panel.attributes("-topmost", True)
        self.panel.attributes("-alpha", 0.92)

        self._panel_width = int(400 * s)
        panel_height = int(300 * s)

        # Position bottom-left
        screen_height = self.root.winfo_screenheight()
        y_pos = screen_height - panel_height - int(50 * s)
        self.panel.geometry(f"{self._panel_width}x{panel_height}+{int(15 * s)}+{y_pos}")

        bg = "#2d2d2d"
        inner_bg = "#252525"
        text_fg = "#e0e0e0"
        muted_fg = "#9a9a9a"
        radius = int(14 * s)

        # Use a transparent color trick for rounded corners on Windows
        transparent = "#f0f0f0"
        self.panel.configure(bg=transparent)
        try:
            self.panel.attributes("-transparentcolor", transparent)
        except:
            # Fallback if transparentcolor not supported
            self.panel.configure(bg=bg)

        pw = self._panel_width

        # Canvas for rounded rectangle background
        canvas = tk.Canvas(
            self.panel, width=pw, height=panel_height,
            bg=transparent, highlightthickness=0, bd=0
        )
        canvas.pack(fill=tk.BOTH, expand=True)

        # Draw rounded rectangle
        self._draw_rounded_rect(canvas, 0, 0, pw, panel_height, radius, bg)

        # Header area on canvas — title + count
        hdr_y = int(8 * s)
        sep_y = int(30 * s)
        text_y = int(36 * s)
        self._title_label = tk.Label(
            self.panel,
            text="Extracted Data",
            font=("Segoe UI", _sf(9, s), "bold"),
            bg=bg, fg=text_fg
        )
        canvas.create_window(int(12 * s), hdr_y, anchor="nw", window=self._title_label)

        self._count_label = tk.Label(
            self.panel,
            text="0 patients",
            font=("Segoe UI", _sf(8, s)),
            bg=bg, fg=muted_fg
        )
        canvas.create_window(pw - int(12 * s), hdr_y + int(2 * s), anchor="ne", window=self._count_label)

        # Thin separator line on canvas
        canvas.create_line(int(10 * s), sep_y, pw - int(10 * s), sep_y, fill="#555555", width=1)

        # Text widget for patient data — no wrapping so lines stay on one line
        self._data_text = tk.Text(
            self.panel,
            bg=inner_bg,
            fg=text_fg,
            font=("Consolas", _sf(8, s)),
            wrap=tk.NONE,
            highlightthickness=0,
            borderwidth=0,
            insertbackground=text_fg,
            padx=int(5 * s),
            pady=int(3 * s)
        )
        canvas.create_window(
            int(10 * s), text_y, anchor="nw", width=pw - int(20 * s), height=panel_height - text_y - int(14 * s),
            window=self._data_text
        )
        self._data_text.insert(tk.END, "Waiting for qualified patients...\n")
        self._data_text.config(state=tk.DISABLED)

        # Make entire panel draggable via canvas
        canvas.bind("<Button-1>", self._start_drag)
        canvas.bind("<B1-Motion>", self._drag)
        self._title_label.bind("<Button-1>", self._start_drag)
        self._title_label.bind("<B1-Motion>", self._drag)
        self._count_label.bind("<Button-1>", self._start_drag)
        self._count_label.bind("<B1-Motion>", self._drag)

    @staticmethod
    def _draw_rounded_rect(canvas, x1, y1, x2, y2, r, fill):
        """Draw a rounded rectangle on a canvas."""
        canvas.create_arc(x1, y1, x1 + 2*r, y1 + 2*r, start=90, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x2 - 2*r, y1, x2, y1 + 2*r, start=0, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x1, y2 - 2*r, x1 + 2*r, y2, start=180, extent=90, fill=fill, outline=fill)
        canvas.create_arc(x2 - 2*r, y2 - 2*r, x2, y2, start=270, extent=90, fill=fill, outline=fill)
        canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline=fill)
        canvas.create_rectangle(x1, y1 + r, x2, y2 - r, fill=fill, outline=fill)

    def _start_drag(self, event):
        """Start dragging the window."""
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def _drag(self, event):
        """Handle window dragging."""
        if self.panel:
            x = self.panel.winfo_x() + (event.x - self._drag_start_x)
            y = self.panel.winfo_y() + (event.y - self._drag_start_y)
            self.panel.geometry(f"+{x}+{y}")

    def _is_complete(self, patient_data):
        """Check if patient data has all required fields populated."""
        for field in self._REQUIRED_FIELDS:
            val = patient_data.get(field, '')
            if not val or val == 'N/A':
                return False
        return True

    def add_patient(self, patient_data):
        """Add a patient's extracted data. Always displays immediately."""
        self._patients.append(patient_data)
        if self._data_text and self.root and self._running:
            try:
                self.root.after(0, lambda: self._do_add_patient(patient_data))
            except:
                pass

    def _do_add_patient(self, patient_data):
        """Replace display with latest patient data."""
        if not self._data_text:
            return
        self._data_text.config(state=tk.NORMAL)

        # Always clear and show only the latest patient
        self._data_text.delete(1.0, tk.END)

        # Extract fields (no pharmacy)
        first = patient_data.get('first_name', '')
        last = patient_data.get('last_name', '')
        dob = patient_data.get('dob', 'N/A')
        mrn = patient_data.get('mrn', 'N/A')
        phone = patient_data.get('cell_phone', 'N/A')
        email = patient_data.get('email', 'N/A')
        gender = patient_data.get('gender', 'N/A')
        insurance = patient_data.get('insurance', 'N/A')
        race = patient_data.get('race', 'N/A')
        ethnicity = patient_data.get('ethnicity', 'N/A')
        language = patient_data.get('language', 'N/A')

        self._data_text.insert(tk.END, f"  Name:      {last}, {first}\n")
        self._data_text.insert(tk.END, f"  DOB:       {dob}\n")
        self._data_text.insert(tk.END, f"  MRN:       {mrn}\n")
        self._data_text.insert(tk.END, f"  Phone:     {phone}\n")
        self._data_text.insert(tk.END, f"  Email:     {email}\n")
        self._data_text.insert(tk.END, f"  Gender:    {gender}\n")
        self._data_text.insert(tk.END, f"  Insurance: {insurance}\n")
        self._data_text.insert(tk.END, f"  Race:      {race}\n")
        self._data_text.insert(tk.END, f"  Ethnicity: {ethnicity}\n")
        self._data_text.insert(tk.END, f"  Language:  {language}\n")

        # Show extraction time if available
        extract_secs = patient_data.get('_extract_seconds')
        if extract_secs is not None:
            self._data_text.insert(tk.END, f"  Extracted in {extract_secs}s\n")

        self._data_text.see(tk.END)
        self._data_text.config(state=tk.DISABLED)

        # Update count
        if self._count_label:
            count = len(self._patients)
            self._count_label.config(text=f"{count} patient{'s' if count != 1 else ''}")

    def stop(self):
        """Stop the overlay."""
        self._running = False
        if self.root:
            try:
                self.root.after(0, self._close)
            except:
                pass

    def _close(self):
        try:
            if self.panel:
                self.panel.destroy()
            if self.root:
                self.root.quit()
                self.root.destroy()
        except:
            pass


# Global instance for singleton access
_control_overlay: Optional[ControlOverlay] = None
_mode_selection: Optional[ModeSelectionOverlay] = None
_analytics_dashboard: Optional[AnalyticsDashboardOverlay] = None

# Global registry for all overlay windows (for hide/show all functionality)
_all_overlay_windows: List[tk.Toplevel] = []

def register_overlay_window(window: tk.Toplevel):
    """Register an overlay window for global hide/show."""
    global _all_overlay_windows
    if window not in _all_overlay_windows:
        _all_overlay_windows.append(window)

def unregister_overlay_window(window: tk.Toplevel):
    """Unregister an overlay window."""
    global _all_overlay_windows
    if window in _all_overlay_windows:
        _all_overlay_windows.remove(window)

def hide_all_overlays():
    """Hide all registered overlay windows."""
    global _all_overlay_windows, _control_overlay, _analytics_dashboard
    # Hide singleton overlays
    if _control_overlay:
        _control_overlay.hide()
    if _analytics_dashboard:
        _analytics_dashboard.hide()
    # Hide all registered windows
    for window in _all_overlay_windows:
        try:
            if window and window.winfo_exists():
                window.withdraw()
        except:
            pass

def show_all_overlays():
    """Show all registered overlay windows."""
    global _all_overlay_windows, _control_overlay, _analytics_dashboard
    # Show singleton overlays
    if _control_overlay:
        _control_overlay.show()
    if _analytics_dashboard:
        _analytics_dashboard.show()
    # Show all registered windows
    for window in _all_overlay_windows:
        try:
            if window and window.winfo_exists():
                window.deiconify()
                window.lift()
                window.attributes("-topmost", True)
        except:
            pass


def get_control_overlay() -> ControlOverlay:
    """Get or create the control overlay singleton."""
    global _control_overlay
    if _control_overlay is None:
        _control_overlay = ControlOverlay()
    return _control_overlay


def get_mode_selection() -> ModeSelectionOverlay:
    """Get or create the mode selection overlay."""
    global _mode_selection
    if _mode_selection is None:
        _mode_selection = ModeSelectionOverlay()
    return _mode_selection


def reset_control_overlay():
    """Reset the control overlay singleton."""
    global _control_overlay
    if _control_overlay:
        _control_overlay.stop()
    _control_overlay = None


def reset_mode_selection():
    """Reset the mode selection singleton."""
    global _mode_selection
    if _mode_selection:
        _mode_selection.stop()
    _mode_selection = None


def get_analytics_dashboard() -> AnalyticsDashboardOverlay:
    """Get or create the analytics dashboard singleton."""
    global _analytics_dashboard
    if _analytics_dashboard is None:
        _analytics_dashboard = AnalyticsDashboardOverlay()
    return _analytics_dashboard


def reset_analytics_dashboard():
    """Reset the analytics dashboard singleton."""
    global _analytics_dashboard
    if _analytics_dashboard:
        _analytics_dashboard.stop()
    _analytics_dashboard = None
