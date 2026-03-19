"""
MHT Agentic Database Module

SQLite database for tracking patient extraction events with full audit trail.
"""

import sqlite3
import json
from datetime import datetime
from enum import IntEnum
from pathlib import Path
from typing import Dict, List, Optional, Union


class EventStatus(IntEnum):
    """
    Status codes for event lifecycle (0-100 scale matching production).

    Inbound (I) - Automation sending TO MHT:
        0   = Initial/created
        10  = Converted to MHT format
        20  = Ready to send
        40  = Sent to MHT API
        100 = Complete

    Outbound (O) - MHT sending results TO us:
        10  = Results received, ready to process
        50  = Processing
        100 = Complete

    Negative values indicate failure at that stage:
        -10 = Failed at conversion
        -40 = Failed at API send
        etc.

    error_count tracks retries (max 4 before marking as failed)
    """
    # Inbound statuses
    INITIAL = 0           # Raw data received
    CONVERTED = 10        # Converted to MHT API format
    READY_TO_SEND = 20    # Queued for API
    SENT = 40             # Sent to MHT API

    # Outbound statuses
    OUTBOUND_READY = 10   # MHT results ready to process
    PROCESSING = 50       # Being processed

    # Terminal statuses
    COMPLETE = 100        # Successfully completed

    # Legacy compatibility (maps old values)
    PENDING = 0           # Alias for INITIAL
    EXPIRED = 100         # Patient discharged = complete for our purposes


class MHTDatabase:
    """SQLite database manager for MHT patient extraction events."""

    def __init__(self, db_path: Union[str, Path]):
        """
        Initialize database connection and create tables if needed.

        Args:
            db_path: Path to SQLite database file
        """
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        """Get a database connection with row factory."""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        """Create database tables if they don't exist."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            # Create common_event table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS common_event (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    received_at DATETIME NOT NULL,
                    direction TEXT NOT NULL CHECK(direction IN ('I', 'O')),
                    raw_data TEXT,
                    converted_at DATETIME,
                    converted_data TEXT,
                    sent_at DATETIME,
                    response_data TEXT,
                    status INTEGER NOT NULL DEFAULT 0,
                    kind TEXT NOT NULL,
                    updated_at DATETIME,
                    error_count INTEGER NOT NULL DEFAULT 0
                )
            """)

            # Create common_eventerror table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS common_eventerror (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at DATETIME NOT NULL,
                    error TEXT NOT NULL,
                    event_id INTEGER NOT NULL,
                    FOREIGN KEY (event_id) REFERENCES common_event(id)
                )
            """)

            # Create bot_slot table for dynamic inbound/outbound assignment
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS bot_slot (
                    slot_name TEXT PRIMARY KEY,
                    session_id INTEGER,
                    claimed_at TEXT,
                    heartbeat_at TEXT,
                    status TEXT DEFAULT 'open'
                )
            """)

            # Create bot_error_log table for tracking login/startup errors per bot
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS bot_error_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    slot_name TEXT NOT NULL,
                    attempt INTEGER NOT NULL,
                    max_attempts INTEGER NOT NULL DEFAULT 5,
                    error TEXT NOT NULL,
                    traceback TEXT,
                    step TEXT,
                    created_at DATETIME NOT NULL
                )
            """)

            # Migrate old role-based slots to user-based slots
            cursor.execute("DELETE FROM bot_slot WHERE slot_name IN ('inbound', 'outbound')")

            # Seed per-user slot rows
            cursor.execute("INSERT OR IGNORE INTO bot_slot (slot_name, status) VALUES ('experityb', 'open')")
            cursor.execute("INSERT OR IGNORE INTO bot_slot (slot_name, status) VALUES ('experityc', 'open')")
            cursor.execute("INSERT OR IGNORE INTO bot_slot (slot_name, status) VALUES ('experityd', 'open')")

            # Add role and location columns to bot_slot (migration-safe)
            existing_cols = {
                row[1] for row in cursor.execute("PRAGMA table_info(bot_slot)").fetchall()
            }
            if "role" not in existing_cols:
                cursor.execute("ALTER TABLE bot_slot ADD COLUMN role TEXT DEFAULT ''")
            if "location" not in existing_cols:
                cursor.execute("ALTER TABLE bot_slot ADD COLUMN location TEXT DEFAULT ''")

            # OTP queue table — coordinates TOTP code usage across RDP sessions
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS otp_queue (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    bot_name TEXT NOT NULL,
                    requested_at TEXT NOT NULL,
                    granted_at TEXT,
                    completed_at TEXT,
                    otp_code TEXT,
                    success INTEGER,
                    totp_period INTEGER
                )
            """)

            # Typing lock table — serializes keyboard/clipboard input across RDP sessions
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS typing_lock (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    holder TEXT,
                    acquired_at TEXT
                )
            """)
            cursor.execute("INSERT OR IGNORE INTO typing_lock (id) VALUES (1)")

            # Clinic locations table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS clinic_locations (
                    location_name TEXT PRIMARY KEY,
                    display_order INTEGER NOT NULL DEFAULT 0,
                    is_default INTEGER NOT NULL DEFAULT 0,
                    assigned_bot TEXT,
                    assigned_at TEXT,
                    bot_role TEXT
                )
            """)
            # Add is_active column (migration-safe)
            loc_cols = {
                row[1] for row in cursor.execute("PRAGMA table_info(clinic_locations)").fetchall()
            }
            if "is_active" not in loc_cols:
                cursor.execute("ALTER TABLE clinic_locations ADD COLUMN is_active INTEGER NOT NULL DEFAULT 1")

            # Seed default locations
            cursor.execute(
                "INSERT OR IGNORE INTO clinic_locations "
                "(location_name, display_order, is_default) VALUES ('ANNISTON', 1, 1)"
            )
            cursor.execute(
                "INSERT OR IGNORE INTO clinic_locations "
                "(location_name, display_order, is_default) VALUES ('ATTALLA', 2, 0)"
            )

            # Bot config table — dashboard-configurable key-value store
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS bot_config (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
            """)
            # Seed defaults
            cursor.execute("INSERT OR IGNORE INTO bot_config (key, value) VALUES ('rdp_count', '3')")
            cursor.execute("INSERT OR IGNORE INTO bot_config (key, value) VALUES ('inbound_count', '2')")
            cursor.execute("INSERT OR IGNORE INTO bot_config (key, value) VALUES ('outbound_count', '1')")

            conn.commit()
        finally:
            conn.close()

    def create_event(self, raw_data: dict, kind: str = "patient_extraction") -> int:
        """
        Create a new event with raw data.

        Args:
            raw_data: Dictionary of raw scraped patient data
            kind: Event type identifier

        Returns:
            Event ID
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            now = datetime.now().isoformat()

            cursor.execute("""
                INSERT INTO common_event (received_at, direction, raw_data, status, kind, updated_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (now, 'I', json.dumps(raw_data), EventStatus.PENDING, kind, now))

            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()

    def update_event_converted(self, event_id: int, converted_data: dict) -> bool:
        """
        Update event with converted MHT API format data.

        Args:
            event_id: Event ID to update
            converted_data: MHT API payload dictionary

        Returns:
            True if update successful
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            now = datetime.now().isoformat()

            cursor.execute("""
                UPDATE common_event
                SET converted_at = ?, converted_data = ?, status = ?, updated_at = ?
                WHERE id = ?
            """, (now, json.dumps(converted_data), EventStatus.CONVERTED, now, event_id))  # status=10

            conn.commit()
            return cursor.rowcount > 0
        finally:
            conn.close()

    def update_event_sent(self, event_id: int, response_data: dict) -> bool:
        """
        Update event after sending to MHT API.

        Args:
            event_id: Event ID to update
            response_data: API response dictionary

        Returns:
            True if update successful
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            now = datetime.now().isoformat()

            cursor.execute("""
                UPDATE common_event
                SET sent_at = ?, response_data = ?, status = ?, updated_at = ?
                WHERE id = ?
            """, (now, json.dumps(response_data), EventStatus.SENT, now, event_id))

            conn.commit()
            return cursor.rowcount > 0
        finally:
            conn.close()

    def expire_event(self, event_id: int) -> bool:
        """
        Mark event as expired (patient discharged).

        Args:
            event_id: Event ID to expire

        Returns:
            True if update successful
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            now = datetime.now().isoformat()

            cursor.execute("""
                UPDATE common_event
                SET status = ?, updated_at = ?
                WHERE id = ?
            """, (EventStatus.EXPIRED, now, event_id))

            conn.commit()
            return cursor.rowcount > 0
        finally:
            conn.close()

    def record_error(self, event_id: int, error_message: str) -> int:
        """
        Record an error for an event and increment error count.

        Args:
            event_id: Event ID with error
            error_message: Error description

        Returns:
            Error record ID
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            now = datetime.now().isoformat()

            # Insert error record
            cursor.execute("""
                INSERT INTO common_eventerror (created_at, error, event_id)
                VALUES (?, ?, ?)
            """, (now, error_message, event_id))

            error_id = cursor.lastrowid

            # Increment error count
            cursor.execute("""
                UPDATE common_event
                SET error_count = error_count + 1, updated_at = ?
                WHERE id = ?
            """, (now, event_id))

            conn.commit()
            return error_id
        finally:
            conn.close()

    def get_event(self, event_id: int) -> Optional[dict]:
        """
        Get event by ID.

        Args:
            event_id: Event ID

        Returns:
            Event dict or None
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM common_event WHERE id = ?", (event_id,))
            row = cursor.fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def get_pending_events(self) -> List[Dict]:
        """Get all events with PENDING status."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM common_event WHERE status = ? ORDER BY received_at",
                (EventStatus.PENDING,)
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            conn.close()

    def get_converted_events(self) -> List[Dict]:
        """Get all events with CONVERTED status (ready to send)."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT * FROM common_event WHERE status = ? ORDER BY converted_at",
                (EventStatus.CONVERTED,)
            )
            return [dict(row) for row in cursor.fetchall()]
        finally:
            conn.close()


# ---------------------------------------------------------------------------
# Standalone slot helpers (work with raw db_path, no MHTDatabase instance)
# ---------------------------------------------------------------------------

def claim_slot(db_path: Union[str, Path], preferred: str) -> Optional[str]:
    """
    Atomically claim a specific bot slot by name.

    Args:
        db_path: Path to SQLite database file.
        preferred: The user-based slot name (e.g. 'experityb', 'experityc', 'experityd').

    Returns:
        Slot name if claimed, or None if not available.
    """
    import os
    if not preferred:
        return None
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        now = datetime.now().isoformat()
        session_id = os.getpid()

        conn.execute("BEGIN EXCLUSIVE")

        row = conn.execute(
            "SELECT status FROM bot_slot WHERE slot_name = ?", (preferred,)
        ).fetchone()
        if row and row[0] == "open":
            conn.execute(
                "UPDATE bot_slot SET status = 'active', session_id = ?, "
                "claimed_at = ?, heartbeat_at = ? WHERE slot_name = ?",
                (session_id, now, now, preferred),
            )
            conn.commit()
            return preferred

        conn.commit()
        return None
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return None
    finally:
        conn.close()


def release_slot(db_path: Union[str, Path], slot_name: str) -> None:
    """Release a claimed slot back to 'open'."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        conn.execute(
            "UPDATE bot_slot SET status = 'open', session_id = NULL, "
            "claimed_at = NULL, heartbeat_at = NULL WHERE slot_name = ?",
            (slot_name,),
        )
        conn.commit()
    finally:
        conn.close()


def heartbeat_slot(db_path: Union[str, Path], slot_name: str) -> None:
    """Update the heartbeat timestamp for a claimed slot."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        now = datetime.now().isoformat()
        conn.execute(
            "UPDATE bot_slot SET heartbeat_at = ? WHERE slot_name = ?",
            (now, slot_name),
        )
        conn.commit()
    finally:
        conn.close()


def get_slots(db_path: Union[str, Path]) -> List[Dict]:
    """Return both slot rows for dashboard display."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            "SELECT slot_name, session_id, claimed_at, heartbeat_at, status, "
            "COALESCE(location, '') as location, COALESCE(role, '') as role "
            "FROM bot_slot ORDER BY slot_name"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def cleanup_stale_slots(db_path: Union[str, Path], stale_seconds: int = 90) -> int:
    """
    Reset slots whose heartbeat is older than *stale_seconds* back to 'open'.

    Returns:
        Number of slots cleaned up.
    """
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        cutoff = datetime.now().isoformat()
        # We compare ISO strings — works because they sort lexicographically
        cursor = conn.execute(
            "SELECT slot_name, heartbeat_at FROM bot_slot WHERE status = 'active'"
        )
        cleaned = 0
        for row in cursor.fetchall():
            hb = row[1]
            if not hb:
                continue
            try:
                hb_dt = datetime.fromisoformat(hb)
                age = (datetime.now() - hb_dt).total_seconds()
                if age > stale_seconds:
                    conn.execute(
                        "UPDATE bot_slot SET status = 'open', session_id = NULL, "
                        "claimed_at = NULL, heartbeat_at = NULL WHERE slot_name = ?",
                        (row[0],),
                    )
                    cleaned += 1
            except (ValueError, TypeError):
                pass
        conn.commit()
        return cleaned
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Bot error log helpers
# ---------------------------------------------------------------------------

def log_bot_error(
    db_path: Union[str, Path],
    slot_name: str,
    attempt: int,
    max_attempts: int,
    error: str,
    tb: str = "",
    step: str = "",
) -> int:
    """Log a bot login/startup error attempt to the DB.

    Returns:
        The error log row ID.
    """
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        now = datetime.now().isoformat()
        cur = conn.execute(
            "INSERT INTO bot_error_log "
            "(slot_name, attempt, max_attempts, error, traceback, step, created_at) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (slot_name, attempt, max_attempts, error, tb, step, now),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def get_bot_errors(
    db_path: Union[str, Path], slot_name: str = None, limit: int = 50
) -> List[Dict]:
    """Return recent bot error log entries, optionally filtered by slot.

    Returns:
        List of error log dicts, newest first.
    """
    conn = sqlite3.connect(str(db_path), timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        if slot_name:
            rows = conn.execute(
                "SELECT * FROM bot_error_log WHERE slot_name = ? "
                "ORDER BY created_at DESC LIMIT ?",
                (slot_name, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM bot_error_log ORDER BY created_at DESC LIMIT ?",
                (limit,),
            ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def clear_bot_errors(db_path: Union[str, Path]) -> int:
    """Delete all bot error log entries. Called on stop-all / clear.

    Returns:
        Number of rows deleted.
    """
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        cur = conn.execute("DELETE FROM bot_error_log")
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# OTP queue helpers
# ---------------------------------------------------------------------------

def otp_queue_request(db_path: Union[str, Path], bot_name: str) -> int:
    """Add this bot to the OTP queue. Returns the queue row ID."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        now = datetime.now().isoformat()
        cur = conn.execute(
            "INSERT INTO otp_queue (bot_name, requested_at) VALUES (?, ?)",
            (bot_name, now),
        )
        conn.commit()
        return cur.lastrowid
    finally:
        conn.close()


def otp_queue_try_grant(
    db_path: Union[str, Path], queue_id: int, totp_secret: str
) -> Optional[str]:
    """
    Try to grant an OTP code to this queue entry.

    Uses BEGIN EXCLUSIVE to ensure only one bot holds a grant at a time.
    Returns the TOTP code if granted, or None if not yet this bot's turn.
    """
    import pyotp

    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        conn.execute("BEGIN EXCLUSIVE")

        # Check no other bot is holding an active grant (granted but not completed)
        active = conn.execute(
            "SELECT id FROM otp_queue WHERE granted_at IS NOT NULL AND completed_at IS NULL"
        ).fetchone()
        if active:
            conn.commit()
            return None

        # Check this bot is next in line (lowest ungranted ID)
        next_row = conn.execute(
            "SELECT id FROM otp_queue WHERE granted_at IS NULL ORDER BY id LIMIT 1"
        ).fetchone()
        if not next_row or next_row[0] != queue_id:
            conn.commit()
            return None

        # Generate TOTP code and check current period not already used successfully
        totp = pyotp.TOTP(totp_secret)
        code = totp.now()
        period = int(datetime.now().timestamp()) // 30

        already_used = conn.execute(
            "SELECT id FROM otp_queue WHERE totp_period = ? AND success = 1",
            (period,),
        ).fetchone()
        if already_used:
            conn.commit()
            return None

        # Grant
        now = datetime.now().isoformat()
        conn.execute(
            "UPDATE otp_queue SET granted_at = ?, otp_code = ?, totp_period = ? WHERE id = ?",
            (now, code, period, queue_id),
        )
        conn.commit()
        return code
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return None
    finally:
        conn.close()


def otp_queue_complete(
    db_path: Union[str, Path], queue_id: int, success: bool
) -> None:
    """Mark an OTP queue entry as completed (success or failure)."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        now = datetime.now().isoformat()
        conn.execute(
            "UPDATE otp_queue SET completed_at = ?, success = ? WHERE id = ?",
            (now, 1 if success else 0, queue_id),
        )
        conn.commit()
    finally:
        conn.close()


def otp_queue_clear(db_path: Union[str, Path]) -> int:
    """Delete all OTP queue entries. Called on start-all / stop-all."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        cur = conn.execute("DELETE FROM otp_queue")
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Typing lock helpers — serialize keyboard/clipboard input across RDP sessions
# ---------------------------------------------------------------------------

def typing_lock_acquire(
    db_path: Union[str, Path], bot_name: str, timeout: int = 60
) -> bool:
    """
    Acquire the typing lock, waiting up to *timeout* seconds.

    Uses BEGIN EXCLUSIVE to atomically check and claim the lock.
    Stale locks (acquired_at > 60s ago) are automatically broken.

    Returns:
        True if lock acquired, False on timeout.
    """
    import time as _time
    import sys as _sys

    def _log(msg):
        print(f"[typing_lock] {bot_name}: {msg}", flush=True)
        _sys.stdout.flush()

    _log(f"acquiring lock (timeout={timeout}s) db={db_path}")
    deadline = _time.time() + timeout
    attempts = 0
    while _time.time() < deadline:
        attempts += 1
        conn = sqlite3.connect(str(db_path), timeout=10)
        try:
            conn.execute("BEGIN EXCLUSIVE")

            row = conn.execute(
                "SELECT holder, acquired_at FROM typing_lock WHERE id = 1"
            ).fetchone()

            if row is None:
                # Table exists but row missing — seed it
                conn.execute(
                    "INSERT OR IGNORE INTO typing_lock (id) VALUES (1)"
                )
                conn.execute(
                    "UPDATE typing_lock SET holder = ?, acquired_at = ? WHERE id = 1",
                    (bot_name, datetime.now().isoformat()),
                )
                conn.commit()
                _log(f"ACQUIRED (seeded row) after {attempts} attempts")
                return True

            holder, acquired_at = row[0], row[1]

            # Lock is free
            if holder is None:
                conn.execute(
                    "UPDATE typing_lock SET holder = ?, acquired_at = ? WHERE id = 1",
                    (bot_name, datetime.now().isoformat()),
                )
                conn.commit()
                _log(f"ACQUIRED (was free) after {attempts} attempts")
                return True

            # Check for stale lock (> 180s old — entire login can take 60-90s)
            if acquired_at:
                try:
                    age = (datetime.now() - datetime.fromisoformat(acquired_at)).total_seconds()
                    if age > 180:
                        conn.execute(
                            "UPDATE typing_lock SET holder = ?, acquired_at = ? WHERE id = 1",
                            (bot_name, datetime.now().isoformat()),
                        )
                        conn.commit()
                        _log(f"ACQUIRED (stale from {holder}, age={age:.0f}s) after {attempts} attempts")
                        return True
                except (ValueError, TypeError):
                    pass

            if attempts == 1 or attempts % 10 == 0:
                _log(f"waiting... held by {holder} (attempt {attempts})")

            conn.commit()
        except Exception as exc:
            if attempts == 1:
                _log(f"ERROR on attempt {attempts}: {exc}")
            try:
                conn.rollback()
            except Exception:
                pass
        finally:
            conn.close()

        _time.sleep(0.5)

    _log(f"TIMEOUT after {attempts} attempts ({timeout}s)")
    return False


def typing_lock_release(db_path: Union[str, Path], bot_name: str) -> None:
    """Release the typing lock if held by *bot_name*."""
    print(f"[typing_lock] {bot_name}: RELEASING lock", flush=True)
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        conn.execute(
            "UPDATE typing_lock SET holder = NULL, acquired_at = NULL "
            "WHERE id = 1 AND holder = ?",
            (bot_name,),
        )
        conn.commit()
        print(f"[typing_lock] {bot_name}: RELEASED", flush=True)
    finally:
        conn.close()


def typing_lock_clear(db_path: Union[str, Path]) -> None:
    """Unconditionally reset the typing lock. Called on stop-all / start-all."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        conn.execute(
            "UPDATE typing_lock SET holder = NULL, acquired_at = NULL WHERE id = 1"
        )
        conn.commit()
    finally:
        conn.close()


class _TypingLockContext:
    """Context manager wrapping typing_lock_acquire / typing_lock_release."""

    def __init__(self, db_path: Union[str, Path], bot_name: str, timeout: int = 60):
        self._db_path = str(db_path)
        self._bot_name = bot_name
        self._timeout = timeout
        self._acquired = False

    def __enter__(self):
        self._acquired = typing_lock_acquire(self._db_path, self._bot_name, self._timeout)
        return self._acquired

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._acquired:
            typing_lock_release(self._db_path, self._bot_name)
        return False


def typing_lock_context(
    db_path: Union[str, Path], bot_name: str, timeout: int = 60
) -> _TypingLockContext:
    """Return a context manager that acquires/releases the typing lock."""
    return _TypingLockContext(db_path, bot_name, timeout)


# ---------------------------------------------------------------------------
# Clinic location helpers
# ---------------------------------------------------------------------------

def assign_location(
    db_path: Union[str, Path], bot_name: str, role: str = "inbound"
) -> Optional[str]:
    """
    Atomically assign the next available location to a bot.

    Default location (is_default=1) is assigned first.
    Returns the location name, or None if all locations are taken.
    """
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        now = datetime.now().isoformat()
        conn.execute("BEGIN EXCLUSIVE")

        # Try default location first, then by display_order
        row = conn.execute(
            "SELECT location_name FROM clinic_locations "
            "WHERE assigned_bot IS NULL "
            "ORDER BY is_default DESC, display_order ASC LIMIT 1"
        ).fetchone()

        if not row:
            conn.commit()
            return None

        location = row[0]
        conn.execute(
            "UPDATE clinic_locations SET assigned_bot = ?, assigned_at = ?, bot_role = ? "
            "WHERE location_name = ?",
            (bot_name, now, role, location),
        )
        conn.commit()
        return location
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        return None
    finally:
        conn.close()


def release_location(db_path: Union[str, Path], bot_name: str) -> None:
    """Release all locations assigned to this bot."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        conn.execute(
            "UPDATE clinic_locations SET assigned_bot = NULL, assigned_at = NULL, bot_role = NULL "
            "WHERE assigned_bot = ?",
            (bot_name,),
        )
        conn.commit()
    finally:
        conn.close()


def get_active_locations(db_path: Union[str, Path]) -> List[Dict]:
    """Return locations that are currently assigned to a bot."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            "SELECT * FROM clinic_locations WHERE assigned_bot IS NOT NULL "
            "ORDER BY display_order"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_all_locations(db_path: Union[str, Path]) -> List[Dict]:
    """Return all configured clinic locations."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            "SELECT * FROM clinic_locations ORDER BY display_order"
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def add_location(
    db_path: Union[str, Path], location_name: str, display_order: int = 99
) -> bool:
    """Add a new clinic location. Returns True if added."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        conn.execute(
            "INSERT OR IGNORE INTO clinic_locations (location_name, display_order, is_default) "
            "VALUES (?, ?, 0)",
            (location_name.upper(), display_order),
        )
        conn.commit()
        return conn.total_changes > 0
    finally:
        conn.close()


def remove_location(db_path: Union[str, Path], location_name: str) -> bool:
    """Remove a clinic location (only if not assigned). Returns True if removed."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        cur = conn.execute(
            "DELETE FROM clinic_locations WHERE location_name = ? AND assigned_bot IS NULL",
            (location_name.upper(),),
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def release_all_locations(db_path: Union[str, Path]) -> int:
    """Release all location assignments. Called on stop-all."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        cur = conn.execute(
            "UPDATE clinic_locations SET assigned_bot = NULL, assigned_at = NULL, bot_role = NULL"
        )
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def toggle_location_active(db_path: Union[str, Path], location_name: str, active: bool) -> bool:
    """Toggle a clinic location active/inactive. Returns True if updated."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        val = 1 if active else 0
        cur = conn.execute(
            "UPDATE clinic_locations SET is_active = ?, is_default = ? WHERE location_name = ?",
            (val, val, location_name.upper()),
        )
        conn.commit()
        return cur.rowcount > 0
    finally:
        conn.close()


def sync_bot_slots(db_path: Union[str, Path]) -> None:
    """Sync bot_slot rows to match active clinic locations.

    Ensures there is one bot_slot per active location (using the standard
    experityb/experityc/experityd slot names). Cleans up any stale
    location-named slots.
    """
    conn = sqlite3.connect(str(db_path), timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        active = conn.execute(
            "SELECT location_name FROM clinic_locations WHERE is_active = 1 ORDER BY display_order"
        ).fetchall()
        active_count = len(active)

        # Standard slot names in priority order
        all_slots = ["experityb", "experityc", "experityd"]

        # Remove any non-standard slot names (e.g. old "anniston", "attalla" rows)
        conn.execute(
            "DELETE FROM bot_slot WHERE slot_name NOT IN (?, ?, ?)",
            tuple(all_slots),
        )

        for i, slot in enumerate(all_slots):
            # Ensure the slot row exists
            conn.execute(
                "INSERT OR IGNORE INTO bot_slot (slot_name, status) VALUES (?, 'open')",
                (slot,),
            )
            if i < active_count:
                # Assign location to this slot
                conn.execute(
                    "UPDATE bot_slot SET location = ? WHERE slot_name = ?",
                    (active[i]["location_name"], slot),
                )
            else:
                # Clear slot — no matching active location
                conn.execute(
                    "UPDATE bot_slot SET status = 'open', session_id = NULL, "
                    "claimed_at = NULL, heartbeat_at = NULL, location = '' "
                    "WHERE slot_name = ?",
                    (slot,),
                )

        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Bot config helpers
# ---------------------------------------------------------------------------

def get_config(db_path: Union[str, Path], key: str = None) -> Union[Dict, Optional[str]]:
    """
    Get bot configuration value(s).

    If key is provided, returns the value string (or None).
    If key is None, returns dict of all config key-value pairs.
    """
    conn = sqlite3.connect(str(db_path), timeout=10)
    conn.row_factory = sqlite3.Row
    try:
        if key:
            row = conn.execute(
                "SELECT value FROM bot_config WHERE key = ?", (key,)
            ).fetchone()
            return row[0] if row else None
        else:
            rows = conn.execute("SELECT key, value FROM bot_config").fetchall()
            return {r["key"]: r["value"] for r in rows}
    finally:
        conn.close()


def set_config(db_path: Union[str, Path], key: str, value: str) -> None:
    """Set a bot configuration value (upsert)."""
    conn = sqlite3.connect(str(db_path), timeout=10)
    try:
        conn.execute(
            "INSERT INTO bot_config (key, value) VALUES (?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
            (key, str(value)),
        )
        conn.commit()
    finally:
        conn.close()
