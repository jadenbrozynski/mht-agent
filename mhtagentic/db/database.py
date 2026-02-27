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
            "SELECT slot_name, session_id, claimed_at, heartbeat_at, status "
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
