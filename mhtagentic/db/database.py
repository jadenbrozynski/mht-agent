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

            # Update event error count and status
            cursor.execute("""
                UPDATE common_event
                SET error_count = error_count + 1, status = ?, updated_at = ?
                WHERE id = ?
            """, (EventStatus.ERROR, now, event_id))

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
