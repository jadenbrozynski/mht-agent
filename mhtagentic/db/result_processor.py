"""
MHT Result Processor

Background script that:
1. Polls for outbound (O) events with status=10 (unprocessed results from MHT)
2. Processes the assessment data
3. Updates status to 100 (complete)

This runs alongside the main automation to handle incoming MHT results.
"""

import sqlite3
import json
import time
import threading
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List, Callable
import logging

logger = logging.getLogger("mhtagentic.result_processor")


class MHTResultProcessor:
    """Processes MHT assessment results from outbound events."""

    # Status codes
    STATUS_OUTBOUND_READY = 10  # Ready to process
    STATUS_PROCESSING = 50      # Being processed
    STATUS_COMPLETE = 100       # Done
    MAX_ERRORS = 4              # Max retries before marking as failed

    def __init__(self, db_path: str, poll_interval_seconds: int = 10):
        """
        Initialize the result processor.

        Args:
            db_path: Path to SQLite database
            poll_interval_seconds: How often to check for new results
        """
        self.db_path = Path(db_path)
        self.poll_interval = poll_interval_seconds
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._on_result_processed: Optional[Callable[[dict], None]] = None

    def set_callback(self, callback: Callable[[dict], None]):
        """Set callback to be called when a result is processed."""
        self._on_result_processed = callback

    def _get_connection(self) -> sqlite3.Connection:
        """Get database connection."""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _get_pending_results(self) -> List[dict]:
        """
        Query for unprocessed outbound results.

        SELECT * FROM common_event WHERE direction='O' AND status=10
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT id, received_at, raw_data, status, kind, error_count
                FROM common_event
                WHERE direction = 'O' AND status = ?
                ORDER BY received_at
            """, (self.STATUS_OUTBOUND_READY,))

            return [dict(row) for row in cursor.fetchall()]
        finally:
            conn.close()

    def _update_status(self, event_id: int, status: int, converted_data: dict = None, response_data: dict = None):
        """Update event status and optional data fields."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            now = datetime.now().isoformat()

            if converted_data and response_data:
                cursor.execute("""
                    UPDATE common_event
                    SET status = ?, converted_at = ?, converted_data = ?,
                        sent_at = ?, response_data = ?, updated_at = ?
                    WHERE id = ?
                """, (status, now, json.dumps(converted_data), now, json.dumps(response_data), now, event_id))
            else:
                cursor.execute("""
                    UPDATE common_event
                    SET status = ?, updated_at = ?
                    WHERE id = ?
                """, (status, now, event_id))

            conn.commit()
            return cursor.rowcount > 0
        finally:
            conn.close()

    def _increment_error(self, event_id: int, current_status: int, error_message: str) -> bool:
        """
        Increment error count. If max errors reached, negate status to mark as failed.

        Returns True if event was marked as failed.
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            now = datetime.now().isoformat()

            # Get current error count
            cursor.execute("SELECT error_count FROM common_event WHERE id = ?", (event_id,))
            row = cursor.fetchone()
            if not row:
                return False

            new_error_count = row['error_count'] + 1

            if new_error_count >= self.MAX_ERRORS:
                # Mark as failed by negating status
                new_status = -abs(current_status)
                cursor.execute("""
                    UPDATE common_event
                    SET status = ?, error_count = ?, updated_at = ?
                    WHERE id = ?
                """, (new_status, new_error_count, now, event_id))
                logger.warning(f"Event {event_id} marked as FAILED (status={new_status}) after {new_error_count} errors")
                conn.commit()
                return True
            else:
                cursor.execute("""
                    UPDATE common_event
                    SET error_count = ?, updated_at = ?
                    WHERE id = ?
                """, (new_error_count, now, event_id))
                logger.info(f"Event {event_id} error count incremented to {new_error_count}")
                conn.commit()
                return False
        finally:
            conn.close()

    def _process_assessment_result(self, event: dict) -> dict:
        """
        Process an assessment result and extract key data.

        Args:
            event: The outbound event dict

        Returns:
            Processed result summary
        """
        raw_data = json.loads(event['raw_data']) if event['raw_data'] else {}
        data = raw_data.get('data', raw_data)

        # Extract assessment info
        patient = data.get('patient', {})
        assessments = data.get('assessment', [])

        result_summary = {
            "event_id": event['id'],
            "processed_at": datetime.now().isoformat(),
            "patient_id": patient.get('patient_id'),
            "patient_name": f"{patient.get('patient_first_name', '')} {patient.get('patient_last_name', '')}".strip(),
            "assessments": []
        }

        for assessment in assessments:
            assessment_summary = {
                "assessment_id": assessment.get('assessment_id'),
                "assessment_name": assessment.get('assessment_name'),
                "assessment_type": assessment.get('assessment_type'),
                "total_score": assessment.get('total_score_value'),
                "severity": assessment.get('total_score_legend') or assessment.get('patient_score_legend_value'),
                "flagged_abnormal": assessment.get('flagged_abnormal', False),
                "clinical_notes": assessment.get('assessment_clinical_notes'),
                "item_count": len(assessment.get('assessment_items', []))
            }
            result_summary["assessments"].append(assessment_summary)

        # Extract PDF URL if present
        result_summary["pdf_url"] = data.get('assessment_response_url')

        return result_summary

    def _process_single_result(self, event: dict) -> bool:
        """
        Process a single outbound result.

        Args:
            event: The event dict

        Returns:
            True if successfully processed
        """
        event_id = event['id']
        logger.info(f"Processing outbound event {event_id}...")

        try:
            # Update to processing status
            self._update_status(event_id, self.STATUS_PROCESSING)

            # Process the assessment data
            result_summary = self._process_assessment_result(event)

            # Create response data (simulating what we'd store after processing)
            response_data = {
                "processed": True,
                "processed_at": datetime.now().isoformat(),
                "summary": result_summary,
                "status": "success"
            }

            # Update to complete
            self._update_status(
                event_id,
                self.STATUS_COMPLETE,
                converted_data=result_summary,
                response_data=response_data
            )

            logger.info(f"Event {event_id} processed successfully - "
                       f"Patient: {result_summary['patient_name']}, "
                       f"Assessments: {len(result_summary['assessments'])}")

            # Call callback if set
            if self._on_result_processed:
                self._on_result_processed(result_summary)

            return True

        except Exception as e:
            logger.error(f"Error processing event {event_id}: {e}")
            self._increment_error(event_id, self.STATUS_OUTBOUND_READY, str(e))
            return False

    def _poll_and_process(self):
        """Poll for pending results and process them."""
        pending = self._get_pending_results()

        if pending:
            logger.info(f"Found {len(pending)} pending outbound results to process")

        for event in pending:
            self._process_single_result(event)

    def _run_loop(self):
        """Main processing loop."""
        logger.info(f"Result Processor started (poll interval: {self.poll_interval}s)")

        while self._running:
            try:
                self._poll_and_process()
            except Exception as e:
                logger.error(f"Processor error: {e}")

            time.sleep(self.poll_interval)

        logger.info("Result Processor stopped")

    def start(self):
        """Start the processor in a background thread."""
        if self._running:
            return

        self._running = True
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()
        logger.info("MHT Result Processor started")

    def stop(self):
        """Stop the processor."""
        self._running = False
        if self._thread:
            self._thread.join(timeout=10)
        logger.info("MHT Result Processor stopped")

    def is_running(self) -> bool:
        """Check if processor is running."""
        return self._running

    def process_now(self) -> int:
        """
        Manually trigger processing of all pending results.

        Returns:
            Number of results processed
        """
        pending = self._get_pending_results()
        processed = 0
        for event in pending:
            if self._process_single_result(event):
                processed += 1
        return processed

    def get_stats(self) -> dict:
        """Get processing statistics."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            stats = {
                "outbound_pending": 0,
                "outbound_complete": 0,
                "outbound_failed": 0,
                "inbound_total": 0,
                "inbound_complete": 0
            }

            # Outbound stats
            cursor.execute("SELECT COUNT(*) FROM common_event WHERE direction='O' AND status=10")
            stats["outbound_pending"] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM common_event WHERE direction='O' AND status=100")
            stats["outbound_complete"] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM common_event WHERE direction='O' AND status<0")
            stats["outbound_failed"] = cursor.fetchone()[0]

            # Inbound stats
            cursor.execute("SELECT COUNT(*) FROM common_event WHERE direction='I'")
            stats["inbound_total"] = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM common_event WHERE direction='I' AND status>=1")
            stats["inbound_complete"] = cursor.fetchone()[0]

            return stats
        finally:
            conn.close()


# Standalone usage
if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(message)s')

    db_path = sys.argv[1] if len(sys.argv) > 1 else "output/mht_data.db"
    interval = int(sys.argv[2]) if len(sys.argv) > 2 else 10

    processor = MHTResultProcessor(db_path, poll_interval_seconds=interval)

    def on_result(result):
        print(f"\n=== RESULT PROCESSED ===")
        print(f"Patient: {result['patient_name']}")
        for a in result['assessments']:
            print(f"  {a['assessment_name']}: {a['total_score']} ({a['severity']})")
        print()

    processor.set_callback(on_result)
    processor.start()

    print(f"Result Processor running (interval: {interval}s). Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        processor.stop()
