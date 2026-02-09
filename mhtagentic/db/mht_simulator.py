"""
MHT Response Simulator

Simulates MHT API responses by:
1. Monitoring for new inbound (I) events with status >= 1 (CONVERTED)
2. After a delay (~30 seconds), creates an outbound (O) event with mock assessment results
3. Sets the outbound event to status=10 (ready to be processed)

This mimics the production flow where MHT processes assessments and sends results back.
"""

import sqlite3
import json
import time
import threading
import random
import base64
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, List
import logging

logger = logging.getLogger("mhtagentic.mht_simulator")


class MHTResponseSimulator:
    """Simulates MHT sending back assessment results."""

    # Status codes matching production
    STATUS_INITIAL = 0
    STATUS_OUTBOUND_READY = 100  # Outbound results ready to be processed by OutboundWorker
    STATUS_COMPLETE = 200

    def __init__(self, db_path: str, response_delay_seconds: int = 30):
        """
        Initialize the simulator.

        Args:
            db_path: Path to SQLite database
            response_delay_seconds: Delay before generating mock response (default 30s)
        """
        self.db_path = Path(db_path)
        self.response_delay = response_delay_seconds
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._processed_event_ids: set = set()  # Track which inbound events we've responded to

    def _get_connection(self) -> sqlite3.Connection:
        """Get database connection."""
        conn = sqlite3.connect(str(self.db_path))
        conn.row_factory = sqlite3.Row
        return conn

    def _generate_mock_assessment(self, patient_data: dict) -> dict:
        """
        Generate mock MHT assessment results.

        Args:
            patient_data: The converted patient data from inbound event

        Returns:
            Mock assessment response matching MHT format
        """
        patient = patient_data.get('patient', {})
        patient_id = patient.get('patient_id', 'unknown')
        first_name = patient.get('patient_first_name', '')
        last_name = patient.get('patient_last_name', '')

        # Generate random but realistic PHQ-9 scores (0-3 each)
        phq9_scores = [random.randint(0, 3) for _ in range(9)]
        total_score = sum(phq9_scores)

        # Determine severity based on total score
        if total_score <= 4:
            severity = "Minimal"
        elif total_score <= 9:
            severity = "Mild"
        elif total_score <= 14:
            severity = "Moderate"
        elif total_score <= 19:
            severity = "Moderately Severe"
        else:
            severity = "Severe"

        # Build assessment items array (PHQ-9 questions)
        phq9_questions = [
            "Little interest or pleasure in doing things",
            "Feeling down, depressed, or hopeless",
            "Trouble falling/staying asleep or sleeping too much",
            "Feeling tired or having little energy",
            "Poor appetite or overeating",
            "Feeling bad about yourself",
            "Trouble concentrating on things",
            "Moving or speaking slowly/being fidgety",
            "Thoughts of self-harm"
        ]

        assessment_items = []
        for i, (question, score) in enumerate(zip(phq9_questions, phq9_scores), 1):
            assessment_items.append({
                "assessment_item_id": 40 + i,
                "assessment_item_name": f"PHQ-9 Q{i}",
                "assessment_item_description": question,
                "assessment_item_score": f"{score}.0",
                "assessment_item_score_range": "0.00 - 3.00",
                "assessment_item_legend_value": ["Not at all", "Several days", "More than half", "Nearly every day"][score],
                "assessment_item_asset": f"PHQ9_{i}",
                "assessment_item_type": "Functional Impairment"
            })

        # Mock PDF URL (in production this would be a real S3/blob URL)
        mock_pdf_url = f"https://mht-assessments.s3.amazonaws.com/reports/{patient_id}_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"

        # Build the full response matching MHT format
        assessment_response = {
            "data": {
                "Status": "Success",
                "clinic": {
                    "protocol": "API",
                    "clinic_id": patient_data.get('clinic_id', 110),
                    "clinic_name": "Southern Immediate Care",
                    "clinic_location": "SOUTHERN IMMEDIATE CARE - ATTALLA"
                },
                "patient": {
                    "patient_id": patient_id,
                    "patient_first_name": first_name,
                    "patient_last_name": last_name,
                    "patient_mobile": patient.get('patient_mobile', ''),
                    "patient_email": patient.get('patient_email', '')
                },
                "assessment_response_url": mock_pdf_url,
                "assessment_preferred_language": "EN",
                "assessment_completed_at": datetime.now().isoformat(),
                "assessment": [
                    {
                        "assessment_id": random.randint(1000000, 9999999),
                        "assessment_name": "PHQ-9",
                        "assessment_type": "Depression Screening",
                        "assessment_items": assessment_items,
                        "total_score_legend": severity,
                        "total_score_value": str(total_score),
                        "assessment_overall_legend_score": f"{total_score}.0",
                        "patient_score_legend_value": severity,
                        "assessment_clinical_notes": f"Patient {first_name} {last_name} completed PHQ-9 screening. Total score: {total_score} ({severity}).",
                        "flagged_abnormal": total_score >= 10
                    }
                ],
                "encounter_id": f"ENC_{patient_id}_{datetime.now().strftime('%Y%m%d')}",
                "current_time": datetime.now().isoformat()
            },
            "_metadata": {
                "simulated": True,
                "simulated_at": datetime.now().isoformat(),
                "source_event_patient": f"{last_name}, {first_name}",
                "response_delay_seconds": self.response_delay
            }
        }

        return assessment_response

    def _create_outbound_event(self, inbound_event_id: int, inbound_converted_data: dict) -> int:
        """
        Create an outbound (O) event with mock assessment results.

        Args:
            inbound_event_id: The source inbound event ID
            inbound_converted_data: The converted data from inbound event

        Returns:
            New outbound event ID
        """
        conn = self._get_connection()
        try:
            cursor = conn.cursor()
            now = datetime.now().isoformat()

            # Generate mock assessment response
            mock_response = self._generate_mock_assessment(inbound_converted_data)

            # Create outbound event
            cursor.execute("""
                INSERT INTO common_event (
                    received_at, direction, raw_data, status, kind, updated_at, error_count
                ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                now,
                'O',  # Outbound - from MHT to us
                json.dumps(mock_response),
                self.STATUS_OUTBOUND_READY,  # status=10, ready to be processed
                'assessment_result',
                now,
                0
            ))

            conn.commit()
            event_id = cursor.lastrowid

            logger.info(f"Created outbound event {event_id} for inbound event {inbound_event_id}")
            return event_id

        finally:
            conn.close()

    def _check_and_respond(self):
        """Check for new inbound events and create responses after delay."""
        conn = self._get_connection()
        try:
            cursor = conn.cursor()

            # Find inbound events with status >= 1 (CONVERTED or higher) that we haven't responded to
            cursor.execute("""
                SELECT id, converted_data, received_at
                FROM common_event
                WHERE direction = 'I' AND status >= 1 AND converted_data IS NOT NULL
                ORDER BY id
            """)

            for row in cursor.fetchall():
                event_id = row['id']

                # Skip if already processed
                if event_id in self._processed_event_ids:
                    continue

                # Check if enough time has passed since the event was received
                received_at = datetime.fromisoformat(row['received_at'])
                elapsed = (datetime.now() - received_at).total_seconds()

                if elapsed >= self.response_delay:
                    # Time to create the response
                    converted_data = json.loads(row['converted_data'])
                    self._create_outbound_event(event_id, converted_data)
                    self._processed_event_ids.add(event_id)

                    patient = converted_data.get('patient', {})
                    logger.info(f"Generated mock MHT response for patient: "
                               f"{patient.get('patient_first_name', '')} {patient.get('patient_last_name', '')}")

        finally:
            conn.close()

    def _run_loop(self):
        """Main simulation loop."""
        logger.info(f"MHT Simulator started (response delay: {self.response_delay}s)")

        while self._running:
            try:
                self._check_and_respond()
            except Exception as e:
                logger.error(f"Simulator error: {e}")

            # Check every 5 seconds
            time.sleep(5)

        logger.info("MHT Simulator stopped")

    def start(self):
        """Start the simulator in a background thread."""
        if self._running:
            return

        self._running = True
        self._thread = threading.Thread(target=self._run_loop, daemon=True)
        self._thread.start()
        logger.info("MHT Response Simulator started")

    def stop(self):
        """Stop the simulator."""
        self._running = False
        if self._thread:
            self._thread.join(timeout=10)
        logger.info("MHT Response Simulator stopped")

    def is_running(self) -> bool:
        """Check if simulator is running."""
        return self._running


# Standalone usage
if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(message)s')

    db_path = sys.argv[1] if len(sys.argv) > 1 else "output/mht_data.db"
    delay = int(sys.argv[2]) if len(sys.argv) > 2 else 30

    simulator = MHTResponseSimulator(db_path, response_delay_seconds=delay)
    simulator.start()

    print(f"MHT Simulator running (delay: {delay}s). Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        simulator.stop()
