"""
Integration Client for remote MHT MySQL server.

Pushes inbound patient events to the integration server's common_event table
so MHT's SmarTest system can process real assessments.

Format: MHT 'modify appointment' API spec (nested patient/clinician/appointment).
Status: 20 = CONVERTED (go direct since we convert in-house).
Kind: 'update' (required for MHT to process the event).
"""

import json
import logging
from datetime import datetime, timedelta

logger = logging.getLogger("mhtagentic.integration")


class IntegrationClient:
    """MySQL client for dual-writing events to the MHT integration server."""

    def __init__(self, host, port, user, password, database,
                 clinic_id=163, npi=""):
        self.config = {
            "host": host,
            "port": port,
            "user": user,
            "password": password,
            "database": database,
        }
        self.clinic_id = clinic_id
        self.npi = npi

    def _connect(self):
        """Get a pymysql connection. Returns None on failure."""
        import pymysql
        try:
            return pymysql.connect(**self.config, connect_timeout=5)
        except Exception as e:
            logger.warning(f"Integration server connection failed: {e}")
            return None

    def build_mht_payload(self, mht_payload: dict) -> dict:
        """Build the MHT modify-appointment payload from our internal format.

        Produces the nested structure that MHT SmarTest expects:
        {clinic_id, patient: {...}, clinician: {...}, appointment: {...}}

        Only includes fields from the modify-appointment API spec.
        """
        patient = mht_payload.get('patient', {})
        clinician = mht_payload.get('clinician', {})
        appointment = mht_payload.get('appointment', {})

        return {
            "clinic_id": self.clinic_id,
            "patient": {
                "patient_id": patient.get('patient_id', ''),
                "patient_first_name": patient.get('patient_first_name', ''),
                "patient_last_name": patient.get('patient_last_name', ''),
                "patient_date_of_birth": patient.get('patient_date_of_birth', ''),
                "patient_sex": patient.get('patient_sex', ''),
                "patient_email": patient.get('patient_email', ''),
                "patient_mobile": patient.get('patient_mobile', ''),
                "patient_race": patient.get('patient_race', ''),
                "patient_ethnicity": patient.get('patient_ethnicity', ''),
                "patient_preferred_language": patient.get('patient_preferred_language', 'English'),
            },
            "clinician": {
                "clinician_alternate_id": clinician.get('clinician_alternate_id', '') or self.npi,
                "clinician_first_name": clinician.get('clinician_first_name', ''),
                "clinician_last_name": clinician.get('clinician_last_name', ''),
                "clinician_email": clinician.get('clinician_email', ''),
            },
            "appointment": {
                "appointment_id": appointment.get('appointment_id', '')
                                  or f"APPT_{datetime.now().strftime('%m%d%H%M%S')}",
                "clinic_location": "",  # Must be empty — MHT silently skips delivery otherwise
                "appointment_date": (datetime.now() + timedelta(days=1)).replace(
                    hour=10, minute=0, second=0, microsecond=0
                ).strftime("%Y-%m-%dT%H:%M:%S"),
                "appointment_reason": appointment.get('appointment_reason', '')
                                      or "Behavioral Health Screening",
            },
        }

    # Keep convert_payload as alias for backwards compatibility
    def convert_payload(self, mht_payload: dict) -> dict:
        """Alias for build_mht_payload (backwards compat)."""
        return self.build_mht_payload(mht_payload)

    def push_event(self, mht_payload: dict) -> bool:
        """
        Insert an inbound event into the remote common_event table.

        Uses MHT modify-appointment nested format.
        Status = 20 (CONVERTED — go direct since we convert in-house).
        Kind = 'update' (required for MHT to process).

        Returns True on success, False on failure.
        Never raises -- failures are logged but don't block the local flow.
        """
        try:
            payload = self.build_mht_payload(mht_payload)
            payload_json = json.dumps(payload)

            conn = self._connect()
            if not conn:
                return False

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with conn.cursor() as cur:
                cur.execute("""
                    INSERT INTO common_event
                    (received_at, direction, raw_data, converted_at,
                     converted_data, status, kind, updated_at, error_count)
                    VALUES (%s, 'I', %s, %s, %s, 20, 'update', %s, 0)
                """, (now, payload_json, now, payload_json, now))
            conn.commit()
            conn.close()

            fn = payload.get('patient', {}).get('patient_first_name', '')
            ln = payload.get('patient', {}).get('patient_last_name', '')
            logger.info(
                f"Pushed event to integration server for "
                f"{fn} {ln} (status=20, kind=update)"
            )
            return True
        except Exception as e:
            logger.warning(f"Integration push failed: {e}")
            return False

    def update_outbound_status(self, event_id: int, status: int) -> bool:
        """
        Update the status of an outbound event on the remote integration server.

        Fire-and-forget: never blocks the main flow. Failures are logged only.

        Args:
            event_id: Remote event ID to update.
            status: New status code (granular step status).

        Returns:
            True on success, False on failure.
        """
        try:
            conn = self._connect()
            if not conn:
                return False

            now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            with conn.cursor() as cur:
                cur.execute(
                    "UPDATE common_event SET status=%s, updated_at=%s WHERE id=%s",
                    (status, now, event_id),
                )
            conn.commit()
            conn.close()
            logger.info(f"Updated remote outbound event {event_id} status → {status}")
            return True
        except Exception as e:
            logger.warning(f"Remote outbound status update failed for event {event_id}: {e}")
            return False
