"""
MHT Agentic Analytics Module

Comprehensive analytics and logging for patient processing operations.
Tracks timing, success rates, daily statistics, and software value metrics.
"""

from __future__ import annotations

import json
import time
import logging
import threading
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Optional, Dict, List, Any, Callable
from dataclasses import dataclass, field, asdict
from enum import Enum
import statistics

logger = logging.getLogger("mhtagentic.analytics")


class ProcessingStage(Enum):
    """Stages in patient processing pipeline."""
    WAITING_ROOM_SCAN = "waiting_room_scan"
    PATIENT_CLICK = "patient_click"
    DEMOGRAPHICS_LOAD = "demographics_load"
    DATA_EXTRACTION = "data_extraction"
    POPUP_DISMISSAL = "popup_dismissal"
    DEMOGRAPHICS_CLOSE = "demographics_close"
    CHART_ICON_CLICK = "chart_icon_click"
    INSURANCE_EXTRACTION = "insurance_extraction"
    JSON_GENERATION = "json_generation"
    TOTAL_PROCESSING = "total_processing"


class ProcessingResult(Enum):
    """Result of patient processing."""
    SUCCESS = "success"
    PARTIAL = "partial"  # Some data extracted
    FAILED = "failed"
    SKIPPED = "skipped"  # Patient not qualified
    ERROR = "error"


@dataclass
class PatientProcessingEvent:
    """Individual patient processing event with timing data."""
    patient_name: str
    patient_mrn: str = ""
    start_time: float = 0.0
    end_time: float = 0.0
    result: ProcessingResult = ProcessingResult.SUCCESS
    error_message: str = ""
    stage_timings: Dict[str, float] = field(default_factory=dict)
    fields_extracted: List[str] = field(default_factory=list)
    fields_missing: List[str] = field(default_factory=list)
    popup_dismissed: bool = False
    patient_moved_to_roomed: bool = False

    @property
    def total_time_seconds(self) -> float:
        """Total processing time in seconds."""
        if self.end_time and self.start_time:
            return self.end_time - self.start_time
        return 0.0

    @property
    def extraction_completeness(self) -> float:
        """Percentage of fields successfully extracted."""
        total_fields = len(self.fields_extracted) + len(self.fields_missing)
        if total_fields == 0:
            return 0.0
        return (len(self.fields_extracted) / total_fields) * 100


@dataclass
class SessionStatistics:
    """Statistics for a single monitoring session."""
    session_id: str
    start_time: datetime
    end_time: Optional[datetime] = None
    total_patients_processed: int = 0
    successful_extractions: int = 0
    partial_extractions: int = 0
    failed_extractions: int = 0
    skipped_patients: int = 0
    errors: int = 0
    total_processing_time_seconds: float = 0.0
    waiting_room_scans: int = 0
    roomed_patients_tracked: int = 0
    discharged_patients: int = 0
    popups_dismissed: int = 0
    patients_moved_during_processing: int = 0

    @property
    def duration_minutes(self) -> float:
        """Session duration in minutes."""
        end = self.end_time or datetime.now()
        return (end - self.start_time).total_seconds() / 60

    @property
    def avg_time_per_patient_seconds(self) -> float:
        """Average processing time per patient."""
        if self.total_patients_processed == 0:
            return 0.0
        return self.total_processing_time_seconds / self.total_patients_processed

    @property
    def success_rate(self) -> float:
        """Percentage of successful extractions."""
        total = self.successful_extractions + self.partial_extractions + self.failed_extractions
        if total == 0:
            return 0.0
        return (self.successful_extractions / total) * 100

    @property
    def patients_per_hour(self) -> float:
        """Patients processed per hour."""
        hours = self.duration_minutes / 60
        if hours == 0:
            return 0.0
        return self.total_patients_processed / hours


@dataclass
class DailyStatistics:
    """Aggregated statistics for a day."""
    date: str
    sessions: int = 0
    total_patients: int = 0
    successful_extractions: int = 0
    partial_extractions: int = 0
    failed_extractions: int = 0
    total_processing_time_seconds: float = 0.0
    total_monitoring_time_minutes: float = 0.0
    avg_time_per_patient_seconds: float = 0.0
    peak_patients_per_hour: float = 0.0
    errors: int = 0

    @property
    def success_rate(self) -> float:
        """Daily success rate."""
        total = self.successful_extractions + self.partial_extractions + self.failed_extractions
        if total == 0:
            return 0.0
        return (self.successful_extractions / total) * 100


class AnalyticsEngine:
    """
    Core analytics engine for MHT Agentic.

    Tracks all patient processing operations, maintains session and daily statistics,
    and persists data for historical analysis.
    """

    REQUIRED_FIELDS = [
        'first_name', 'last_name', 'dob', 'mrn', 'cell_phone',
        'email', 'gender', 'insurance', 'race', 'ethnicity'
    ]

    def __init__(self, data_dir: Optional[Path] = None):
        """Initialize analytics engine."""
        self.data_dir = data_dir or Path("output/analytics")
        self.data_dir.mkdir(parents=True, exist_ok=True)

        # Current session tracking
        self.current_session: Optional[SessionStatistics] = None
        self.current_patient_event: Optional[PatientProcessingEvent] = None
        self._stage_start_time: float = 0.0

        # Event history for current session
        self.patient_events: List[PatientProcessingEvent] = []

        # Timing data for analysis
        self.stage_timings: Dict[str, List[float]] = {stage.value: [] for stage in ProcessingStage}

        # Callbacks for real-time updates
        self._on_stats_update: Optional[Callable[[Dict], None]] = None

        # Thread safety
        self._lock = threading.Lock()

        # Load existing daily stats
        self._daily_stats: Dict[str, DailyStatistics] = {}
        self._load_daily_stats()

        # Log buffer for deep logging
        self._log_buffer: List[Dict] = []
        self._max_log_buffer = 10000

    def set_stats_callback(self, callback: Callable[[Dict], None]):
        """Set callback for real-time stats updates."""
        self._on_stats_update = callback

    def start_session(self) -> str:
        """Start a new monitoring session."""
        with self._lock:
            session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.current_session = SessionStatistics(
                session_id=session_id,
                start_time=datetime.now()
            )
            self.patient_events = []
            self.stage_timings = {stage.value: [] for stage in ProcessingStage}

            self._log_event("SESSION_START", {
                "session_id": session_id,
                "start_time": self.current_session.start_time.isoformat()
            })

            return session_id

    def end_session(self):
        """End current monitoring session and save statistics."""
        with self._lock:
            if not self.current_session:
                return

            self.current_session.end_time = datetime.now()

            # Update daily stats
            self._update_daily_stats()

            # Save session data
            self._save_session_data()

            self._log_event("SESSION_END", {
                "session_id": self.current_session.session_id,
                "duration_minutes": self.current_session.duration_minutes,
                "total_patients": self.current_session.total_patients_processed,
                "success_rate": self.current_session.success_rate
            })

            # Notify callback
            if self._on_stats_update:
                self._on_stats_update(self.get_current_stats())

    def start_patient_processing(self, patient_name: str, mrn: str = "") -> PatientProcessingEvent:
        """Start tracking a patient processing operation."""
        with self._lock:
            self.current_patient_event = PatientProcessingEvent(
                patient_name=patient_name,
                patient_mrn=mrn,
                start_time=time.time()
            )

            self._log_event("PATIENT_PROCESSING_START", {
                "patient_name": patient_name,
                "mrn": mrn,
                "timestamp": datetime.now().isoformat()
            })

            return self.current_patient_event

    def start_stage(self, stage: ProcessingStage):
        """Start timing a processing stage."""
        self._stage_start_time = time.time()
        self._log_event("STAGE_START", {
            "stage": stage.value,
            "patient": self.current_patient_event.patient_name if self.current_patient_event else "unknown"
        })

    def end_stage(self, stage: ProcessingStage, success: bool = True, details: str = ""):
        """End timing a processing stage."""
        if self._stage_start_time > 0:
            elapsed = time.time() - self._stage_start_time

            with self._lock:
                # Record timing
                self.stage_timings[stage.value].append(elapsed)

                if self.current_patient_event:
                    self.current_patient_event.stage_timings[stage.value] = elapsed

                self._log_event("STAGE_END", {
                    "stage": stage.value,
                    "elapsed_seconds": round(elapsed, 3),
                    "success": success,
                    "details": details,
                    "patient": self.current_patient_event.patient_name if self.current_patient_event else "unknown"
                })

            self._stage_start_time = 0.0

    def record_field_extraction(self, field_name: str, success: bool, value: str = ""):
        """Record extraction result for a specific field."""
        with self._lock:
            if self.current_patient_event:
                if success:
                    self.current_patient_event.fields_extracted.append(field_name)
                else:
                    self.current_patient_event.fields_missing.append(field_name)

            self._log_event("FIELD_EXTRACTION", {
                "field": field_name,
                "success": success,
                "has_value": bool(value),
                "patient": self.current_patient_event.patient_name if self.current_patient_event else "unknown"
            })

    def record_popup_dismissed(self, popup_type: str = ""):
        """Record that a popup was dismissed."""
        with self._lock:
            if self.current_patient_event:
                self.current_patient_event.popup_dismissed = True
            if self.current_session:
                self.current_session.popups_dismissed += 1

            self._log_event("POPUP_DISMISSED", {
                "popup_type": popup_type,
                "patient": self.current_patient_event.patient_name if self.current_patient_event else "unknown"
            })

    def record_patient_moved(self):
        """Record that patient moved from Waiting Room to Roomed during processing."""
        with self._lock:
            if self.current_patient_event:
                self.current_patient_event.patient_moved_to_roomed = True
            if self.current_session:
                self.current_session.patients_moved_during_processing += 1

            self._log_event("PATIENT_MOVED", {
                "patient": self.current_patient_event.patient_name if self.current_patient_event else "unknown"
            })

    def end_patient_processing(self, result: ProcessingResult, error_message: str = ""):
        """Complete tracking for current patient."""
        should_notify = False
        with self._lock:
            if not self.current_patient_event:
                return

            self.current_patient_event.end_time = time.time()
            self.current_patient_event.result = result
            self.current_patient_event.error_message = error_message

            # Determine fields missing
            extracted_set = set(self.current_patient_event.fields_extracted)
            for field in self.REQUIRED_FIELDS:
                if field not in extracted_set and field not in self.current_patient_event.fields_missing:
                    self.current_patient_event.fields_missing.append(field)

            # Update session stats
            if self.current_session:
                self.current_session.total_patients_processed += 1
                self.current_session.total_processing_time_seconds += self.current_patient_event.total_time_seconds

                if result == ProcessingResult.SUCCESS:
                    self.current_session.successful_extractions += 1
                elif result == ProcessingResult.PARTIAL:
                    self.current_session.partial_extractions += 1
                elif result == ProcessingResult.FAILED:
                    self.current_session.failed_extractions += 1
                elif result == ProcessingResult.SKIPPED:
                    self.current_session.skipped_patients += 1
                elif result == ProcessingResult.ERROR:
                    self.current_session.errors += 1

            # Store event
            self.patient_events.append(self.current_patient_event)

            self._log_event("PATIENT_PROCESSING_END", {
                "patient_name": self.current_patient_event.patient_name,
                "result": result.value,
                "total_time_seconds": round(self.current_patient_event.total_time_seconds, 3),
                "fields_extracted": len(self.current_patient_event.fields_extracted),
                "fields_missing": len(self.current_patient_event.fields_missing),
                "extraction_completeness": round(self.current_patient_event.extraction_completeness, 1),
                "error_message": error_message
            })

            self.current_patient_event = None
            should_notify = self._on_stats_update is not None

        # Notify callback OUTSIDE the lock to prevent deadlock
        if should_notify:
            self._on_stats_update(self.get_current_stats())

    def record_waiting_room_scan(self, patients_found: int, qualified_count: int):
        """Record a waiting room scan event."""
        with self._lock:
            if self.current_session:
                self.current_session.waiting_room_scans += 1

            self._log_event("WAITING_ROOM_SCAN", {
                "patients_found": patients_found,
                "qualified_count": qualified_count,
                "scan_number": self.current_session.waiting_room_scans if self.current_session else 0
            })

    def record_roomed_patient(self, patient_name: str, room: str = ""):
        """Record a patient being moved to roomed."""
        with self._lock:
            if self.current_session:
                self.current_session.roomed_patients_tracked += 1

            self._log_event("PATIENT_ROOMED", {
                "patient_name": patient_name,
                "room": room
            })

    def record_discharged_patient(self, patient_name: str):
        """Record a patient being discharged."""
        with self._lock:
            if self.current_session:
                self.current_session.discharged_patients += 1

            self._log_event("PATIENT_DISCHARGED", {
                "patient_name": patient_name
            })

    def record_error(self, error_type: str, error_message: str, context: str = ""):
        """Record an error event."""
        with self._lock:
            self._log_event("ERROR", {
                "error_type": error_type,
                "error_message": error_message,
                "context": context,
                "patient": self.current_patient_event.patient_name if self.current_patient_event else "none"
            })

    def _log_event(self, event_type: str, data: Dict):
        """Add event to log buffer."""
        event = {
            "timestamp": datetime.now().isoformat(),
            "event_type": event_type,
            **data
        }

        self._log_buffer.append(event)

        # Trim buffer if too large
        if len(self._log_buffer) > self._max_log_buffer:
            self._log_buffer = self._log_buffer[-self._max_log_buffer // 2:]

        # Also log to standard logger
        logger.debug(f"{event_type}: {data}")

    def get_current_stats(self) -> Dict[str, Any]:
        """Get current session statistics as dictionary."""
        with self._lock:
            if not self.current_session:
                return {}

            # Calculate stage averages
            stage_averages = {}
            for stage, timings in self.stage_timings.items():
                if timings:
                    stage_averages[stage] = {
                        "avg": round(statistics.mean(timings), 3),
                        "min": round(min(timings), 3),
                        "max": round(max(timings), 3),
                        "count": len(timings)
                    }

            return {
                "session_id": self.current_session.session_id,
                "duration_minutes": round(self.current_session.duration_minutes, 1),
                "total_patients_processed": self.current_session.total_patients_processed,
                "successful_extractions": self.current_session.successful_extractions,
                "partial_extractions": self.current_session.partial_extractions,
                "failed_extractions": self.current_session.failed_extractions,
                "skipped_patients": self.current_session.skipped_patients,
                "errors": self.current_session.errors,
                "success_rate": round(self.current_session.success_rate, 1),
                "avg_time_per_patient_seconds": round(self.current_session.avg_time_per_patient_seconds, 1),
                "patients_per_hour": round(self.current_session.patients_per_hour, 1),
                "waiting_room_scans": self.current_session.waiting_room_scans,
                "roomed_patients_tracked": self.current_session.roomed_patients_tracked,
                "discharged_patients": self.current_session.discharged_patients,
                "popups_dismissed": self.current_session.popups_dismissed,
                "patients_moved_during_processing": self.current_session.patients_moved_during_processing,
                "stage_averages": stage_averages
            }

    def get_daily_stats(self, date_str: Optional[str] = None) -> Dict[str, Any]:
        """Get daily statistics."""
        if date_str is None:
            date_str = date.today().isoformat()

        with self._lock:
            if date_str in self._daily_stats:
                stats = self._daily_stats[date_str]
                return asdict(stats)
            return {}

    def get_weekly_summary(self) -> Dict[str, Any]:
        """Get summary for the past 7 days."""
        summary = {
            "days": [],
            "total_patients": 0,
            "total_successful": 0,
            "avg_success_rate": 0.0,
            "avg_time_per_patient": 0.0,
            "total_sessions": 0
        }

        with self._lock:
            today = date.today()
            success_rates = []
            times_per_patient = []

            for i in range(7):
                day = today - timedelta(days=i)
                date_str = day.isoformat()

                if date_str in self._daily_stats:
                    stats = self._daily_stats[date_str]
                    day_summary = {
                        "date": date_str,
                        "patients": stats.total_patients,
                        "successful": stats.successful_extractions,
                        "success_rate": stats.success_rate,
                        "sessions": stats.sessions
                    }
                    summary["days"].append(day_summary)
                    summary["total_patients"] += stats.total_patients
                    summary["total_successful"] += stats.successful_extractions
                    summary["total_sessions"] += stats.sessions

                    if stats.success_rate > 0:
                        success_rates.append(stats.success_rate)
                    if stats.avg_time_per_patient_seconds > 0:
                        times_per_patient.append(stats.avg_time_per_patient_seconds)

            if success_rates:
                summary["avg_success_rate"] = round(statistics.mean(success_rates), 1)
            if times_per_patient:
                summary["avg_time_per_patient"] = round(statistics.mean(times_per_patient), 1)

        return summary

    def get_value_metrics(self) -> Dict[str, Any]:
        """
        Calculate metrics to evaluate if the software is worth pursuing.

        Returns metrics like:
        - Time saved per patient (estimated manual time vs automated)
        - Projected daily/monthly time savings
        - Success/reliability metrics
        - ROI indicators
        """
        ESTIMATED_MANUAL_TIME_SECONDS = 180  # 3 minutes to manually enter patient data

        with self._lock:
            weekly = self.get_weekly_summary()
            current = self.get_current_stats() if self.current_session else {}

            total_patients = weekly.get("total_patients", 0)
            avg_automated_time = current.get("avg_time_per_patient_seconds", 15)

            # Calculate time savings
            manual_time_total = total_patients * ESTIMATED_MANUAL_TIME_SECONDS
            automated_time_total = total_patients * avg_automated_time
            time_saved_seconds = manual_time_total - automated_time_total
            time_saved_hours = time_saved_seconds / 3600

            # Project to monthly
            avg_daily_patients = total_patients / 7 if total_patients > 0 else 0
            projected_monthly_patients = avg_daily_patients * 22  # Working days
            projected_monthly_savings_hours = (projected_monthly_patients * (ESTIMATED_MANUAL_TIME_SECONDS - avg_automated_time)) / 3600

            return {
                "time_analysis": {
                    "estimated_manual_time_per_patient_seconds": ESTIMATED_MANUAL_TIME_SECONDS,
                    "actual_automated_time_per_patient_seconds": round(avg_automated_time, 1),
                    "time_saved_per_patient_seconds": round(ESTIMATED_MANUAL_TIME_SECONDS - avg_automated_time, 1),
                    "efficiency_multiplier": round(ESTIMATED_MANUAL_TIME_SECONDS / avg_automated_time, 1) if avg_automated_time > 0 else 0
                },
                "weekly_metrics": {
                    "patients_processed": total_patients,
                    "time_saved_hours": round(time_saved_hours, 1),
                    "success_rate": weekly.get("avg_success_rate", 0),
                    "sessions_run": weekly.get("total_sessions", 0)
                },
                "monthly_projection": {
                    "projected_patients": round(projected_monthly_patients),
                    "projected_time_saved_hours": round(projected_monthly_savings_hours, 1),
                    "equivalent_fte_days_saved": round(projected_monthly_savings_hours / 8, 1)  # 8-hour workday
                },
                "reliability": {
                    "success_rate": weekly.get("avg_success_rate", 0),
                    "uptime_indicator": "stable" if weekly.get("avg_success_rate", 0) > 85 else "needs_improvement"
                },
                "recommendation": self._generate_recommendation(weekly, avg_automated_time)
            }

    def _generate_recommendation(self, weekly: Dict, avg_time: float) -> str:
        """Generate a recommendation based on metrics."""
        success_rate = weekly.get("avg_success_rate", 0)
        total_patients = weekly.get("total_patients", 0)

        if total_patients < 10:
            return "INSUFFICIENT_DATA: Need more usage data for accurate assessment"

        if success_rate >= 90 and avg_time < 30:
            return "STRONG_VALUE: High success rate and fast processing - worth pursuing"
        elif success_rate >= 75 and avg_time < 45:
            return "GOOD_VALUE: Solid performance - continue development"
        elif success_rate >= 60:
            return "MODERATE_VALUE: Acceptable but needs improvement in reliability"
        else:
            return "NEEDS_WORK: Success rate below threshold - focus on bug fixes"

    def _update_daily_stats(self):
        """Update daily statistics from current session."""
        if not self.current_session:
            return

        date_str = self.current_session.start_time.date().isoformat()

        if date_str not in self._daily_stats:
            self._daily_stats[date_str] = DailyStatistics(date=date_str)

        daily = self._daily_stats[date_str]
        daily.sessions += 1
        daily.total_patients += self.current_session.total_patients_processed
        daily.successful_extractions += self.current_session.successful_extractions
        daily.partial_extractions += self.current_session.partial_extractions
        daily.failed_extractions += self.current_session.failed_extractions
        daily.total_processing_time_seconds += self.current_session.total_processing_time_seconds
        daily.total_monitoring_time_minutes += self.current_session.duration_minutes
        daily.errors += self.current_session.errors

        # Recalculate averages
        if daily.total_patients > 0:
            daily.avg_time_per_patient_seconds = daily.total_processing_time_seconds / daily.total_patients

        # Update peak rate
        if self.current_session.patients_per_hour > daily.peak_patients_per_hour:
            daily.peak_patients_per_hour = self.current_session.patients_per_hour

        self._save_daily_stats()

    def _save_session_data(self):
        """Save session data to file."""
        if not self.current_session:
            return

        session_file = self.data_dir / f"session_{self.current_session.session_id}.json"

        data = {
            "session": asdict(self.current_session),
            "patient_events": [asdict(e) for e in self.patient_events],
            "stage_timings": {k: v for k, v in self.stage_timings.items() if v},
            "log_buffer": self._log_buffer[-1000:]  # Last 1000 events
        }

        # Convert enums and datetime
        data["session"]["start_time"] = self.current_session.start_time.isoformat()
        if self.current_session.end_time:
            data["session"]["end_time"] = self.current_session.end_time.isoformat()

        for event in data["patient_events"]:
            event["result"] = event["result"].value if isinstance(event["result"], ProcessingResult) else event["result"]

        try:
            with open(session_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, default=str)
            logger.info(f"Session data saved: {session_file}")
        except Exception as e:
            logger.error(f"Failed to save session data: {e}")

    def _save_daily_stats(self):
        """Save daily statistics to file."""
        stats_file = self.data_dir / "daily_stats.json"

        data = {date_str: asdict(stats) for date_str, stats in self._daily_stats.items()}

        try:
            with open(stats_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            logger.error(f"Failed to save daily stats: {e}")

    def _load_daily_stats(self):
        """Load daily statistics from file."""
        stats_file = self.data_dir / "daily_stats.json"

        if not stats_file.exists():
            return

        try:
            with open(stats_file, 'r', encoding='utf-8') as f:
                data = json.load(f)

            for date_str, stats_dict in data.items():
                self._daily_stats[date_str] = DailyStatistics(**stats_dict)

            logger.info(f"Loaded daily stats for {len(self._daily_stats)} days")
        except Exception as e:
            logger.error(f"Failed to load daily stats: {e}")

    def export_logs(self, filepath: Optional[Path] = None) -> Path:
        """Export all logs to a file."""
        if filepath is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filepath = self.data_dir / f"logs_export_{timestamp}.json"

        with self._lock:
            data = {
                "export_time": datetime.now().isoformat(),
                "session_stats": self.get_current_stats(),
                "daily_stats": {k: asdict(v) for k, v in self._daily_stats.items()},
                "value_metrics": self.get_value_metrics(),
                "log_events": self._log_buffer
            }

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, default=str)

        return filepath

    def get_log_summary(self) -> str:
        """Get a human-readable summary of recent logs."""
        with self._lock:
            lines = []
            lines.append("=" * 60)
            lines.append("MHT AGENTIC ANALYTICS SUMMARY")
            lines.append("=" * 60)

            if self.current_session:
                stats = self.get_current_stats()
                lines.append(f"\nCurrent Session: {stats.get('session_id', 'N/A')}")
                lines.append(f"Duration: {stats.get('duration_minutes', 0):.1f} minutes")
                lines.append(f"Patients Processed: {stats.get('total_patients_processed', 0)}")
                lines.append(f"Successful: {stats.get('successful_extractions', 0)}")
                lines.append(f"Success Rate: {stats.get('success_rate', 0):.1f}%")
                lines.append(f"Avg Time/Patient: {stats.get('avg_time_per_patient_seconds', 0):.1f}s")
                lines.append(f"Patients/Hour: {stats.get('patients_per_hour', 0):.1f}")

            # Value metrics
            value = self.get_value_metrics()
            if value:
                lines.append("\n" + "-" * 40)
                lines.append("VALUE ASSESSMENT")
                lines.append("-" * 40)
                time_analysis = value.get("time_analysis", {})
                lines.append(f"Efficiency Multiplier: {time_analysis.get('efficiency_multiplier', 0)}x faster")
                lines.append(f"Time Saved/Patient: {time_analysis.get('time_saved_per_patient_seconds', 0):.0f}s")

                weekly = value.get("weekly_metrics", {})
                lines.append(f"Weekly Time Saved: {weekly.get('time_saved_hours', 0):.1f} hours")

                monthly = value.get("monthly_projection", {})
                lines.append(f"Monthly Projection: {monthly.get('projected_time_saved_hours', 0):.1f} hours saved")

                lines.append(f"\nRecommendation: {value.get('recommendation', 'N/A')}")

            lines.append("\n" + "=" * 60)

            return "\n".join(lines)


# Global analytics instance
_analytics: Optional[AnalyticsEngine] = None


def get_analytics() -> AnalyticsEngine:
    """Get or create the global analytics engine."""
    global _analytics
    if _analytics is None:
        _analytics = AnalyticsEngine()
    return _analytics


def reset_analytics():
    """Reset the global analytics engine."""
    global _analytics
    if _analytics:
        _analytics.end_session()
    _analytics = None
