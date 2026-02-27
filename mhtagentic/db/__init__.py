"""MHT Database package for event tracking."""

from .database import (
    MHTDatabase,
    EventStatus,
    claim_slot,
    release_slot,
    heartbeat_slot,
    get_slots,
    cleanup_stale_slots,
    log_bot_error,
    get_bot_errors,
    clear_bot_errors,
)
from .mht_simulator import MHTResponseSimulator
from .result_processor import MHTResultProcessor

__all__ = [
    "MHTDatabase",
    "EventStatus",
    "MHTResponseSimulator",
    "MHTResultProcessor",
    "claim_slot",
    "release_slot",
    "heartbeat_slot",
    "get_slots",
    "cleanup_stale_slots",
    "log_bot_error",
    "get_bot_errors",
    "clear_bot_errors",
]
