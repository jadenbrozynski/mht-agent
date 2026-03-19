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
    # OTP queue
    otp_queue_request,
    otp_queue_try_grant,
    otp_queue_complete,
    otp_queue_clear,
    # Typing lock
    typing_lock_acquire,
    typing_lock_release,
    typing_lock_clear,
    typing_lock_context,
    # Clinic locations
    assign_location,
    release_location,
    get_active_locations,
    get_all_locations,
    add_location,
    remove_location,
    release_all_locations,
    toggle_location_active,
    sync_bot_slots,
    # Bot config
    get_config,
    set_config,
)
from .mht_simulator import MHTResponseSimulator
from .result_processor import MHTResultProcessor
from .integration_client import IntegrationClient

__all__ = [
    "MHTDatabase",
    "EventStatus",
    "MHTResponseSimulator",
    "MHTResultProcessor",
    "IntegrationClient",
    "claim_slot",
    "release_slot",
    "heartbeat_slot",
    "get_slots",
    "cleanup_stale_slots",
    "log_bot_error",
    "get_bot_errors",
    "clear_bot_errors",
    # OTP queue
    "otp_queue_request",
    "otp_queue_try_grant",
    "otp_queue_complete",
    "otp_queue_clear",
    # Typing lock
    "typing_lock_acquire",
    "typing_lock_release",
    "typing_lock_clear",
    "typing_lock_context",
    # Clinic locations
    "assign_location",
    "release_location",
    "get_active_locations",
    "get_all_locations",
    "add_location",
    "remove_location",
    "release_all_locations",
    "toggle_location_active",
    "sync_bot_slots",
    # Bot config
    "get_config",
    "set_config",
]
