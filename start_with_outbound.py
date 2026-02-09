"""
MHT Agentic - Full Monitoring with Outbound Processing

Starts both:
1. Inbound monitoring (patient extraction from Tracking Board)
2. Outbound worker (processes completed MHT assessments back to Experity)

Usage:
    python start_with_outbound.py

Prerequisites:
    - Experity EMR must be open
    - Must be on the Tracking Board view
"""
import sys
import os
import time
import logging
from pathlib import Path

# Setup paths
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(SCRIPT_DIR))
os.chdir(str(SCRIPT_DIR))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(name)s] %(levelname)s: %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger("mht_agentic")

# Database path
DB_PATH = SCRIPT_DIR / "output" / "mht_data.db"

from mhtagentic.outbound.outbound_worker import OutboundWorker
from mhtagentic.desktop.control_overlay import ControlOverlay

def main():
    print("=" * 60)
    print("MHT AGENTIC - FULL MONITORING WITH OUTBOUND")
    print("=" * 60)
    print(f"Database: {DB_PATH}")
    print()
    print("This will:")
    print("1. Monitor for completed outbound events (status=100)")
    print("2. Automatically process them into Experity")
    print()
    print("Press Ctrl+C to stop")
    print("=" * 60)
    print()

    # Create and start control overlay for UI feedback
    overlay = ControlOverlay()
    overlay.start()
    overlay.set_status("Outbound Monitor Ready", "#4ECDC4")
    overlay.set_step("Polling for completed assessments...")

    # Create outbound worker with overlay integration
    outbound = OutboundWorker(
        str(DB_PATH),
        poll_interval=5.0,  # Check every 5 seconds
        overlay=overlay
    )

    # Callback when outbound event is processed
    def on_outbound_processed(result):
        if result['success']:
            logger.info(f"✓ Processed: {result['patient']} (event {result['event_id']})")
            overlay.set_status(f"✓ Patient Processed", "#5cb85c")
            overlay.set_step(f"{result['patient']} - Complete")
        else:
            logger.error(f"✗ Failed: {result['patient']} (event {result['event_id']})")
            overlay.set_status(f"✗ Processing Failed", "#d9534f")

    outbound.set_callback(on_outbound_processed)

    # Start outbound worker
    outbound.start()
    logger.info("Outbound worker started - polling every 5 seconds")

    try:
        while True:
            # Update overlay when idle (not processing)
            if not outbound.is_processing():
                overlay.set_status("Monitoring...", "#4ECDC4")
                overlay.set_step("Waiting for completed assessments (status=100)")
            time.sleep(2)
    except KeyboardInterrupt:
        print("\n\nShutting down...")
        outbound.stop()
        overlay.stop()
        print("Done.")


if __name__ == "__main__":
    main()
