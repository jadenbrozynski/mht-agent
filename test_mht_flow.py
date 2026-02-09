"""
MHT Full Flow Test

Tests the complete flow:
1. Start MHT Response Simulator (creates outbound events after delay)
2. Start Result Processor (processes outbound events, updates to status=100)
3. Run monitoring to scrape patients (creates inbound events)

This simulates the full production workflow locally.
"""

import sys
import os
import time
import logging
from pathlib import Path
from datetime import datetime

# Setup paths
SCRIPT_DIR = Path(__file__).parent.resolve()
sys.path.insert(0, str(SCRIPT_DIR))
os.chdir(str(SCRIPT_DIR))

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(name)s] %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger("mht_flow_test")

from mhtagentic.db.mht_simulator import MHTResponseSimulator
from mhtagentic.db.result_processor import MHTResultProcessor

DB_PATH = SCRIPT_DIR / "output" / "mht_data.db"


def print_status(simulator: MHTResponseSimulator, processor: MHTResultProcessor):
    """Print current status."""
    stats = processor.get_stats()
    print("\n" + "=" * 60)
    print("MHT FLOW STATUS")
    print("=" * 60)
    print(f"Simulator Running: {simulator.is_running()}")
    print(f"Processor Running: {processor.is_running()}")
    print()
    print("INBOUND (I) - Automation -> MHT:")
    print(f"  Total:    {stats['inbound_total']}")
    print(f"  Complete: {stats['inbound_complete']}")
    print()
    print("OUTBOUND (O) - MHT -> Automation:")
    print(f"  Pending:  {stats['outbound_pending']} (status=10)")
    print(f"  Complete: {stats['outbound_complete']} (status=100)")
    print(f"  Failed:   {stats['outbound_failed']} (status<0)")
    print("=" * 60 + "\n")


def main():
    print("\n" + "=" * 60)
    print("MHT FULL FLOW TEST")
    print("=" * 60)
    print(f"Database: {DB_PATH}")
    print()
    print("This test will:")
    print("1. Start MHT Response Simulator (30s delay)")
    print("2. Start Result Processor (10s poll interval)")
    print("3. Monitor for inbound events and generate mock responses")
    print()
    print("Press Ctrl+C to stop")
    print("=" * 60 + "\n")

    # Initialize components
    simulator = MHTResponseSimulator(
        str(DB_PATH),
        response_delay_seconds=30  # Wait 30s before generating response
    )

    processor = MHTResultProcessor(
        str(DB_PATH),
        poll_interval_seconds=10  # Check every 10s for new results
    )

    # Callback when results are processed
    def on_result_processed(result):
        logger.info(f"RESULT PROCESSED: {result['patient_name']}")
        for a in result['assessments']:
            severity = a['severity']
            score = a['total_score']
            flagged = " [FLAGGED]" if a['flagged_abnormal'] else ""
            logger.info(f"  -> {a['assessment_name']}: Score {score} ({severity}){flagged}")

    processor.set_callback(on_result_processed)

    # Start both services
    simulator.start()
    processor.start()

    logger.info("Both services started. Waiting for events...")
    logger.info("Run start_monitoring.py in another terminal to create inbound events")

    try:
        # Print status periodically
        last_status_time = 0
        while True:
            current_time = time.time()

            # Print status every 30 seconds
            if current_time - last_status_time >= 30:
                print_status(simulator, processor)
                last_status_time = current_time

            time.sleep(1)

    except KeyboardInterrupt:
        print("\n\nShutting down...")
        simulator.stop()
        processor.stop()
        print_status(simulator, processor)
        print("Done.")


if __name__ == "__main__":
    main()
