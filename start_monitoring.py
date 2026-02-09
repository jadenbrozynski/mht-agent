"""
MHT Agentic - Resume Monitoring

Starts monitoring directly from the Tracking Board without the login flow.
Use this when Experity is already open and logged in.

Includes:
1. Inbound monitoring (patient extraction from Tracking Board)
2. Outbound worker (processes completed MHT assessments back to Experity)

Usage:
    python start_monitoring.py

Prerequisites:
    - Experity EMR must be open
    - Must be on the Tracking Board view
    - Waiting Room tab must be visible
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

from mhtagentic.desktop.control_overlay import ControlOverlay, reset_control_overlay
from mhtagentic.outbound.outbound_worker import OutboundWorker
from mhtagentic.db.mht_simulator import MHTResponseSimulator

# Database path
DB_PATH = SCRIPT_DIR / "output" / "mht_data.db"

# Reset and create control overlay
reset_control_overlay()
control = ControlOverlay(on_kill=lambda: None)
control.start()
time.sleep(0.3)

# Start MHT Response Simulator (creates outbound events ~30 seconds after inbound)
simulator = MHTResponseSimulator(
    str(DB_PATH),
    response_delay_seconds=15  # 15 seconds delay for testing (30 in production)
)
simulator.start()
logger.info("MHT Simulator started - will create outbound events 15 seconds after inbound")

# Create outbound worker (but DON'T auto-start polling - will be called during refresh wait)
outbound = OutboundWorker(
    str(DB_PATH),
    poll_interval=5.0,
    overlay=control
)

def on_outbound_processed(result):
    if result['success']:
        logger.info(f"✓ Outbound processed: {result['patient']} (event {result['event_id']})")
    else:
        logger.error(f"✗ Outbound failed: {result['patient']} (event {result['event_id']})")

outbound.set_callback(on_outbound_processed)
# Don't auto-start - outbound will be processed during refresh wait in monitoring loop
logger.info("Outbound worker ready - will process during refresh wait periods")

# Import launcher module to get _start_monitoring
import importlib.util
spec = importlib.util.spec_from_file_location('launcher', SCRIPT_DIR / 'launcher.pyw')
launcher = importlib.util.module_from_spec(spec)
launcher.SCRIPT_DIR = SCRIPT_DIR
spec.loader.exec_module(launcher)

print('Starting monitoring (inbound + outbound)...')
print('Click X on control panel to stop')
print()

# Run the full monitoring function from launcher.pyw (pass outbound worker for refresh-time processing)
launcher._start_monitoring(control, outbound_worker=outbound)

# Stop workers when monitoring stops
simulator.stop()
print('Monitoring stopped')
