"""
MHT Agentic - RDP Monitor Dashboard

Web dashboard for monitoring bot activity inside RDP sessions.
Shows live screenshots, session status, and analytics.

Usage:
    python start_dashboard.py

Then open http://localhost:5555 in your browser.

Prerequisites:
    pip install flask
"""
import sys
import os
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
logger = logging.getLogger("mht_dashboard")

PORT = 5555

if __name__ == "__main__":
    from dashboard.server import create_app

    app = create_app(project_root=SCRIPT_DIR, port=PORT)

    logger.info(f"Starting MHT Dashboard on http://localhost:{PORT}")
    print(f"\n  MHT Agentic - RDP Monitor Dashboard")
    print(f"  Open http://localhost:{PORT} in your browser\n")

    app.run(host="0.0.0.0", port=PORT, debug=False, threaded=True)
