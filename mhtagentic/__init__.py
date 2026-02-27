"""MHTAgentic - Automated Experity patient data extraction agent."""

import os
from pathlib import Path

__version__ = "0.1.0"

# Shared output directory — C:\ProgramData\MHTAgentic is writable by all
# authenticated users on Windows, so any RDP user account can write here.
OUTPUT_DIR = Path(os.environ.get("MHT_OUTPUT_DIR",
                                 os.path.join(os.environ.get("PROGRAMDATA", r"C:\ProgramData"),
                                              "MHTAgentic")))
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
