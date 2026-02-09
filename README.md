# MHT Agentic

Desktop automation tool for extracting patient data from Experity EMR to trigger MHT SmarTest behavioral health assessments.

## Overview

MHT Agentic monitors the Experity Tracking Board, identifies qualified patients (age 12+), extracts their demographic data, and generates JSON files for the MHT API to trigger PHQ-9 and GAD-7 assessments.

## Requirements

- Python 3.9+
- Windows 10/11
- Experity EMR (desktop application)

## Installation

```bash
pip install -r requirements.txt
```

## Usage

### Full Application (with login)

```bash
# Silent mode (no console) - use desktop shortcut
pythonw launcher.pyw

# With console output
python launcher.pyw
```

### Resume Monitoring (skip login)

Use when Experity is already open and on the Tracking Board:

```bash
python start_monitoring.py
```

## Configuration

### Login Macro

The application uses a recorded login macro stored in `config/login_macro.json`. To record a new macro, use the Record mode in the application.

### Output

Extracted patient data is saved as JSON files in `output/mht_api/` with the format:

```
mht_api_{LASTNAME}_{FIRSTNAME}_{TIMESTAMP}.json
```

## Project Structure

```
MHTAgentic/
├── launcher.pyw           # Main application entry point
├── start_monitoring.py    # Resume monitoring script
├── requirements.txt       # Python dependencies
├── config/
│   └── login_macro.json   # Recorded login actions
└── mhtagentic/
    └── desktop/
        ├── automation.py      # Desktop automation utilities
        ├── control_overlay.py # UI overlay components
        └── macro_recorder.py  # Login macro recording
```

## Features

- Monitors Waiting Room for qualified patients
- Color-coded visual overlays showing patient status
- Automatic popup dialog dismissal
- Patient data extraction from Demographics tab
- Roomed/discharged patient tracking
- JSON generation for MHT API integration

## License

Proprietary - Mental Health Technologies Inc.
