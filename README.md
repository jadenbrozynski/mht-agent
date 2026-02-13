# MHTAgentic

Desktop automation for Experity EMR that extracts patient demographics and triggers MHT SmarTest behavioral health assessments (PHQ-9, GAD-7).

## What It Does

- Monitors the Experity Tracking Board for qualified patients (chart created)
- Extracts demographics from patient charts (name, DOB, MRN, phone, email, insurance, race, ethnicity, language)
- Generates MHT API JSON payloads for assessment triggers
- Tracks patient lifecycle: waiting room -> roomed -> discharged
- Processes completed assessment results back into Experity (outbound flow)
- Floating overlay UI showing extraction status and session analytics

## Setup

```
pip install -r requirements.txt
```

Copy `config/.env.example` to `config/.env` and fill in your Experity credentials:

```
EXPERITY_USERNAME=your_username
EXPERITY_PASSWORD=your_password
```

## Usage

**Full launch with login flow:**
```
pythonw launcher.pyw       # silent (desktop shortcut)
python launcher.pyw        # with console output
```

**Resume monitoring** (Experity already open on Tracking Board):
```
python start_monitoring.py
```

**Outbound only** (process completed assessments back into Experity):
```
python start_with_outbound.py
```

**Demo mode** (minimal overlays, simulated MHT responses):
```
python start_demo.py
```

## Project Structure

```
launcher.pyw                 Main entry point (login + monitoring)
start_monitoring.py          Skip login, start from Tracking Board
start_with_outbound.py       Outbound assessment entry only
start_demo.py                Demo mode with simulated responses

config/
  .env.example               Credential template
  login_macro.json           Recorded login macro

mhtagentic/
  db/
    database.py              SQLite event tracking
    mht_simulator.py         Mock MHT API responses (testing)
    result_processor.py      Outbound result processing
  desktop/
    automation.py            Window management, screenshots
    control_overlay.py       Floating UI panels
    analytics.py             Session metrics and daily stats
    macro_recorder.py        Login macro recording
  outbound/
    outbound_worker.py       Assessment entry automation

output/                      Runtime data (gitignored)
  mht_api/                   Generated JSON payloads
  analytics/                 Session and daily stats
  mht_data.db                Event database
```

## Architecture

Two data flows:

**Inbound** -- Extracts patient data from Experity, converts to MHT API format, writes JSON + database events.

**Outbound** -- When MHT returns assessment results, the outbound worker enters scores back into the patient chart via Procedures/Supplies.

Events tracked in SQLite with status codes (0=initial, 10=converted, 40=sent, 100=complete). Negative values indicate failure at that stage. Max 4 retries before marking failed.

## Requirements

- Windows 10/11
- Python 3.9+
- Experity EMR desktop application
