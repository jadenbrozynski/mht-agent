"""
Flask web dashboard for monitoring MHT bot RDP sessions.

Provides:
- Live screenshots of mstsc windows (captured without focus change)
- Session status and bot health
- Analytics from SQLite DB and daily_stats.json
"""

import json
import logging
import threading
import time
from pathlib import Path
from typing import Dict

from flask import Flask, Response, jsonify, render_template, request

from mhtagentic.db import (
    cleanup_stale_slots,
    get_slots,
    release_slot,
    log_bot_error,
    get_bot_errors,
    clear_bot_errors,
)

from .screenshot_capture import capture_window_jpeg
from .session_monitor import (
    check_bot_health,
    clear_all_otp_signals,
    ensure_autostart_task,
    find_rdp_files,
    find_rdp_windows,
    _find_rdp_session_ids,
    get_recent_events,
    launch_bot_in_session,
    start_monitoring_in_sessions,
    start_rdp_session,
    stop_all_rdp_sessions,
    stop_rdp_session,
    wait_for_otp_complete,
    wait_for_rdp_session,
    _logoff_rdp_user,
    _parse_rdp_username,
)

logger = logging.getLogger("mht_dashboard.server")

# In-memory screenshot store: {session_id: jpeg_bytes}
_screenshots: Dict[str, bytes] = {}
_sessions: list = []
_lock = threading.Lock()

# In-memory start-all progress tracker
# Structure: { "running": bool, "bots": { "ExperityB": { "status": str, "attempt": int, "max_attempts": int, "error": str }, ... } }
_start_all_status: Dict = {"running": False, "bots": {}}

# Abort flag: set by stop-all to tell a running start-all thread to quit immediately
_start_all_abort: Dict = {"abort": False}


def create_app(project_root: Path, port: int = 5555) -> Flask:
    """
    Create and configure the Flask dashboard app.

    Args:
        project_root: Path to MHTAgentic directory
        port: Port to serve on

    Returns:
        Configured Flask app
    """
    app = Flask(
        __name__,
        template_folder=str(Path(__file__).parent / "templates"),
        static_folder=str(Path(__file__).parent / "static"),
    )

    from mhtagentic import OUTPUT_DIR
    from mhtagentic.db import MHTDatabase
    db_path = OUTPUT_DIR / "mht_data.db"
    analytics_dir = OUTPUT_DIR / "analytics"

    # Ensure DB tables + slot migration runs on dashboard startup
    MHTDatabase(str(db_path))

    # Directories to scan for .rdp files
    rdp_search_dirs = [
        project_root,
        project_root / "config",
        Path.home() / "Desktop",
    ]

    # --- Background session + screenshot thread ---
    # Only capture screenshots when DB state changes or on a slow fallback
    _last_db_state = {"updated_at": None, "count": 0}
    _last_capture_time = [0.0]
    CAPTURE_ON_CHANGE_INTERVAL = 3     # Check for DB changes every 3s
    FALLBACK_CAPTURE_INTERVAL = 30     # Capture anyway every 30s even if no change

    def _db_state_changed() -> bool:
        """Check if DB state has changed since last screenshot."""
        if not db_path.exists():
            return False
        try:
            import sqlite3 as _sql
            conn = _sql.connect(f"file:{db_path}?mode=ro", uri=True)
            conn.row_factory = _sql.Row
            cur = conn.cursor()
            cur.execute("SELECT MAX(updated_at) as u, COUNT(*) as c FROM common_event")
            row = cur.fetchone()
            conn.close()
            new_state = {"updated_at": row["u"], "count": row["c"]}
            if new_state != _last_db_state:
                _last_db_state.update(new_state)
                return True
        except Exception:
            pass
        return False

    def screenshot_loop():
        while True:
            try:
                windows = find_rdp_windows()

                # Decide whether to capture screenshots this tick
                now = time.time()
                db_changed = _db_state_changed()
                fallback_due = (now - _last_capture_time[0]) >= FALLBACK_CAPTURE_INTERVAL
                should_capture = db_changed or fallback_due

                with _lock:
                    _sessions.clear()
                    active_ids = set()

                    for win in windows:
                        session_id = str(win["hwnd"])
                        active_ids.add(session_id)

                        # Screenshots disabled to reduce CPU/memory load
                        # if should_capture:
                        #     jpeg = capture_window_jpeg(win["hwnd"], quality=70)
                        #     if jpeg:
                        #         _screenshots[session_id] = jpeg

                        # Detect agent role from window title
                        title_lower = win["title"].lower()
                        if "experityb" in title_lower:
                            agent_role = "inbound"
                        elif "experityd" in title_lower:
                            agent_role = "inbound"
                        elif "experityc" in title_lower:
                            agent_role = "outbound"
                        else:
                            agent_role = None

                        _sessions.append({
                            "id": session_id,
                            "title": win["title"],
                            "width": win["width"],
                            "height": win["height"],
                            "has_screenshot": session_id in _screenshots,
                            "agent_role": agent_role,
                        })

                    # Clean up screenshots for closed windows
                    stale = set(_screenshots.keys()) - active_ids
                    for sid in stale:
                        del _screenshots[sid]

                if should_capture and windows:
                    _last_capture_time[0] = now

            except Exception as e:
                logger.error(f"Screenshot loop error: {e}")

            time.sleep(CAPTURE_ON_CHANGE_INTERVAL)

    bg_thread = threading.Thread(target=screenshot_loop, daemon=True)
    bg_thread.start()
    logger.info("Background thread started (screenshots on DB change + 30s fallback)")

    # --- Routes ---

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/api/sessions")
    def api_sessions():
        """Active RDP sessions + bot health."""
        health = check_bot_health(db_path)

        with _lock:
            sessions = []
            for s in _sessions:
                sessions.append({**s, "bot_health": health})

        # If no RDP windows but DB is active, show a "headless" status
        if not sessions:
            sessions = [{
                "id": "none",
                "title": "No RDP sessions detected",
                "width": 0,
                "height": 0,
                "has_screenshot": False,
                "bot_health": health,
            }]

        return jsonify({"sessions": sessions, "bot_health": health})

    @app.route("/api/slots")
    def api_slots():
        """Return current bot slot status (read-only, no auto-cleanup)."""
        if db_path.exists():
            try:
                slots = get_slots(str(db_path))
            except Exception as e:
                logger.error(f"Slot query failed: {e}")
                slots = []
        else:
            slots = []
        return jsonify({"slots": slots})

    @app.route("/api/screenshots/<session_id>")
    def api_screenshot(session_id: str):
        """JPEG screenshot of a specific mstsc window."""
        with _lock:
            jpeg = _screenshots.get(session_id)

        if jpeg is None:
            return Response(status=404)

        return Response(jpeg, mimetype="image/jpeg")

    @app.route("/api/analytics/current")
    def api_analytics_current():
        """Today's stats from daily_stats.json + DB health."""
        today_str = time.strftime("%Y-%m-%d")
        stats = {}

        # Read daily_stats.json
        daily_stats_path = analytics_dir / "daily_stats.json"
        if daily_stats_path.exists():
            try:
                with open(daily_stats_path, "r") as f:
                    all_stats = json.load(f)
                stats = all_stats.get(today_str, {})
            except Exception as e:
                logger.error(f"Failed to read daily stats: {e}")

        # Merge live DB counts
        health = check_bot_health(db_path)

        return jsonify({
            "date": today_str,
            "daily_stats": stats,
            "live": {
                "events_today": health["events_today"],
                "inbound_today": health["inbound_today"],
                "outbound_today": health["outbound_today"],
                "errors_today": health["errors_today"],
                "bot_active": health["active"],
                "bot_status": health["status"],
                "last_event_at": health["last_event_at"],
            },
        })

    @app.route("/api/analytics/events")
    def api_analytics_events():
        """Recent 50 patient events from DB."""
        events = get_recent_events(db_path, limit=50)
        return jsonify({"events": events})

    # --- RDP Control Routes ---

    @app.route("/api/rdp/files")
    def api_rdp_files():
        """List available .rdp files."""
        files = find_rdp_files(rdp_search_dirs)
        return jsonify({"files": files})

    @app.route("/api/rdp/start", methods=["POST"])
    def api_rdp_start():
        """Disabled — use Start All (start_all_clean.py) instead."""
        return jsonify({"success": False, "error": "Use Start All instead"}), 400

    @app.route("/api/rdp/start-all", methods=["POST"])
    def api_rdp_start_all():
        """Just launch start_all_clean.py. It handles everything."""
        import sys as _sys
        import subprocess as _sp

        script_path = project_root / "start_all_clean.py"
        if not script_path.exists():
            return jsonify({"success": False, "error": "start_all_clean.py not found"}), 404

        python_exe = Path(_sys.executable).parent / "python.exe"
        if not python_exe.exists():
            import shutil as _shutil
            _found = _shutil.which("python.exe")
            python_exe = Path(_found) if _found else Path("python.exe")

        log_file = Path(r"C:\ProgramData\MHTAgentic\start_all_clean.log")

        try:
            proc = _sp.Popen(
                [str(python_exe), str(script_path)],
                stdout=open(str(log_file), "w"),
                stderr=_sp.STDOUT,
                cwd=str(project_root),
            )
            logger.info(f"[Start All] start_all_clean.py launched (pid={proc.pid})")
            return jsonify({"success": True, "pid": proc.pid})
        except Exception as e:
            logger.error(f"[Start All] Failed: {e}")
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/bot-errors")
    def api_bot_errors():
        """Return recent bot login/startup errors from the DB."""
        slot = request.args.get("slot", None)
        limit = int(request.args.get("limit", 50))
        if db_path.exists():
            try:
                errors = get_bot_errors(str(db_path), slot_name=slot, limit=limit)
            except Exception as e:
                logger.error(f"Bot errors query failed: {e}")
                errors = []
        else:
            errors = []
        return jsonify({"errors": errors})

    @app.route("/api/rdp/stop/<session_id>", methods=["POST"])
    def api_rdp_stop(session_id: str):
        """Stop a specific RDP session: kill bots, logoff, close mstsc."""
        from dashboard.session_monitor import (
            _find_rdp_session_ids, _kill_bot_processes_in_session,
            _logoff_sessions_elevated,
        )
        try:
            hwnd = int(session_id)
        except ValueError:
            return jsonify({"success": False, "error": "Invalid session id"}), 400

        # Find the WTS session ID for this hwnd
        rdp_sessions = _find_rdp_session_ids()
        # Try to match by finding which session owns this window
        # If session_id looks like a WTS session ID (small number), use it directly
        # Otherwise treat as hwnd and fall back to stop_rdp_session
        if hwnd < 1000:
            # Likely a WTS session ID
            _kill_bot_processes_in_session(hwnd)
            import time; time.sleep(1)
            _logoff_sessions_elevated([hwnd])
            return jsonify({"success": True, "session_id": hwnd})
        else:
            # hwnd — just close the mstsc window
            result = stop_rdp_session(hwnd)
            return jsonify(result)

    @app.route("/api/rdp/stop-all", methods=["POST"])
    def api_rdp_stop_all():
        """Logoff all RDP sessions (kills apps inside), then close mstsc.
        Force-resets all bot slots, clears error logs, and verifies clean state.
        NUCLEAR: immediately aborts any running start-all, marks all bots aborted."""
        logger.info("[Stop All] ========== STOP ALL TRIGGERED ==========")

        # 1. Signal any running start-all thread to abort IMMEDIATELY
        _start_all_abort["abort"] = True
        logger.info("[Stop All] Abort flag set — start-all thread will exit at next check")

        # 2. Immediately mark ALL bots as aborted so no more work is queued
        was_running = _start_all_status.get("running", False)
        if was_running:
            logger.info("[Stop All] Start-all was running — marking all bots as aborted")
            for bot_name, bot_status in _start_all_status.get("bots", {}).items():
                if bot_status.get("status") not in ("running", "complete"):
                    bot_status["status"] = "aborted"
                    bot_status["error"] = "Cancelled by Stop All"
                    bot_status["step"] = "aborted"
                    logger.info(f"[Stop All] Marked {bot_name} as aborted")
        _start_all_status["running"] = False

        # 3. Stop all RDP sessions (elevated kill + logoff)
        result = stop_all_rdp_sessions(rdp_search_dirs)

        # 4. Final sweep: force-reset any slots that survived (stale_seconds=0)
        extra_cleaned = 0
        errors_cleared = 0
        if db_path.exists():
            try:
                extra_cleaned = cleanup_stale_slots(str(db_path), stale_seconds=0)
                if extra_cleaned:
                    logger.info(f"[Stop All] Extra slot cleanup: {extra_cleaned}")
            except Exception as e:
                logger.error(f"[Stop All] Post-stop slot cleanup failed: {e}")

            # Clear bot error logs — fresh session starts clean
            try:
                errors_cleared = clear_bot_errors(str(db_path))
                if errors_cleared:
                    logger.info(f"[Stop All] Cleared {errors_cleared} bot error log entries")
            except Exception as e:
                logger.error(f"[Stop All] Failed to clear bot errors: {e}")

        # 5. Reset start-all status tracker completely
        _start_all_status["running"] = False
        _start_all_status["bots"] = {}
        _start_all_status["finished_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")

        result["slots_reset"] = result.get("slots_reset", 0) + extra_cleaned
        result["errors_cleared"] = errors_cleared
        result["start_all_was_running"] = was_running
        logger.info(f"[Stop All] ========== DONE: {result} ==========")
        return jsonify(result)

    @app.route("/api/start-monitoring", methods=["POST"])
    def api_start_monitoring():
        """Disabled — use Start All (start_all_clean.py) instead."""
        return jsonify({"success": False, "error": "Use Start All instead"}), 400

    @app.route("/api/clear-demo-db", methods=["POST"])
    def api_clear_demo_db():
        """Delete all rows from common_event, common_eventerror, and bot_error_log."""
        if not db_path.exists():
            return jsonify({"success": True, "deleted": 0})
        try:
            import sqlite3
            conn = sqlite3.connect(str(db_path), timeout=5)
            cur = conn.cursor()
            cur.execute("PRAGMA journal_mode=OFF")
            cur.execute("SELECT COUNT(*) FROM common_event")
            count = cur.fetchone()[0]
            cur.execute("DELETE FROM common_eventerror")
            cur.execute("DELETE FROM common_event")
            cur.execute("DELETE FROM bot_error_log")
            # Reset bot slots back to open
            cur.execute("UPDATE bot_slot SET status = 'open', session_id = NULL, claimed_at = NULL, heartbeat_at = NULL")
            conn.commit()
            conn.close()

            # Reset start-all status tracker
            _start_all_status["running"] = False
            _start_all_status["bots"] = {}

            logger.info(f"Cleared demo DB: {count} events deleted + bot errors cleared")
            return jsonify({"success": True, "deleted": count})
        except Exception as e:
            logger.error(f"Failed to clear demo DB: {e}")
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/locations")
    def api_locations():
        from mhtagentic.db import get_all_locations
        if db_path.exists():
            try:
                locations = get_all_locations(str(db_path))
            except Exception as e:
                logger.error(f"Locations query failed: {e}")
                locations = []
        else:
            locations = []
        return jsonify({"locations": locations})

    @app.route("/api/locations/<name>/toggle", methods=["POST"])
    def api_toggle_location(name):
        from mhtagentic.db import toggle_location_active, get_all_locations
        data = request.get_json(silent=True) or {}
        active = data.get("active", True)
        success = toggle_location_active(str(db_path), name, active)
        locations = get_all_locations(str(db_path)) if db_path.exists() else []
        return jsonify({"success": success, "locations": locations})

    @app.route("/api/locations/add", methods=["POST"])
    def api_add_location():
        from mhtagentic.db import add_location, get_all_locations
        data = request.get_json(silent=True) or {}
        name = data.get("name", "").strip().upper()
        if not name:
            return jsonify({"success": False, "error": "No name"}), 400
        success = add_location(str(db_path), name)
        locations = get_all_locations(str(db_path)) if db_path.exists() else []
        return jsonify({"success": success, "locations": locations})

    @app.route("/api/config")
    def api_config():
        from mhtagentic.db import get_config
        if db_path.exists():
            try:
                config = get_config(str(db_path))
            except Exception:
                config = {}
        else:
            config = {}
        return jsonify(config)

    @app.route("/api/config", methods=["POST"])
    def api_config_update():
        from mhtagentic.db import set_config, get_config
        data = request.get_json(silent=True) or {}
        for k, v in data.items():
            set_config(str(db_path), k, str(v))
        config = get_config(str(db_path)) if db_path.exists() else {}
        return jsonify({"success": True, "config": config})

    @app.route("/api/rdp/reboot/<slot_name>", methods=["POST"])
    def api_reboot_bot(slot_name):
        """Kill a specific bot's processes and relaunch it."""
        if db_path.exists():
            try:
                release_slot(str(db_path), slot_name.lower())
            except Exception:
                pass

        # All RDPs use the same username — find session by matching slot name in window titles
        from dashboard.session_monitor import (
            _find_rdp_session_ids, _kill_bot_processes_in_session,
            _logoff_sessions_elevated,
        )
        rdp_sessions = _find_rdp_session_ids()
        # Match by slot name (e.g. "experityb" matches session with username containing it)
        target_sids = [s["session_id"] for s in rdp_sessions
                       if slot_name.lower() in s.get("username", "").lower()
                       or slot_name.lower() in s.get("station", "").lower()]

        if not target_sids:
            # Fallback: log off by username directly
            _logoff_rdp_user(slot_name)
        else:
            for sid in target_sids:
                _kill_bot_processes_in_session(sid)
            import time; time.sleep(1)
            _logoff_sessions_elevated(target_sids)

        return jsonify({"success": True, "message": f"Rebooting {slot_name}..."})

    return app
