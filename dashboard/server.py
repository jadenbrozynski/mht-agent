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
        """Return current bot slot status (auto-releases stale slots)."""
        if db_path.exists():
            try:
                cleanup_stale_slots(str(db_path))
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
        """Start an RDP session from an .rdp file."""
        data = request.get_json(silent=True) or {}
        rdp_path = data.get("path", "")
        if not rdp_path:
            return jsonify({"success": False, "error": "No path provided"}), 400
        result = start_rdp_session(rdp_path, project_root=project_root)
        return jsonify(result)

    MAX_LOGIN_ATTEMPTS = 5
    SESSION_APPEAR_TIMEOUT = 120  # seconds: wait for new Windows session after mstsc opens
    BOT_READY_TIMEOUT = 180      # seconds: wait for Google Auth + bot fully ready (signal file)

    @app.route("/api/rdp/start-all", methods=["POST"])
    def api_rdp_start_all():
        """Start exactly 3 RDP sessions: ExperityB → ExperityD → ExperityC.
        Each gets a fresh logon (logoff first), startup task fires, OTP chained.
        Retries each bot up to MAX_LOGIN_ATTEMPTS times before moving on.
        Each attempt has 2-phase OTP wait: Task Scheduler (45s) → PsExec fallback (75s)."""
        # Hardcoded launch order — exactly 3 sessions, no extras
        desktop = Path.home() / "Desktop"
        launch_order = [
            {"username": "ExperityB", "path": str(desktop / "ExperityB_MHT.rdp"), "location": "ANNISTON"},
            {"username": "ExperityD", "path": str(desktop / "ExperityD_MHT.rdp"), "location": "ATTALLA"},
            {"username": "ExperityC", "path": str(desktop / "ExperityC_MHT.rdp"), "location": "ATTALLA"},
        ]

        # Verify all .rdp files exist
        for entry in launch_order:
            if not Path(entry["path"]).exists():
                return jsonify({"success": False, "error": f"Missing: {entry['path']}"}), 400

        # Prevent double-start
        if _start_all_status.get("running"):
            return jsonify({"success": False, "error": "Start All already in progress"}), 409

        # Clear stale OTP signals before starting
        clear_all_otp_signals()

        # Clear previous error logs for a fresh session
        if db_path.exists():
            try:
                clear_bot_errors(str(db_path))
            except Exception:
                pass

        # Initialize status tracker + reset abort flag
        _start_all_status["running"] = True
        _start_all_status["started_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
        _start_all_abort["abort"] = False
        _start_all_status["bots"] = {}
        for entry in launch_order:
            _start_all_status["bots"][entry["username"]] = {
                "status": "pending",
                "attempt": 0,
                "max_attempts": MAX_LOGIN_ATTEMPTS,
                "error": "",
                "step": "",
            }

        def _is_aborted():
            """Check if stop-all has been triggered."""
            return _start_all_abort.get("abort", False)

        # Launch in a background thread so the HTTP response returns immediately
        def _sequential_start():
            import time as _time
            import traceback as _tb

            def _abortable_sleep(seconds: float):
                """Sleep in 0.5s increments, checking abort flag each tick."""
                deadline = _time.time() + seconds
                while _time.time() < deadline:
                    if _is_aborted():
                        return
                    _time.sleep(min(0.5, deadline - _time.time()))

            def _wait_for_rdp_window_gone(target_username: str, timeout: int = 15):
                """Wait until no mstsc window exists for target_username."""
                deadline = _time.time() + timeout
                while _time.time() < deadline:
                    if _is_aborted():
                        return True  # don't block on abort
                    windows = find_rdp_windows()
                    still_open = any(
                        target_username.lower() in w["title"].lower()
                        for w in windows
                    )
                    if not still_open:
                        return True
                    _time.sleep(1)
                return False

            # ── Step 0: Ensure scheduled tasks exist (auto-repair) ──
            logger.info("[Start All] ========== STARTING ==========")
            logger.info("[Start All] Verifying scheduled tasks...")
            try:
                task_result = ensure_autostart_task(project_root)
                created = task_result.get("created", [])
                if created:
                    logger.info(f"[Start All] Created missing scheduled tasks: {created}")
                    _abortable_sleep(3)
                else:
                    logger.info("[Start All] All scheduled tasks already exist")
            except Exception as e:
                logger.error(f"[Start All] ensure_autostart_task failed: {e}")

            if _is_aborted():
                logger.info("[Start All] ABORTED before pre-check")
                _start_all_status["running"] = False
                return

            # ── Step 0b: Check which bots are already running healthy ──
            already_running = set()
            try:
                slots = get_slots(str(db_path))
                current_windows = find_rdp_windows()
                logger.info(
                    f"[Start All] Pre-check: {len(slots)} slots, "
                    f"{len(current_windows)} RDP windows visible"
                )
                for s in slots:
                    logger.info(
                        f"[Start All]   Slot: {s['slot_name']} status={s['status']} "
                        f"heartbeat={s.get('heartbeat_at')} session_id={s.get('session_id')}"
                    )
                for w in current_windows:
                    logger.info(
                        f"[Start All]   Window: hwnd={w['hwnd']} title='{w['title']}' "
                        f"size={w['width']}x{w['height']}"
                    )

                for entry in launch_order:
                    uname = entry["username"]
                    slot_name = uname.lower()
                    slot_info = next(
                        (s for s in slots if s["slot_name"] == slot_name),
                        None,
                    )
                    if slot_info and slot_info["status"] == "active" and slot_info.get("heartbeat_at"):
                        from datetime import datetime
                        try:
                            hb = datetime.fromisoformat(slot_info["heartbeat_at"])
                            age = (datetime.now() - hb).total_seconds()
                            has_window = any(
                                uname.lower() in w["title"].lower()
                                for w in current_windows
                            )
                            logger.info(
                                f"[Start All]   {uname}: slot=active, heartbeat_age={age:.0f}s, "
                                f"has_window={has_window}"
                            )
                            if age < 60 and has_window:
                                already_running.add(uname)
                                logger.info(f"[Start All]   → SKIPPING {uname} (already healthy)")
                            else:
                                logger.info(
                                    f"[Start All]   → NOT skipping {uname} "
                                    f"(age={age:.0f}s, window={has_window})"
                                )
                        except (ValueError, TypeError) as e:
                            logger.warning(f"[Start All]   {uname}: heartbeat parse error: {e}")
                    else:
                        status = slot_info["status"] if slot_info else "no_slot"
                        logger.info(f"[Start All]   {uname}: slot={status} — will launch")
            except Exception as e:
                logger.error(f"[Start All] Pre-check for running bots failed: {e}")

            logger.info(f"[Start All] Already running (will skip): {already_running or 'none'}")
            logger.info(f"[Start All] Launch order: {[e['username'] for e in launch_order]}")

            for i, entry in enumerate(launch_order):
                if _is_aborted():
                    logger.info("[Start All] ABORTED between bots")
                    for remaining in launch_order[i:]:
                        rname = remaining["username"]
                        if rname in _start_all_status["bots"]:
                            _start_all_status["bots"][rname]["status"] = "aborted"
                            _start_all_status["bots"][rname]["error"] = "Cancelled by Stop All"
                    break

                username = entry["username"]
                rdp_path = entry["path"]
                slot_name = username.lower()
                bot_status = _start_all_status["bots"][username]

                logger.info(f"[Start All] ── Bot {i+1}/3: {username} ──")

                # Skip bots that are already running healthy
                if username in already_running:
                    bot_status["status"] = "running"
                    bot_status["step"] = "already_running"
                    bot_status["attempt"] = 0
                    bot_status["error"] = ""
                    logger.info(f"[Start All] {username} → skipped (already running)")
                    continue

                login_success = False

                for attempt in range(1, MAX_LOGIN_ATTEMPTS + 1):
                    if _is_aborted():
                        logger.info(f"[Start All] {username} ABORTED at attempt {attempt}")
                        bot_status["status"] = "aborted"
                        bot_status["error"] = "Cancelled by Stop All"
                        login_success = True  # prevent "failed" status
                        break

                    bot_status["attempt"] = attempt
                    bot_status["status"] = "starting"
                    bot_status["step"] = "logoff"
                    bot_status["error"] = ""

                    logger.info(
                        f"[Start All] {username} — attempt {attempt}/{MAX_LOGIN_ATTEMPTS}"
                    )

                    try:
                        # 1. Log off any existing session
                        bot_status["step"] = "logoff"
                        logger.info(f"[Start All] {username}: logging off existing session...")
                        _logoff_rdp_user(username)

                        if _is_aborted(): break

                        # 3. Wait for the old RDP window to close
                        bot_status["step"] = "wait_window_close"
                        logger.info(f"[Start All] {username}: waiting for old window to close...")
                        gone = _wait_for_rdp_window_gone(username, timeout=15)
                        if not gone:
                            logger.warning(f"[Start All] {username}: old window didn't close in 15s")
                        else:
                            logger.info(f"[Start All] {username}: old window closed")
                        _abortable_sleep(2)

                        if _is_aborted(): break

                        # 4. Clear OTP signal
                        signal_path = Path(r"C:\ProgramData\MHTAgentic\session_status") / f"{username}_otp_complete"
                        try:
                            signal_path.unlink(missing_ok=True)
                            logger.info(f"[Start All] {username}: cleared OTP signal")
                        except Exception:
                            pass

                        # 5. Snapshot existing session IDs BEFORE opening mstsc
                        #    so we can detect the NEW session that appears after.
                        #    (WTS may report wrong username during Google Auth,
                        #    so we match by "new session ID" not by username)
                        pre_session_ids = set(
                            s["session_id"] for s in _find_rdp_session_ids()
                        )
                        logger.info(
                            f"[Start All] {username}: existing session IDs before open: "
                            f"{pre_session_ids}"
                        )

                        # Open the RDP connection
                        bot_status["step"] = "rdp_connect"
                        logger.info(f"[Start All] {username}: opening RDP from {rdp_path}...")
                        result = start_rdp_session(rdp_path, project_root=project_root)
                        if not result.get("success"):
                            raise RuntimeError(
                                f"mstsc.exe failed to launch: {result.get('error', 'unknown')}"
                            )
                        logger.info(f"[Start All] {username}: mstsc started (pid={result.get('pid')})")

                        if _is_aborted(): break

                        # 6. Wait for a NEW Windows session to appear
                        #    (one that wasn't in the pre-snapshot).
                        bot_status["step"] = "wait_session"
                        logger.info(
                            f"[Start All] {username}: waiting up to "
                            f"{SESSION_APPEAR_TIMEOUT}s for NEW session "
                            f"(ignoring existing: {pre_session_ids})..."
                        )

                        sid = None
                        _wait_deadline = _time.time() + SESSION_APPEAR_TIMEOUT
                        _poll_n = 0
                        while _time.time() < _wait_deadline:
                            if _is_aborted(): break
                            all_sessions = _find_rdp_session_ids()
                            new_sessions = [
                                s for s in all_sessions
                                if s["session_id"] not in pre_session_ids
                            ]
                            _poll_n += 1
                            if _poll_n <= 3 or _poll_n % 5 == 0:
                                logger.info(
                                    f"[Start All] {username}: poll #{_poll_n} — "
                                    f"all={[(s['username'], s['session_id']) for s in all_sessions]}, "
                                    f"new={[(s['username'], s['session_id']) for s in new_sessions]}"
                                )
                            if new_sessions:
                                sid = new_sessions[0]["session_id"]
                                logger.info(
                                    f"[Start All] {username}: NEW session found! "
                                    f"session_id={sid} (WTS user={new_sessions[0]['username']})"
                                )
                                break
                            _time.sleep(2)

                        if _is_aborted(): break

                        if sid is None:
                            all_sessions = _find_rdp_session_ids()
                            all_windows = find_rdp_windows()
                            logger.error(
                                f"[Start All] {username}: NO NEW SESSION! "
                                f"All sessions: {[(s['username'], s['session_id']) for s in all_sessions]}, "
                                f"All windows: {[(w['title'], w['hwnd']) for w in all_windows]}, "
                                f"Pre-existing IDs: {pre_session_ids}"
                            )
                            raise RuntimeError(
                                f"No new Windows session appeared for {username} "
                                f"after {SESSION_APPEAR_TIMEOUT}s — RDP may have "
                                f"failed to connect or needs manual credential entry."
                            )

                        # 7. PsExec inject the bot into this session.
                        #    Task Scheduler may also start a bot on first logon —
                        #    that's fine, whichever claims the slot first wins.
                        #    For 2nd/3rd RDP (same account), Task Scheduler
                        #    won't fire, so PsExec is the only way.
                        bot_status["step"] = "inject_bot"
                        logger.info(
                            f"[Start All] {username}: session {sid} found. "
                            f"Injecting bot via PsExec (FORCE_BOT_USER={username})..."
                        )
                        launch_result = launch_bot_in_session(
                            sid, project_root, username=username,
                            location=entry.get("location", "ATTALLA")
                        )
                        logger.info(
                            f"[Start All] {username}: PsExec result: {launch_result}"
                        )

                        if _is_aborted(): break

                        # 8. Wait for bot to become ready.
                        #    Check signal file AND DB slot — either means success.
                        #    Task Scheduler bot may claim slot for first session,
                        #    PsExec bot claims slot for subsequent sessions.
                        bot_status["step"] = "wait_bot_ready"
                        signal_file = Path(r"C:\ProgramData\MHTAgentic\session_status") / f"{username}_otp_complete"
                        logger.info(
                            f"[Start All] {username}: waiting up to {BOT_READY_TIMEOUT}s "
                            f"for bot ready (signal file OR DB slot)..."
                        )
                        _otp_deadline = _time.time() + BOT_READY_TIMEOUT
                        otp_ok = False
                        _otp_poll = 0
                        while _time.time() < _otp_deadline:
                            if _is_aborted(): break
                            _otp_poll += 1

                            # Check 1: OTP signal file
                            if signal_file.exists():
                                logger.info(
                                    f"[Start All] {username}: OTP signal FILE "
                                    f"found after {_otp_poll * 3}s!"
                                )
                                otp_ok = True
                                break

                            # Check 2: DB slot active
                            try:
                                _diag = get_slots(str(db_path))
                                slot_info = next(
                                    (s for s in _diag if s["slot_name"] == slot_name),
                                    None,
                                )
                                if slot_info and slot_info["status"] == "active":
                                    logger.info(
                                        f"[Start All] {username}: DB slot ACTIVE "
                                        f"after {_otp_poll * 3}s "
                                        f"(PID={slot_info.get('session_id')})"
                                    )
                                    otp_ok = True
                                    break
                            except Exception:
                                pass

                            if _otp_poll % 10 == 0:
                                slot_st = "?"
                                try:
                                    _diag = get_slots(str(db_path))
                                    slot_st = next(
                                        (s["status"] for s in _diag if s["slot_name"] == slot_name),
                                        "?"
                                    )
                                except Exception:
                                    pass
                                logger.info(
                                    f"[Start All] {username}: still waiting "
                                    f"({_otp_poll * 3}s elapsed, slot={slot_st}, "
                                    f"signal_file={signal_file.exists()})..."
                                )
                            _time.sleep(3)

                        if _is_aborted(): break

                        if otp_ok:
                            bot_status["status"] = "running"
                            bot_status["step"] = "complete"
                            bot_status["error"] = ""
                            logger.info(
                                f"[Start All] {username}: ✓ BOT READY "
                                f"(attempt {attempt}, session {sid})"
                            )
                            login_success = True
                            break
                        else:
                            diag_slots = get_slots(str(db_path))
                            diag_sessions = _find_rdp_session_ids()
                            logger.error(
                                f"[Start All] {username}: bot NOT ready after {BOT_READY_TIMEOUT}s! "
                                f"Signal: {signal_file.exists()}, "
                                f"Slots: {[(s['slot_name'], s['status']) for s in diag_slots]}, "
                                f"Sessions: {[(s['username'], s['session_id']) for s in diag_sessions]}"
                            )
                            raise RuntimeError(
                                f"Bot did not become active after {BOT_READY_TIMEOUT}s. "
                                f"Task Scheduler may not have started the bot."
                            )

                    except Exception as exc:
                        # If aborted, don't log errors or retry — just bail
                        if _is_aborted():
                            logger.info(f"[Start All] {username}: aborted during exception handler")
                            bot_status["status"] = "aborted"
                            bot_status["error"] = "Cancelled by Stop All"
                            login_success = True
                            break

                        error_msg = str(exc)
                        tb_str = _tb.format_exc()
                        step = bot_status.get("step", "unknown")

                        logger.error(
                            f"[Start All] {username} — attempt {attempt}/{MAX_LOGIN_ATTEMPTS} "
                            f"FAILED at step '{step}': {error_msg}"
                        )
                        logger.error(f"[Start All] {username} — traceback:\n{tb_str}")

                        # Only log to DB if not aborted
                        if not _is_aborted():
                            try:
                                log_bot_error(
                                    str(db_path),
                                    slot_name=slot_name,
                                    attempt=attempt,
                                    max_attempts=MAX_LOGIN_ATTEMPTS,
                                    error=error_msg,
                                    tb=tb_str,
                                    step=step,
                                )
                            except Exception as db_exc:
                                logger.error(f"Failed to log bot error to DB: {db_exc}")

                        bot_status["error"] = error_msg
                        bot_status["status"] = (
                            "retrying" if attempt < MAX_LOGIN_ATTEMPTS else "failed"
                        )

                        if attempt < MAX_LOGIN_ATTEMPTS:
                            if _is_aborted(): break
                            logger.info(
                                f"[Start All] {username} — retrying in 5s "
                                f"(attempt {attempt + 1}/{MAX_LOGIN_ATTEMPTS})"
                            )
                            _abortable_sleep(5)

                if _is_aborted():
                    if bot_status["status"] not in ("running", "aborted"):
                        bot_status["status"] = "aborted"
                        bot_status["error"] = "Cancelled by Stop All"
                    continue

                if not login_success and username not in already_running:
                    bot_status["status"] = "failed"
                    if not bot_status["error"]:
                        bot_status["error"] = (
                            f"All {MAX_LOGIN_ATTEMPTS} attempts exhausted."
                        )

            _start_all_status["running"] = False
            _start_all_status["finished_at"] = _time.strftime("%Y-%m-%dT%H:%M:%S")

            # Summary
            succeeded = sum(
                1 for b in _start_all_status["bots"].values()
                if b["status"] in ("running",)
            )
            failed = sum(
                1 for b in _start_all_status["bots"].values()
                if b["status"] == "failed"
            )
            aborted = sum(
                1 for b in _start_all_status["bots"].values()
                if b["status"] == "aborted"
            )
            logger.info(
                f"[Start All] ========== DONE: {succeeded} running, "
                f"{failed} failed, {aborted} aborted =========="
            )

        thread = threading.Thread(target=_sequential_start, daemon=True)
        thread.start()

        return jsonify({
            "results": [{"success": True, "message": "Sequential start initiated with retry"}],
            "sequential": True,
            "file_count": len(launch_order),
            "max_attempts": MAX_LOGIN_ATTEMPTS,
        })

    @app.route("/api/start-all-status")
    def api_start_all_status():
        """Return the current start-all progress (per-bot attempt status)."""
        return jsonify(_start_all_status)

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
        """Stop a specific RDP session by hwnd."""
        try:
            hwnd = int(session_id)
        except ValueError:
            return jsonify({"success": False, "error": "Invalid session id"}), 400
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
        """Launch monitor-only mode inside each active RDP session."""
        result = start_monitoring_in_sessions(project_root)
        return jsonify(result)

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

    return app
