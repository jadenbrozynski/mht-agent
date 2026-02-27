/**
 * MHT Agentic Dashboard - Polling & rendering logic.
 *
 * Sessions polled every 5s (screenshots update only on state change server-side).
 * Stats/events polled every 10s.
 * DOM updates are incremental to avoid page jolt.
 */

(function () {
  "use strict";

  var SESSION_INTERVAL = 5000;
  var STATS_INTERVAL = 10000;

  // Track current state to avoid unnecessary DOM rewrites
  var _currentSessionIds = [];
  var _noSessionsShown = false;

  // --- Helpers ---

  function $(sel) {
    return document.querySelector(sel);
  }

  function setText(id, val) {
    var el = document.getElementById(id);
    if (el && el.textContent !== String(val)) el.textContent = val;
  }

  function formatTime(iso) {
    if (!iso) return "--";
    var d = new Date(iso);
    return d.toLocaleTimeString();
  }

  function statusLabel(code) {
    if (code >= 100) return { text: "Complete", cls: "complete" };
    if (code >= 40) return { text: "Sent", cls: "processing" };
    if (code >= 10) return { text: "Converted", cls: "processing" };
    if (code === 0) return { text: "Pending", cls: "pending" };
    if (code < 0) return { text: "Failed", cls: "failed" };
    return { text: String(code), cls: "pending" };
  }

  function escapeHtml(str) {
    var div = document.createElement("div");
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
  }

  function postJSON(url) {
    return fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    }).then(function (r) { return r.json(); });
  }

  // --- RDP Control ---

  var _startAllPollTimer = null;

  window.startAllRDP = function () {
    var btn = document.getElementById("btn-start-all");
    btn.disabled = true;
    btn.textContent = "Starting...";

    postJSON("/api/rdp/start-all")
      .then(function (data) {
        if (data.error) {
          btn.textContent = data.error;
          setTimeout(function () {
            btn.textContent = "Start All";
            btn.disabled = false;
          }, 3000);
          return;
        }
        btn.textContent = "Starting " + data.file_count + " sessions (max " + data.max_attempts + " retries each)...";
        // Start polling start-all status
        _startPollStartAll();
      })
      .catch(function () {
        btn.textContent = "Start All";
        btn.disabled = false;
      });
  };

  function _startPollStartAll() {
    if (_startAllPollTimer) clearInterval(_startAllPollTimer);
    _startAllPollTimer = setInterval(refreshStartAllStatus, 2000);
    refreshStartAllStatus();
  }

  window.stopAllRDP = function () {
    var btn = document.getElementById("btn-stop-all");
    btn.disabled = true;
    btn.textContent = "Killing processes...";

    // Show progress steps while waiting for the backend
    var progressTimer = setTimeout(function () {
      btn.textContent = "Resetting slots...";
    }, 2000);

    postJSON("/api/rdp/stop-all")
      .then(function (data) {
        clearTimeout(progressTimer);
        var killed = data.killed || 0;
        var slotsReset = data.slots_reset || 0;
        var clean = data.verified_clean;

        // Clear start-all progress and error panels immediately
        if (_startAllPollTimer) {
          clearInterval(_startAllPollTimer);
          _startAllPollTimer = null;
        }
        document.getElementById("start-all-panel").style.display = "none";
        document.getElementById("start-all-bots").innerHTML = "";
        document.getElementById("bot-errors-panel").style.display = "none";
        document.getElementById("bot-errors-list").innerHTML = "";

        // Re-enable start button
        var startBtn = document.getElementById("btn-start-all");
        startBtn.textContent = "Start All";
        startBtn.disabled = false;

        if (clean === false) {
          btn.textContent = "Warning: cleanup incomplete";
          btn.classList.add("btn-warn");
          setTimeout(function () {
            btn.textContent = "Stop All";
            btn.classList.remove("btn-warn");
            btn.disabled = false;
          }, 4000);
        } else {
          btn.textContent = "Done \u2014 " + killed + " killed, " + slotsReset + " slots reset";
          setTimeout(function () {
            btn.textContent = "Stop All";
            btn.disabled = false;
          }, 3000);
        }

        // Refresh everything
        refreshSlots();
        refreshBotErrors();
        refreshStartAllStatus();
      })
      .catch(function () {
        clearTimeout(progressTimer);
        btn.textContent = "Stop All";
        btn.disabled = false;
      });
  };

  window.clearDemoDB = function () {
    if (!confirm("Clear all inbound and outbound events from the demo DB?")) return;
    var btn = document.getElementById("btn-clear-db");
    btn.disabled = true;
    btn.textContent = "Clearing...";

    postJSON("/api/clear-demo-db")
      .then(function (data) {
        if (data.success) {
          btn.textContent = "Cleared " + (data.deleted || 0) + " events";
          refreshStats();
          refreshEvents();
        } else {
          btn.textContent = data.error || "Failed";
        }
        setTimeout(function () {
          btn.textContent = "Clear Demo DB";
          btn.disabled = false;
        }, 3000);
      })
      .catch(function () {
        btn.textContent = "Clear Demo DB";
        btn.disabled = false;
      });
  };

  window.startMonitoring = function () {
    var btn = document.getElementById("btn-start-monitor");
    btn.disabled = true;
    btn.textContent = "Launching...";

    postJSON("/api/start-monitoring")
      .then(function (data) {
        if (data.success) {
          var sessions = data.sessions || [];
          var names = sessions.map(function (s) { return s.username + " (" + s.mode + ")"; });
          btn.textContent = "Started: " + names.join(", ");
        } else {
          btn.textContent = data.error || "Failed";
        }
        setTimeout(function () {
          btn.textContent = "Start Monitoring";
          btn.disabled = false;
        }, 4000);
      })
      .catch(function () {
        btn.textContent = "Start Monitoring";
        btn.disabled = false;
      });
  };

  window.stopSession = function (sessionId) {
    postJSON("/api/rdp/stop/" + sessionId);
  };

  window.startRDP = function (path) {
    fetch("/api/rdp/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: path }),
    });
  };

  // --- Sessions + Screenshots (incremental DOM updates) ---

  function refreshSessions() {
    fetch("/api/sessions")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        var container = $("#sessions-container");
        var health = data.bot_health || {};

        // Update header status badge
        var badge = $("#bot-status");
        var statusText = (health.status || "unknown").replace("_", " ");
        if (badge.textContent !== statusText) {
          badge.textContent = statusText;
          badge.className = "status-badge " + (health.status || "unknown");
        }

        var sessions = data.sessions || [];
        var hasReal = sessions.some(function (s) { return s.id !== "none"; });

        if (!hasReal) {
          // Only rewrite if we weren't already showing the placeholder
          if (!_noSessionsShown) {
            _noSessionsShown = true;
            _currentSessionIds = [];
            fetch("/api/rdp/files")
              .then(function (r) { return r.json(); })
              .then(function (fileData) {
                var files = fileData.files || [];
                var html = '<div class="session-card placeholder">';
                html += "<p>No RDP sessions detected</p>";
                if (files.length > 0) {
                  html += '<div class="rdp-file-list">';
                  for (var i = 0; i < files.length; i++) {
                    html +=
                      '<button class="btn btn-green btn-sm" onclick="startRDP(\'' +
                      escapeHtml(files[i].path.replace(/\\/g, "\\\\").replace(/'/g, "\\'")) +
                      "')\">" +
                      escapeHtml(files[i].name) +
                      "</button> ";
                  }
                  html += "</div>";
                }
                html += "</div>";
                container.innerHTML = html;
              });
          }
          return;
        }

        _noSessionsShown = false;

        // Build list of new session IDs
        var newIds = sessions
          .filter(function (s) { return s.id !== "none"; })
          .map(function (s) { return s.id; });

        // Check if session list structure changed
        var structureChanged =
          newIds.length !== _currentSessionIds.length ||
          newIds.some(function (id, i) { return id !== _currentSessionIds[i]; });

        if (structureChanged) {
          // Full rebuild needed (session added/removed)
          var html = "";
          for (var i = 0; i < sessions.length; i++) {
            var s = sessions[i];
            if (s.id === "none") continue;

            html += '<div class="session-card" data-session="' + s.id + '">';
            html += '  <div class="session-header">';
            html += '    <span class="session-title">' + escapeHtml(s.title);
            if (s.agent_role === "inbound") {
              html += ' <span class="role-badge role-inbound">INBOUND</span>';
            } else if (s.agent_role === "outbound") {
              html += ' <span class="role-badge role-outbound">OUTBOUND</span>';
            }
            html += "</span>";
            html += '    <div class="session-header-right">';
            html += '      <span class="session-meta">' + s.width + "x" + s.height + "</span>";
            html +=
              '      <button class="btn btn-red btn-sm" onclick="stopSession(\'' +
              s.id + "')\">Close</button>";
            html += "    </div>";
            html += "  </div>";
            html += '  <div class="session-screenshot">';
            if (s.has_screenshot) {
              html +=
                '    <img src="/api/screenshots/' + s.id + "?t=" + Date.now() +
                '" alt="RDP Screenshot" loading="lazy">';
            } else {
              html += '    <span class="no-screenshot">No screenshot available</span>';
            }
            html += "  </div>";
            html += "</div>";
          }
          container.innerHTML = html;
          _currentSessionIds = newIds;
        } else {
          // Same sessions — just update screenshot src (no reflow)
          for (var j = 0; j < sessions.length; j++) {
            var s2 = sessions[j];
            if (s2.id === "none") continue;

            var card = container.querySelector('[data-session="' + s2.id + '"]');
            if (!card) continue;

            if (s2.has_screenshot) {
              var img = card.querySelector("img");
              var newSrc = "/api/screenshots/" + s2.id + "?t=" + Date.now();
              if (img) {
                img.src = newSrc;
              } else {
                // Had no-screenshot, now has one
                var wrap = card.querySelector(".session-screenshot");
                wrap.innerHTML = '<img src="' + newSrc + '" alt="RDP Screenshot" loading="lazy">';
              }
            }
          }
        }
      })
      .catch(function (err) {
        console.error("Sessions fetch error:", err);
      });
  }

  // --- Analytics ---

  function refreshStats() {
    fetch("/api/analytics/current")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        var live = data.live || {};
        var daily = data.daily_stats || {};

        setText("stat-events", live.events_today || 0);
        setText("stat-inbound", live.inbound_today || 0);
        setText("stat-outbound", live.outbound_today || 0);
        setText("stat-errors", live.errors_today || 0);

        setText(
          "stat-patients",
          daily.total_patients != null ? daily.total_patients : "--"
        );
        setText(
          "stat-avg-time",
          daily.avg_time_per_patient_seconds != null
            ? daily.avg_time_per_patient_seconds.toFixed(1)
            : "--"
        );

        setText("last-update", new Date().toLocaleTimeString());
      })
      .catch(function (err) {
        console.error("Stats fetch error:", err);
      });
  }

  // --- Events Table ---

  function refreshEvents() {
    fetch("/api/analytics/events")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        var tbody = $("#events-body");
        var events = data.events || [];

        if (events.length === 0) {
          tbody.innerHTML =
            '<tr><td colspan="6" class="empty">No events yet</td></tr>';
          return;
        }

        var html = "";
        for (var i = 0; i < events.length; i++) {
          var e = events[i];
          var st = statusLabel(e.status);
          var dirCls = e.direction === "I" ? "inbound" : "outbound";
          var dirText = e.direction === "I" ? "IN" : "OUT";

          html += "<tr>";
          html += "<td>" + e.id + "</td>";
          html += "<td>" + formatTime(e.received_at) + "</td>";
          html +=
            '<td><span class="dir-badge ' + dirCls + '">' + dirText + "</span></td>";
          html += "<td>" + escapeHtml(e.patient_name || "--") + "</td>";
          html +=
            '<td><span class="status-pill ' + st.cls + '">' + st.text + "</span></td>";
          html += "<td>" + (e.error_count || 0) + "</td>";
          html += "</tr>";
        }

        tbody.innerHTML = html;
      })
      .catch(function (err) {
        console.error("Events fetch error:", err);
      });
  }

  // --- Bot Slots ---

  function refreshSlots() {
    fetch("/api/slots")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        var slots = data.slots || [];
        for (var i = 0; i < slots.length; i++) {
          var slot = slots[i];
          var el = document.getElementById("slot-" + slot.slot_name);
          if (!el) continue;

          var statusEl = el.querySelector(".slot-status");
          var metaEl = el.querySelector(".slot-meta");

          if (slot.status === "active") {
            el.className = "slot-card slot-active";
            statusEl.textContent = "Active \u2014 PID " + (slot.session_id || "?");

            // Show heartbeat age
            if (slot.heartbeat_at) {
              var hbAge = Math.round((Date.now() - new Date(slot.heartbeat_at).getTime()) / 1000);
              metaEl.textContent = "last seen " + hbAge + "s ago";
            } else {
              metaEl.textContent = "";
            }
          } else {
            el.className = "slot-card slot-open";
            statusEl.textContent = "Available";
            metaEl.textContent = "";
          }
        }
      })
      .catch(function (err) {
        console.error("Slots fetch error:", err);
      });
  }

  // --- Start-All Status ---

  function refreshStartAllStatus() {
    fetch("/api/start-all-status")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        var panel = document.getElementById("start-all-panel");
        var container = document.getElementById("start-all-bots");
        var btn = document.getElementById("btn-start-all");
        var bots = data.bots || {};
        var botNames = Object.keys(bots);

        if (botNames.length === 0) {
          panel.style.display = "none";
          return;
        }

        panel.style.display = "";
        var html = "";

        for (var i = 0; i < botNames.length; i++) {
          var name = botNames[i];
          var b = bots[name];
          var statusCls = "sa-" + (b.status || "pending");

          html += '<div class="sa-bot-card ' + statusCls + '">';
          html += '  <div class="sa-bot-name">' + escapeHtml(name) + '</div>';
          html += '  <div class="sa-bot-attempt">Attempt ' + b.attempt + '/' + b.max_attempts + '</div>';
          html += '  <div class="sa-bot-step">' + escapeHtml(b.step || "") + '</div>';
          html += '  <div class="sa-bot-status">' + escapeHtml(b.status || "pending") + '</div>';

          if (b.error) {
            html += '  <div class="sa-bot-error">' + escapeHtml(b.error) + '</div>';
          }
          html += '</div>';
        }

        container.innerHTML = html;

        // Update button state based on running status
        if (!data.running && botNames.length > 0) {
          // Finished — show summary
          var succeeded = 0;
          var failed = 0;
          for (var j = 0; j < botNames.length; j++) {
            var st = bots[botNames[j]].status;
            if (st === "running" || st === "fallback") succeeded++;
            if (st === "failed") failed++;
          }

          btn.textContent = succeeded + " running, " + failed + " failed";
          setTimeout(function () {
            btn.textContent = "Start All";
            btn.disabled = false;
          }, 5000);

          // Stop polling
          if (_startAllPollTimer) {
            clearInterval(_startAllPollTimer);
            _startAllPollTimer = null;
          }

          // Refresh errors now
          refreshBotErrors();
        }
      })
      .catch(function (err) {
        console.error("Start-all status error:", err);
      });
  }

  // --- Bot Errors ---

  function refreshBotErrors() {
    fetch("/api/bot-errors?limit=30")
      .then(function (r) { return r.json(); })
      .then(function (data) {
        var panel = document.getElementById("bot-errors-panel");
        var container = document.getElementById("bot-errors-list");
        var badge = document.getElementById("error-count-badge");
        var errors = data.errors || [];

        if (errors.length === 0) {
          panel.style.display = "none";
          return;
        }

        panel.style.display = "";
        badge.textContent = errors.length;

        var html = "";
        for (var i = 0; i < errors.length; i++) {
          var e = errors[i];
          var attemptStr = e.attempt + "/" + e.max_attempts;
          var isLast = e.attempt >= e.max_attempts;

          html += '<div class="error-card' + (isLast ? " error-final" : "") + '">';
          html += '  <div class="error-card-header">';
          html += '    <span class="error-slot">' + escapeHtml(e.slot_name || "").toUpperCase() + '</span>';
          html += '    <span class="error-attempt">Attempt ' + attemptStr + '</span>';
          html += '    <span class="error-step">' + escapeHtml(e.step || "") + '</span>';
          html += '    <span class="error-time">' + formatTime(e.created_at) + '</span>';
          html += '  </div>';
          html += '  <div class="error-card-body">' + escapeHtml(e.error || "") + '</div>';

          if (e.traceback) {
            html += '  <details class="error-traceback"><summary>Traceback</summary>';
            html += '    <pre>' + escapeHtml(e.traceback) + '</pre>';
            html += '  </details>';
          }
          html += '</div>';
        }

        container.innerHTML = html;
      })
      .catch(function (err) {
        console.error("Bot errors fetch error:", err);
      });
  }

  // --- Clock ---

  function updateClock() {
    setText("clock", new Date().toLocaleTimeString());
  }

  // --- Init ---

  refreshSessions();
  refreshSlots();
  refreshStats();
  refreshEvents();
  refreshBotErrors();
  updateClock();

  setInterval(refreshSessions, SESSION_INTERVAL);
  setInterval(refreshSlots, SESSION_INTERVAL);
  setInterval(refreshStats, STATS_INTERVAL);
  setInterval(refreshEvents, STATS_INTERVAL);
  setInterval(refreshBotErrors, STATS_INTERVAL);
  setInterval(updateClock, 1000);
})();
