/**
 * MHT Agentic Dashboard — Stripe-inspired SaaS admin panel.
 * Tabs: Monitor, Configure, Health, Events, Errors
 */

(function () {
  "use strict";

  var SESSION_INTERVAL = 5000;
  var STATS_INTERVAL = 10000;
  var HEALTH_INTERVAL = 8000;

  var _currentSessionIds = [];
  var _noSessionsShown = false;

  // ──────────────────────────────────────────────────────────
  //  Tab switching
  // ──────────────────────────────────────────────────────────

  window.switchTab = function (tabName) {
    var panels = document.querySelectorAll(".tab-panel");
    for (var i = 0; i < panels.length; i++) panels[i].classList.remove("active");

    var navItems = document.querySelectorAll(".nav-item");
    for (var j = 0; j < navItems.length; j++) navItems[j].classList.remove("active");

    var panel = document.getElementById("tab-" + tabName);
    if (panel) panel.classList.add("active");

    var navItem = document.querySelector('.nav-item[data-tab="' + tabName + '"]');
    if (navItem) navItem.classList.add("active");

    if (tabName === "configure") {
      refreshLocations();
      loadConfig();
    } else if (tabName === "health") {
      refreshHealth();
    } else if (tabName === "events") {
      refreshEvents();
    } else if (tabName === "errors") {
      refreshBotErrors();
    }
  };

  // ──────────────────────────────────────────────────────────
  //  Helpers
  // ──────────────────────────────────────────────────────────

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
    if (code >= 100) return { text: "Complete" };
    if (code >= 40) return { text: "Sent" };
    if (code >= 10) return { text: "Converted" };
    if (code === 0) return { text: "Pending" };
    if (code < 0) return { text: "Failed" };
    return { text: String(code) };
  }

  function escapeHtml(str) {
    var div = document.createElement("div");
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
  }

  function postJSON(url, body) {
    var opts = {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    };
    if (body !== undefined) opts.body = JSON.stringify(body);
    return fetch(url, opts).then(function (r) {
      return r.json();
    });
  }

  // ──────────────────────────────────────────────────────────
  //  Monitor: RDP Control
  // ──────────────────────────────────────────────────────────

  window.startAllRDP = function () {
    var btn = document.getElementById("btn-start-all");
    btn.disabled = true;
    btn.textContent = "Launching...";

    postJSON("/api/rdp/start-all")
      .then(function (data) {
        if (data.error) {
          btn.textContent = data.error;
        } else {
          btn.textContent = "Running (pid " + data.pid + ")";
        }
        setTimeout(function () {
          btn.textContent = "Start All";
          btn.disabled = false;
        }, 5000);
      })
      .catch(function () {
        btn.textContent = "Start All";
        btn.disabled = false;
      });
  };

  window.stopAllRDP = function () {
    var btn = document.getElementById("btn-stop-all");
    btn.disabled = true;
    btn.textContent = "Stopping...";

    postJSON("/api/rdp/stop-all")
      .then(function (data) {
        var killed = data.killed || 0;
        btn.textContent = "Done — " + killed + " killed";
        setTimeout(function () {
          btn.textContent = "Stop All";
          btn.disabled = false;
        }, 3000);
        refreshSlots();
      })
      .catch(function () {
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
          var names = sessions.map(function (s) {
            return s.username + " (" + s.mode + ")";
          });
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

  window.rebootSession = function (sessionId, sessionTitle) {
    // Extract slot name from title (e.g. "ExperityB_MHT - localhost..." -> "experityb")
    var match = sessionTitle.match(/Experity(\w)/i);
    var slotName = match ? "experity" + match[1].toLowerCase() : "";
    var label = slotName.toUpperCase() || "Session " + sessionId;

    if (!confirm("Reboot " + label + "? This will kill the bot, log off the session, and relaunch."))
      return;

    // Use the session's WTS session ID if it's small, otherwise use slot-based reboot
    var endpoint = slotName
      ? "/api/rdp/reboot/" + slotName
      : "/api/rdp/stop/" + sessionId;

    postJSON(endpoint)
      .then(function (data) {
        if (data.success) {
          alert(data.message || "Rebooting " + label + "...");
        } else {
          alert("Failed: " + (data.error || "Unknown error"));
        }
      })
      .catch(function (err) {
        console.error("Reboot error:", err);
      });
  };

  window.startRDP = function (path) {
    fetch("/api/rdp/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path: path }),
    });
  };

  // ──────────────────────────────────────────────────────────
  //  Monitor: Sessions
  // ──────────────────────────────────────────────────────────

  function refreshSessions() {
    fetch("/api/sessions")
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        var container = $("#sessions-container");
        var health = data.bot_health || {};

        var badge = $("#bot-status");
        var statusText = (health.status || "unknown").replace("_", " ");
        if (badge.textContent !== statusText) {
          badge.textContent = statusText;
          badge.className = "status-indicator " + (health.status || "unknown");
        }

        var sessions = data.sessions || [];
        var hasReal = sessions.some(function (s) {
          return s.id !== "none";
        });

        if (!hasReal) {
          if (!_noSessionsShown) {
            _noSessionsShown = true;
            _currentSessionIds = [];
            var html = '<div class="session-card placeholder">';
            html += "<p>No RDP sessions detected — click Start All</p>";
            html += "</div>";
            container.innerHTML = html;
            if (false) {
            fetch("/api/rdp/files")
              .then(function (r) {
                return r.json();
              })
              .then(function (fileData) {
                var files = fileData.files || [];
                var html = '<div class="session-card placeholder">';
                html += "<p>No RDP sessions detected</p>";
                if (files.length > 0) {
                  html += '<div class="rdp-file-list">';
                  for (var i = 0; i < files.length; i++) {
                    html +=
                      '<button class="btn btn-primary btn-sm" onclick="startRDP(\'' +
                      escapeHtml(
                        files[i].path.replace(/\\/g, "\\\\").replace(/'/g, "\\'")
                      ) +
                      "')\">" +
                      escapeHtml(files[i].name) +
                      "</button> ";
                  }
                  html += "</div>";
                }
                html += "</div>";
                container.innerHTML = html;
              });
          } // end if(false)
          }
          return;
        }

        _noSessionsShown = false;

        var newIds = sessions
          .filter(function (s) {
            return s.id !== "none";
          })
          .map(function (s) {
            return s.id;
          });

        var structureChanged =
          newIds.length !== _currentSessionIds.length ||
          newIds.some(function (id, i) {
            return id !== _currentSessionIds[i];
          });

        if (structureChanged) {
          var html = "";
          for (var i = 0; i < sessions.length; i++) {
            var s = sessions[i];
            if (s.id === "none") continue;

            html += '<div class="session-card" data-session="' + s.id + '">';
            html += '  <div class="session-header">';
            html +=
              '    <span class="session-title">' + escapeHtml(s.title);
            if (s.agent_role === "inbound") {
              html += ' <span class="role-badge">INBOUND</span>';
            } else if (s.agent_role === "outbound") {
              html += ' <span class="role-badge">OUTBOUND</span>';
            }
            html += "</span>";
            html += '    <div class="session-header-right">';
            html +=
              '      <span class="session-meta">' +
              s.width +
              "x" +
              s.height +
              "</span>";
            html +=
              '      <button class="btn btn-reboot btn-sm" onclick="rebootSession(\'' +
              s.id + "', '" + escapeHtml(s.title).replace(/'/g, "\\'") +
              "')\">Reboot</button>";
            html +=
              '      <button class="btn btn-secondary btn-sm" onclick="stopSession(\'' +
              s.id +
              "')\">Close</button>";
            html += "    </div>";
            html += "  </div>";
            html += '  <div class="session-screenshot">';
            if (s.has_screenshot) {
              html +=
                '    <img src="/api/screenshots/' +
                s.id +
                "?t=" +
                Date.now() +
                '" alt="RDP Screenshot" loading="lazy">';
            } else {
              html +=
                '    <span class="no-screenshot">No screenshot available</span>';
            }
            html += "  </div>";
            html += "</div>";
          }
          container.innerHTML = html;
          _currentSessionIds = newIds;
        } else {
          for (var j = 0; j < sessions.length; j++) {
            var s2 = sessions[j];
            if (s2.id === "none") continue;

            var card = container.querySelector(
              '[data-session="' + s2.id + '"]'
            );
            if (!card) continue;

            if (s2.has_screenshot) {
              var img = card.querySelector("img");
              var newSrc =
                "/api/screenshots/" + s2.id + "?t=" + Date.now();
              if (img) {
                img.src = newSrc;
              } else {
                var wrap = card.querySelector(".session-screenshot");
                wrap.innerHTML =
                  '<img src="' +
                  newSrc +
                  '" alt="RDP Screenshot" loading="lazy">';
              }
            }
          }
        }
      })
      .catch(function (err) {
        console.error("Sessions fetch error:", err);
      });
  }

  // ──────────────────────────────────────────────────────────
  //  Monitor: Stats
  // ──────────────────────────────────────────────────────────

  function refreshStats() {
    fetch("/api/analytics/current")
      .then(function (r) {
        return r.json();
      })
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

  // ──────────────────────────────────────────────────────────
  //  Monitor: Slots
  // ──────────────────────────────────────────────────────────

  function refreshSlots() {
    fetch("/api/slots")
      .then(function (r) {
        return r.json();
      })
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
            statusEl.textContent = "Active - PID " + (slot.session_id || "?");

            if (slot.heartbeat_at) {
              var hbAge = Math.round(
                (Date.now() - new Date(slot.heartbeat_at).getTime()) / 1000
              );
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

  // ──────────────────────────────────────────────────────────
  //  Events Tab
  // ──────────────────────────────────────────────────────────

  function refreshEvents() {
    fetch("/api/analytics/events")
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        var tbody = $("#events-body");
        var events = data.events || [];

        if (events.length === 0) {
          tbody.innerHTML =
            '<tr><td colspan="12" class="empty">No events yet</td></tr>';
          return;
        }

        var html = "";
        for (var i = 0; i < events.length; i++) {
          var e = events[i];
          var st = statusLabel(e.status);
          var dirText = e.direction === "I" ? "IN" : "OUT";

          html += "<tr>";
          html += "<td>" + e.id + "</td>";
          html += "<td>" + formatTime(e.received_at) + "</td>";
          html +=
            '<td><span class="dir-badge">' + dirText + "</span></td>";
          html +=
            "<td><strong>" + escapeHtml(e.patient_name || "--") + "</strong></td>";
          html += "<td>" + escapeHtml(e.dob || "--") + "</td>";
          html += "<td>" + escapeHtml(e.cell_phone || e.home_phone || "--") + "</td>";
          html += "<td>" + escapeHtml(e.email || "--") + "</td>";
          html += "<td>" + escapeHtml(e.address || "--") + "</td>";
          html += "<td>" + escapeHtml(e.zip || "--") + "</td>";
          html += "<td>" + escapeHtml(e.location || "--") + "</td>";
          html +=
            '<td><span class="status-pill">' +
            st.text +
            "</span></td>";
          html += "<td>" + (e.error_count || 0) + "</td>";
          html += "</tr>";
        }

        tbody.innerHTML = html;
      })
      .catch(function (err) {
        console.error("Events fetch error:", err);
      });
  }

  // ──────────────────────────────────────────────────────────
  //  Errors Tab
  // ──────────────────────────────────────────────────────────

  function refreshBotErrors() {
    fetch("/api/bot-errors?limit=30")
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        var container = document.getElementById("bot-errors-list");
        var badge = document.getElementById("error-count-badge");
        var navBadge = document.getElementById("nav-error-badge");
        var errors = data.errors || [];

        badge.textContent = errors.length;

        if (errors.length === 0) {
          container.innerHTML =
            '<div class="loading-msg">No errors</div>';
          if (navBadge) navBadge.style.display = "none";
          return;
        }

        if (navBadge) {
          navBadge.style.display = "";
          navBadge.textContent = errors.length;
        }

        var html = "";
        for (var i = 0; i < errors.length; i++) {
          var e = errors[i];
          var attemptStr = e.attempt + "/" + e.max_attempts;
          var isLast = e.attempt >= e.max_attempts;

          html +=
            '<div class="error-card' +
            (isLast ? " error-final" : "") +
            '">';
          html += '  <div class="error-card-header">';
          html +=
            '    <span class="error-slot">' +
            escapeHtml(e.slot_name || "").toUpperCase() +
            "</span>";
          html +=
            '    <span class="error-attempt">Attempt ' +
            attemptStr +
            "</span>";
          html +=
            '    <span class="error-step">' +
            escapeHtml(e.step || "") +
            "</span>";
          html +=
            '    <span class="error-time">' +
            formatTime(e.created_at) +
            "</span>";
          html += "  </div>";
          html +=
            '  <div class="error-card-body">' +
            escapeHtml(e.error || "") +
            "</div>";

          if (e.traceback) {
            html +=
              '  <details class="error-traceback"><summary>Traceback</summary>';
            html +=
              "    <pre>" + escapeHtml(e.traceback) + "</pre>";
            html += "  </details>";
          }
          html += "</div>";
        }

        container.innerHTML = html;
      })
      .catch(function (err) {
        console.error("Bot errors fetch error:", err);
      });
  }

  // ──────────────────────────────────────────────────────────
  //  Configure Tab: Locations
  // ──────────────────────────────────────────────────────────

  function refreshLocations() {
    fetch("/api/locations")
      .then(function (r) {
        return r.json();
      })
      .then(function (data) {
        renderLocations(data.locations || []);
      })
      .catch(function (err) {
        console.error("Locations fetch error:", err);
      });
  }

  function renderLocations(locations) {
    var container = document.getElementById("locations-list");

    if (locations.length === 0) {
      container.innerHTML =
        '<div class="loading-msg">No locations configured</div>';
      return;
    }

    var html = "";
    for (var i = 0; i < locations.length; i++) {
      var loc = locations[i];
      var name = loc.name || loc.location_name || "";
      var isActive = loc.is_active === 1 || loc.is_active === true;
      var statusCls = isActive ? "active" : "";
      var statusText = isActive ? "Active" : "Off";
      var checked = isActive ? "checked" : "";

      html += '<div class="location-row">';
      html += '  <div class="location-left">';
      html += '    <span class="location-name">' + escapeHtml(name) + "</span>";
      html += '    <span class="location-status ' + statusCls + '">' + statusText + "</span>";
      html += "  </div>";
      html += '  <label class="toggle-switch">';
      html += '    <input type="checkbox" ' + checked + " onchange=\"toggleLocation('" + escapeHtml(name) + "', this.checked)\">";
      html += '    <span class="toggle-slider"></span>';
      html += "  </label>";
      html += "</div>";
    }

    container.innerHTML = html;
  }

  window.toggleLocation = function (name, active) {
    postJSON("/api/locations/" + encodeURIComponent(name) + "/toggle", {
      active: active,
    })
      .then(function (data) {
        if (data.locations) renderLocations(data.locations);
      })
      .catch(function (err) {
        console.error("Toggle location error:", err);
        refreshLocations();
      });
  };

  // ──────────────────────────────────────────────────────────
  //  Configure Tab: Bot Config
  // ──────────────────────────────────────────────────────────

  function loadConfig() {
    fetch("/api/config")
      .then(function (r) {
        return r.json();
      })
      .then(function (config) {
        if (config.rdp_count !== undefined) {
          document.getElementById("cfg-rdp-count").value = config.rdp_count;
        }
        if (config.inbound_count !== undefined) {
          document.getElementById("cfg-inbound-count").value =
            config.inbound_count;
        }
        if (config.outbound_count !== undefined) {
          document.getElementById("cfg-outbound-count").value =
            config.outbound_count;
        }
      })
      .catch(function (err) {
        console.error("Config fetch error:", err);
      });
  }

  window.saveConfig = function () {
    var payload = {
      rdp_count: document.getElementById("cfg-rdp-count").value,
      inbound_count: document.getElementById("cfg-inbound-count").value,
      outbound_count: document.getElementById("cfg-outbound-count").value,
    };

    postJSON("/api/config", payload)
      .then(function (data) {
        if (data.success) {
          var btn = document.querySelector(".config-actions .btn-primary");
          if (btn) {
            btn.textContent = "Saved";
            setTimeout(function () {
              btn.textContent = "Save Configuration";
            }, 2000);
          }
        }
      })
      .catch(function (err) {
        console.error("Config save error:", err);
      });
  };

  window.loadConfig = loadConfig;

  // ──────────────────────────────────────────────────────────
  //  Health Tab
  // ──────────────────────────────────────────────────────────

  function refreshHealth() {
    Promise.all([
      fetch("/api/slots").then(function (r) {
        return r.json();
      }),
      fetch("/api/sessions").then(function (r) {
        return r.json();
      }),
    ])
      .then(function (results) {
        var slotsData = results[0];
        var sessionsData = results[1];
        renderHealthCards(slotsData.slots || [], sessionsData);
      })
      .catch(function (err) {
        console.error("Health fetch error:", err);
      });
  }

  function renderHealthCards(slots, sessionsData) {
    var container = document.getElementById("health-cards");
    var overallBadge = document.getElementById("health-overall");

    var botDefs = [
      { slot: "experityb", label: "ExperityB", role: "inbound" },
      { slot: "experityd", label: "ExperityD", role: "inbound" },
      { slot: "experityc", label: "ExperityC", role: "outbound" },
    ];

    var activeCount = 0;
    var html = "";

    for (var i = 0; i < botDefs.length; i++) {
      var def = botDefs[i];
      var slotData = null;

      for (var j = 0; j < slots.length; j++) {
        if (slots[j].slot_name === def.slot) {
          slotData = slots[j];
          break;
        }
      }

      var isActive = slotData && slotData.status === "active";
      var isStale = false;
      var hbAge = "--";

      if (isActive && slotData.heartbeat_at) {
        var age = Math.round(
          (Date.now() - new Date(slotData.heartbeat_at).getTime()) / 1000
        );
        hbAge = age + "s ago";
        if (age > 60) isStale = true;
        activeCount++;
      } else if (isActive) {
        activeCount++;
      }

      var cardClass = "health-card";
      var dotClass = "health-status-dot";
      var statusText = "Open";

      if (isActive && !isStale) {
        cardClass += " health-active";
        dotClass += " dot-active";
        statusText = "Active";
      } else if (isActive && isStale) {
        cardClass += " health-stale";
        dotClass += " dot-stale";
        statusText = "Stale";
      } else {
        cardClass += " health-open";
        dotClass += " dot-open";
        statusText = "Open";
      }

      var sessionId = slotData ? slotData.session_id || "--" : "--";
      var location = slotData ? slotData.location || "--" : "--";

      html += '<div class="' + cardClass + '">';
      html += '  <div class="health-card-header">';
      html +=
        '    <span class="health-bot-name">' +
        escapeHtml(def.label) +
        "</span>";
      html +=
        '    <span class="' + dotClass + '">' + statusText + "</span>";
      html += "  </div>";

      html += '  <div class="health-detail">';
      html += '    <span class="health-detail-label">Role</span>';
      html +=
        '    <span class="health-detail-value">' +
        def.role.toUpperCase() +
        "</span>";
      html += "  </div>";

      html += '  <div class="health-detail">';
      html += '    <span class="health-detail-label">Heartbeat</span>';
      html +=
        '    <span class="health-detail-value">' + hbAge + "</span>";
      html += "  </div>";

      html += '  <div class="health-detail">';
      html +=
        '    <span class="health-detail-label">Session ID</span>';
      html +=
        '    <span class="health-detail-value">' +
        escapeHtml(String(sessionId)) +
        "</span>";
      html += "  </div>";

      html += '  <div class="health-detail">';
      html += '    <span class="health-detail-label">Location</span>';
      html +=
        '    <span class="health-detail-value">' +
        escapeHtml(String(location)) +
        "</span>";
      html += "  </div>";

      html += '  <div class="health-card-actions">';
      html +=
        '    <button class="btn btn-secondary btn-sm" onclick="rebootBot(\'' +
        def.slot +
        "')\">Reboot</button>";
      html += "  </div>";
      html += "</div>";
    }

    container.innerHTML = html;

    if (activeCount === botDefs.length) {
      overallBadge.textContent = "All Bots Active";
      overallBadge.className = "health-overall-badge all-active";
    } else if (activeCount > 0) {
      overallBadge.textContent =
        activeCount + "/" + botDefs.length + " Active";
      overallBadge.className = "health-overall-badge partial";
    } else {
      overallBadge.textContent = "No Bots Active";
      overallBadge.className = "health-overall-badge";
    }
  }

  window.rebootBot = function (slotName) {
    if (
      !confirm(
        "Reboot " +
          slotName.toUpperCase() +
          "? This will kill its processes and relaunch."
      )
    )
      return;

    postJSON("/api/rdp/reboot/" + slotName)
      .then(function (data) {
        if (data.success) {
          alert(data.message || "Rebooting...");
        } else {
          alert("Failed: " + (data.error || "Unknown error"));
        }
        setTimeout(refreshHealth, 3000);
      })
      .catch(function (err) {
        console.error("Reboot error:", err);
      });
  };

  // ──────────────────────────────────────────────────────────
  //  Clock
  // ──────────────────────────────────────────────────────────

  function updateClock() {
    setText("clock", new Date().toLocaleTimeString());
  }

  // ──────────────────────────────────────────────────────────
  //  Init + Polling
  // ──────────────────────────────────────────────────────────

  refreshSessions();
  refreshSlots();
  refreshStats();
  refreshEvents();
  refreshBotErrors();
  refreshHealth();
  loadConfig();
  refreshLocations();
  updateClock();

  setInterval(refreshSessions, SESSION_INTERVAL);
  setInterval(refreshSlots, SESSION_INTERVAL);
  setInterval(refreshStats, STATS_INTERVAL);
  setInterval(refreshEvents, STATS_INTERVAL);
  setInterval(refreshBotErrors, STATS_INTERVAL);
  setInterval(refreshHealth, HEALTH_INTERVAL);
  setInterval(updateClock, 1000);
})();
