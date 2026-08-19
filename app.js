/* ---------------------------------------------------------------------------
 * app.js — entry logging, pay-period maths, and pushing rows into the workbook.
 *
 * Entries live in localStorage first and sync to Excel second, so logging a day
 * never depends on the network or on being signed in. Anything unsynced is
 * retried on the next load, on the next save, and whenever "Sync now" is hit.
 * ------------------------------------------------------------------------- */
(function () {
  "use strict";

  var STORE = "wh.entries";
  var CLOCK = "wh.clock";
  var RATE_KEY = "wh.payRate";
  var DEC = (window.CONFIG && window.CONFIG.decimals) || 2;
  var DAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  var MONTHS = ["January", "February", "March", "April", "May", "June", "July",
    "August", "September", "October", "November", "December"];

  var $ = function (id) { return document.getElementById(id); };

  // Carry over anything saved while the app was still called "internship hours".
  // Runs once and is a no-op on a fresh device.
  (function migrate() {
    ["entries", "clock", "settings", "token", "book"].forEach(function (k) {
      var from = "ih." + k, to = "wh." + k;
      if (localStorage.getItem(from) !== null && localStorage.getItem(to) === null) {
        localStorage.setItem(to, localStorage.getItem(from));
      }
      localStorage.removeItem(from);
    });
  })();

  /* -- dates and hours ---------------------------------------------------- */

  function todayISO() {
    var d = new Date();
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  }

  function pad(n) { return (n < 10 ? "0" : "") + n; }

  function parseISO(iso) {
    var p = iso.split("-");
    return { y: +p[0], m: +p[1], d: +p[2] };
  }

  // Excel's day zero is 1899-12-30. Built from UTC so a timezone shift can
  // never slide an entry onto the day before.
  function dateSerial(iso) {
    var p = parseISO(iso);
    return Math.round((Date.UTC(p.y, p.m - 1, p.d) - Date.UTC(1899, 11, 30)) / 86400000);
  }

  function minutes(hhmm) {
    var p = hhmm.split(":");
    return (+p[0]) * 60 + (+p[1]);
  }

  function timeFraction(hhmm) { return minutes(hhmm) / 1440; }

  // Inverse of dateSerial/timeFraction, for reading rows back out of the
  // workbook. Rounded rather than trusted exactly, because Excel round-trips
  // these as floats and can hand back e.g. 46252.0000000002.
  function serialToISO(serial) {
    var ms = Math.round(serial) * 86400000 + Date.UTC(1899, 11, 30);
    var d = new Date(ms);
    return d.getUTCFullYear() + "-" + pad(d.getUTCMonth() + 1) + "-" + pad(d.getUTCDate());
  }

  function fracToHHMM(frac) {
    var total = ((Math.round(frac * 1440) % 1440) + 1440) % 1440;
    return pad(Math.floor(total / 60)) + ":" + pad(total % 60);
  }

  // Loose equality for comparing a local entry's serial/fraction against a
  // value that has round-tripped through Excel, which can drift in the last
  // float bit even for the same nominal time.
  function nearlyEqual(a, b) { return Math.abs(Number(a) - Number(b)) < 1e-6; }

  function matchesRow(e, row) {
    return nearlyEqual(dateSerial(e.date), row[0]) &&
      nearlyEqual(timeFraction(e.tin), row[2]) &&
      nearlyEqual(timeFraction(e.tout), row[3]);
  }

  function decimalHours(tin, tout) {
    return round((minutes(tout) - minutes(tin)) / 60);
  }

  function round(n) {
    var f = Math.pow(10, DEC);
    return Math.round(n * f) / f;
  }

  function fmt(n) { return Number(n).toFixed(DEC); }

  // Kept separate from DEC/fmt on purpose: hours-decimal precision is a
  // config knob, but a dollar figure is always two decimals regardless.
  function fmtMoney(n) { return Number(n).toFixed(2); }

  // Local-only estimate of what an hour is worth, never part of any entry,
  // never sent to Excel or the CSV backup. Defaults to $40 the first time the
  // app runs; after that, whatever the Pay rate field was last set to.
  function getPayRate() {
    var v = parseFloat(localStorage.getItem(RATE_KEY));
    return isFinite(v) && v >= 0 ? v : 40;
  }

  function setPayRate(v) {
    var n = parseFloat(v);
    if (!isFinite(n) || n < 0) return;
    localStorage.setItem(RATE_KEY, String(n));
  }

  function dayName(iso) {
    var p = parseISO(iso);
    return DAYS[new Date(p.y, p.m - 1, p.d).getDay()];
  }

  function time12(hhmm) {
    var p = hhmm.split(":"), h = +p[0], suffix = h < 12 ? "AM" : "PM";
    var hour = h % 12; if (hour === 0) hour = 12;
    return hour + ":" + p[1] + " " + suffix;
  }

  function lastDay(y, m) { return new Date(y, m, 0).getDate(); }

  /* -- pay periods: 1st-15th and 16th-end of month ------------------------ */

  function periodOf(iso) {
    var p = parseISO(iso);
    return p.d <= 15 ? 1 : 2;
  }

  // Excel auto-detects a cell value that looks like a date the same way it
  // would if typed by hand, regardless of how it arrives through the Graph
  // API. "1-15" was silently read as January 15 and stored as a real date,
  // which then set a date display format on the whole column. Letters in the
  // label make that misread impossible.
  function periodLabel(iso) {
    return periodOf(iso) === 1 ? "1st - 15th" : "16th - End";
  }

  function periodBounds(y, m, which) {
    if (which === 1) {
      return { start: y + "-" + pad(m) + "-01", end: y + "-" + pad(m) + "-15" };
    }
    return { start: y + "-" + pad(m) + "-16",
             end: y + "-" + pad(m) + "-" + pad(lastDay(y, m)) };
  }

  function currentPeriod() {
    var d = new Date();
    return periodBounds(d.getFullYear(), d.getMonth() + 1,
      d.getDate() <= 15 ? 1 : 2);
  }

  function prettyRange(range) {
    var a = parseISO(range.start), b = parseISO(range.end);
    return MONTHS[a.m - 1].slice(0, 3) + " " + a.d + " to " +
      MONTHS[b.m - 1].slice(0, 3) + " " + b.d + ", " + b.y;
  }

  /* -- store -------------------------------------------------------------- */

  var entries = load();

  function load() {
    try { return JSON.parse(localStorage.getItem(STORE) || "[]"); }
    catch (e) { return []; }
  }

  function save() {
    localStorage.setItem(STORE, JSON.stringify(entries));
  }

  function sorted() {
    return entries.slice().sort(function (a, b) {
      if (a.date === b.date) return a.tin < b.tin ? 1 : -1;
      return a.date < b.date ? 1 : -1;
    });
  }

  function inRange(e, range) {
    return e.date >= range.start && e.date <= range.end;
  }

  function sumHours(list) {
    return round(list.reduce(function (t, e) {
      return t + decimalHours(e.tin, e.tout);
    }, 0));
  }

  function rowValues(e) {
    return [
      dateSerial(e.date),
      dayName(e.date),
      timeFraction(e.tin),
      timeFraction(e.tout),
      decimalHours(e.tin, e.tout),
      periodLabel(e.date),
      e.notes || ""
    ];
  }

  /* -- rendering ---------------------------------------------------------- */

  function renderPeriod() {
    var range = currentPeriod();
    var list = entries.filter(function (e) { return inRange(e, range); });
    var total = sumHours(list);
    var days = {};
    list.forEach(function (e) { days[e.date] = true; });
    var dayCount = Object.keys(days).length;

    $("periodRange").textContent = prettyRange(range);
    $("periodHours").textContent = fmt(total);
    $("periodPay").textContent = fmtMoney(total * getPayRate());
    $("periodDays").textContent = dayCount;
    $("periodAvg").textContent = fmt(dayCount ? total / dayCount : 0);
  }

  function renderMonth() {
    var y = +$("mYear").value, m = +$("mMonth").value;
    var first = periodBounds(y, m, 1), second = periodBounds(y, m, 2);
    var a = sumHours(entries.filter(function (e) { return inRange(e, first); }));
    var b = sumHours(entries.filter(function (e) { return inRange(e, second); }));

    $("mFirst").textContent = fmt(a);
    $("mSecond").textContent = fmt(b);
    $("mTotal").textContent = fmt(round(a + b));
    $("mPay").textContent = fmtMoney(round(a + b) * getPayRate());
    renderEntries(y, m);
  }

  function renderEntries(y, m) {
    var prefix = y + "-" + pad(m);
    var list = sorted().filter(function (e) { return e.date.indexOf(prefix) === 0; });
    var box = $("entryList");

    $("entryScope").textContent = MONTHS[m - 1] + " " + y;

    if (!list.length) {
      box.innerHTML = '<p class="hint">No entries this month.</p>';
      return;
    }

    var rate = getPayRate();
    box.innerHTML = list.map(function (e) {
      return '<div class="entry' + (e.synced ? "" : " unsynced") + '">' +
        '<div class="entry-main">' +
          '<div class="entry-date">' + dayName(e.date) + " " +
            e.date.slice(5).replace("-", "/") + '</div>' +
          '<div class="entry-time">' + time12(e.tin) + " to " + time12(e.tout) +
            (e.notes ? ' <span class="entry-notes">' + escapeHtml(e.notes) + '</span>' : '') +
          '</div>' +
        '</div>' +
        '<div class="entry-hours">' + fmt(decimalHours(e.tin, e.tout)) +
          (e.synced ? '' : '<span class="dot" title="Not yet in Excel"></span>') +
          '<div class="entry-pay">$' + fmtMoney(decimalHours(e.tin, e.tout) * rate) + '</div>' +
        '</div>' +
        '<button class="del" type="button" data-id="' + e.id + '" aria-label="Delete entry">&times;</button>' +
      '</div>';
    }).join("");
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  function renderAll() {
    renderPeriod();
    renderMonth();
    renderSyncState();
  }

  /* -- Excel sync --------------------------------------------------------- */

  var syncing = false;

  function pending() {
    return entries.filter(function (e) { return !e.synced; });
  }

  // Reconciles local storage against a full read of the workbook, so logging
  // from more than one device converges instead of each device only ever
  // showing what it personally typed in.
  //
  // A row present remotely but not locally was added elsewhere and is pulled
  // in as already synced. A LOCAL entry marked synced but absent from the
  // fetch was deleted elsewhere and is removed here too. Never touches an
  // unsynced entry: that is work this device hasn't pushed yet, and a row
  // missing from the workbook because it hasn't been sent yet is not the same
  // thing as a row that was deleted after being sent.
  function mergeFromWorkbook(rows) {
    var added = 0, removed = 0;

    rows.forEach(function (row) {
      var exists = entries.some(function (e) { return matchesRow(e, row); });
      if (exists) return;
      entries.push({
        id: String(Date.now()) + Math.random().toString(36).slice(2, 7),
        date: serialToISO(row[0]),
        tin: fracToHHMM(row[2]),
        tout: fracToHHMM(row[3]),
        notes: row[6] || "",
        synced: true
      });
      added++;
    });

    entries = entries.filter(function (e) {
      if (!e.synced) return true;
      var stillThere = rows.some(function (row) { return matchesRow(e, row); });
      if (stillThere) return true;
      removed++;
      return false;
    });

    if (added || removed) save();
    return { added: added, removed: removed };
  }

  function renderSyncState() {
    var badge = $("syncBadge");
    var n = pending().length;

    if (!window.Excel.configured()) {
      badge.textContent = "Setup";
      badge.className = "badge warn";
      return;
    }
    if (!window.Excel.isSignedIn()) {
      badge.textContent = n ? n + " local" : "Not connected";
      badge.className = "badge";
      return;
    }
    if (syncing) {
      badge.textContent = "Syncing";
      badge.className = "badge busy";
      return;
    }
    badge.textContent = n ? n + " pending" : "Synced";
    badge.className = "badge " + (n ? "busy" : "ok");
  }

  function setExcelState(msg, kind) {
    var el = $("excelState");
    el.textContent = msg;
    el.className = "hint" + (kind ? " " + kind : "");
  }

  // Pulls the full workbook first so another device's additions and deletions
  // are reflected here, then pushes whatever is still pending. The push step
  // reuses that same fetch to check for duplicates, so a retry after a
  // half-finished sync cannot double-post, and a second Graph read isn't
  // needed just to do it.
  function syncNow(silent) {
    if (syncing || !window.Excel.configured() || !window.Excel.isSignedIn()) return Promise.resolve();

    syncing = true;
    renderSyncState();
    if (!silent) setExcelState("Checking the workbook...");

    return window.Excel.fetchAllRows().then(function (rows) {
      var merge = mergeFromWorkbook(rows);
      if (merge.added || merge.removed) renderAll();

      var queue = pending();
      if (!queue.length) {
        syncing = false;
        setExcelState(merge.added
          ? ("Pulled " + merge.added + " entr" + (merge.added === 1 ? "y" : "ies") + " logged on another device.")
          : "Everything is in the workbook.", "ok");
        renderSyncState();
        return;
      }

      setExcelState("Sending " + queue.length + " entr" + (queue.length === 1 ? "y" : "ies") + "...");

      var chain = Promise.resolve();
      queue.forEach(function (e) {
        chain = chain.then(function () {
          var already = rows.some(function (row) { return matchesRow(e, row); });
          if (already) { e.synced = true; save(); return; }
          return window.Excel.appendRow(rowValues(e)).then(function () {
            e.synced = true;
            save();
          });
        });
      });

      return chain.then(function () {
        syncing = false;
        setExcelState("Workbook is up to date.", "ok");
        renderAll();
      });
    }).catch(function (err) {
      syncing = false;
      setExcelState("Sync failed: " + err.message, "bad");
      renderAll();
    });
  }

  /* -- clock in / out ----------------------------------------------------- */

  var tick = null;

  function clockState() {
    try { return JSON.parse(localStorage.getItem(CLOCK) || "null"); }
    catch (e) { return null; }
  }

  function nowHHMM() {
    var d = new Date();
    return pad(d.getHours()) + ":" + pad(d.getMinutes());
  }

  function startClock() {
    localStorage.setItem(CLOCK, JSON.stringify({
      at: new Date().toISOString(), date: todayISO(), hhmm: nowHHMM()
    }));
    renderClock();
  }

  function stopClock() {
    var c = clockState();
    localStorage.removeItem(CLOCK);
    renderClock();
    if (!c) return;

    // Drop the stamps into the form rather than saving blind, so a note can be
    // added and a mis-tap can be corrected before it reaches payroll's copy.
    $("fDate").value = c.date;
    $("fIn").value = c.hhmm;
    $("fOut").value = nowHHMM();
    previewForm();
    $("fNotes").focus();
    $("entryForm").scrollIntoView({ behavior: "smooth", block: "center" });
  }

  function renderClock() {
    var c = clockState();
    var btn = $("clockBtn"), status = $("clockStatus");

    if (tick) { clearInterval(tick); tick = null; }

    if (!c) {
      btn.textContent = "Clock In";
      btn.classList.remove("running");
      status.hidden = true;
      return;
    }

    btn.textContent = "Clock Out";
    btn.classList.add("running");
    status.hidden = false;
    $("clockStart").textContent = time12(c.hhmm);

    var paint = function () {
      var secs = Math.max(0, Math.floor((Date.now() - new Date(c.at).getTime()) / 1000));
      $("clockElapsed").textContent =
        Math.floor(secs / 3600) + ":" + pad(Math.floor(secs / 60) % 60) + ":" + pad(secs % 60);
    };
    paint();
    tick = setInterval(paint, 1000);
  }

  /* -- form --------------------------------------------------------------- */

  function validate() {
    var date = $("fDate").value;
    var tin = $("fIn").value;
    var tout = $("fOut").value;

    if (!date) return { error: "Pick a date." };
    if (!tin) return { error: "Enter a time in." };
    if (!tout) return { error: "Enter a time out." };
    if (minutes(tout) <= minutes(tin)) {
      return { error: "Time Out must be after Time In." };
    }
    var dup = entries.some(function (e) {
      return e.date === date && e.tin === tin && e.tout === tout;
    });
    if (dup) return { error: "That exact entry is already logged." };

    return { date: date, tin: tin, tout: tout, hours: decimalHours(tin, tout) };
  }

  // Gives the answer (or the reason there isn't one) before Save is ever tapped.
  function previewForm() {
    var el = $("formPreview"), err = $("formError");
    var tin = $("fIn").value, tout = $("fOut").value;
    err.hidden = true;
    el.hidden = true;

    if (!tin || !tout) return;

    if (minutes(tout) === minutes(tin)) {
      err.textContent = "That is a zero-length shift. Adjust the times before saving.";
      err.hidden = false;
      return;
    }
    if (minutes(tout) < minutes(tin)) {
      err.textContent = "Time Out must be after Time In.";
      err.hidden = false;
      return;
    }
    el.hidden = false;
    el.textContent = fmt(decimalHours(tin, tout)) + " hours";
  }

  function submitForm(ev) {
    ev.preventDefault();
    var result = validate();
    var err = $("formError");

    if (result.error) {
      err.textContent = result.error;
      err.hidden = false;
      return;
    }
    err.hidden = true;

    entries.push({
      id: String(Date.now()) + Math.random().toString(36).slice(2, 7),
      date: result.date,
      tin: result.tin,
      tout: result.tout,
      notes: $("fNotes").value.trim(),
      synced: false
    });
    save();

    $("fIn").value = "";
    $("fOut").value = "";
    $("fNotes").value = "";
    $("formPreview").hidden = true;

    // Show the month the entry landed in, not whatever was being browsed.
    var p = parseISO(result.date);
    $("mYear").value = String(p.y);
    $("mMonth").value = String(p.m);

    renderAll();
    syncNow(true);
  }

  function deleteEntry(id) {
    var idx = -1;
    for (var i = 0; i < entries.length; i++) if (entries[i].id === id) idx = i;
    if (idx < 0) return;

    var e = entries[idx];
    if (!confirm("Delete " + e.date + ", " + time12(e.tin) + " to " + time12(e.tout) + "?")) return;

    entries.splice(idx, 1);
    save();
    renderAll();

    if (e.synced && window.Excel.isSignedIn()) {
      setExcelState("Removing the row from the workbook...");
      window.Excel.deleteRow(dateSerial(e.date), timeFraction(e.tin))
        .then(function () { setExcelState("Workbook is up to date.", "ok"); })
        .catch(function (err) {
          setExcelState("Deleted here, but the workbook row remains: " + err.message, "bad");
        });
    }
  }

  /* -- CSV backup --------------------------------------------------------- */

  function downloadCsv() {
    var rows = [["Date", "Day", "Time In", "Time Out", "Hours", "Period", "Notes"]];
    sorted().slice().reverse().forEach(function (e) {
      rows.push([e.date, dayName(e.date), time12(e.tin), time12(e.tout),
        fmt(decimalHours(e.tin, e.tout)), periodLabel(e.date), e.notes || ""]);
    });

    var csv = rows.map(function (r) {
      return r.map(function (c) { return '"' + String(c).replace(/"/g, '""') + '"'; }).join(",");
    }).join("\n");

    var url = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    var a = document.createElement("a");
    a.href = url;
    a.download = "work-hours-" + todayISO() + ".csv";
    a.click();
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  }



  // Entries carry a single synced flag, which is only meaningful relative to
  // one workbook. Repointing the app at a different file (switching tenants,
  // say) has to re-open the whole backlog, or the new workbook stays empty
  // while every entry still claims to be uploaded. Re-sending is safe: the
  // sync pass skips rows already present in the target.
  function reconcileWorkbookTarget() {
    var url = window.Excel.getSettings().workbookUrl;
    if (!url) return;

    var last = localStorage.getItem("wh.lastBook");
    if (last && last !== url) {
      entries.forEach(function (e) { e.synced = false; });
      save();
    }
    localStorage.setItem("wh.lastBook", url);
  }


  /* -- bulk import via #import= link --------------------------------------- */

  // Accepts a batch of entries handed over in the URL fragment, for backfilling
  // days that were tracked somewhere else. Same transport as the setup link, so
  // nothing is sent to a server. Imported rows arrive unsynced and then travel
  // the ordinary sync path, which means the workbook stays a product of the app
  // rather than something edited behind its back.
  function importEntriesFromHash() {
    var h = location.hash || "";
    if (h.indexOf("#import=") !== 0) return null;

    var raw = h.slice(8);
    history.replaceState({}, "", location.origin +
      location.pathname.replace(/index\.html$/, ""));

    var incoming;
    try {
      var pad = raw.replace(/-/g, "+").replace(/_/g, "/");
      while (pad.length % 4) pad += "=";
      incoming = JSON.parse(decodeURIComponent(escape(atob(pad))));
    } catch (e) {
      return { added: 0, skipped: 0, bad: true };
    }
    if (!incoming || !incoming.length) return null;

    var added = 0, skipped = 0;
    incoming.forEach(function (r) {
      if (!r || !r.date || !r.tin || !r.tout) { skipped++; return; }
      if (minutes(r.tout) <= minutes(r.tin)) { skipped++; return; }

      var dup = entries.some(function (e) {
        return e.date === r.date && e.tin === r.tin && e.tout === r.tout;
      });
      if (dup) { skipped++; return; }

      entries.push({
        id: String(Date.now()) + Math.random().toString(36).slice(2, 7),
        date: r.date,
        tin: r.tin,
        tout: r.tout,
        notes: r.notes || "",
        synced: false
      });
      added++;
    });

    if (added) save();
    return { added: added, skipped: skipped };
  }

  function showImportResult(res) {
    if (!res) return;
    var card = $("importCard"), msg = $("importMsg");
    card.hidden = false;

    if (res.bad) {
      msg.textContent = "That import link could not be read. Nothing was changed.";
      return;
    }

    var total = round(entries.filter(function (e) { return !e.synced; })
      .reduce(function (t, e) { return t + decimalHours(e.tin, e.tout); }, 0));

    msg.textContent = res.added + " entr" + (res.added === 1 ? "y" : "ies") + " added" +
      (res.skipped ? ", " + res.skipped + " skipped as duplicates or invalid" : "") +
      ". " + fmt(total) + " hours are queued for the workbook.";
  }

  /* -- settings ----------------------------------------------------------- */

  function loadSettingsForm() {
    var c = window.Excel.getSettings();
    $("sClient").value = c.clientId || "";
    $("sTenant").value = c.tenantId || "";
    $("sBook").value = c.workbookUrl || "";
    $("redirectUri").textContent = window.Excel.redirectUri();
    $("sRate").value = getPayRate();
  }

  function saveSettingsForm() {
    window.Excel.saveSettings({
      clientId: $("sClient").value,
      tenantId: $("sTenant").value,
      workbookUrl: $("sBook").value
    });
    reconcileWorkbookTarget();
    var msg = $("settingsMsg");
    msg.hidden = false;
    msg.textContent = window.Excel.configured()
      ? "Saved. Now tap Connect Excel."
      : "Saved, but one of the three is still blank.";
    refreshExcelUi();
  }


  function applyImported() {
    reconcileWorkbookTarget();
    loadSettingsForm();
    refreshExcelUi();
    var m = $("settingsMsg");
    m.hidden = false;
    m.textContent = window.Excel.configured()
      ? "Settings loaded from the setup link. Tap Connect Excel."
      : "Setup link loaded. Add the workbook URL to finish.";
  }

  /* -- wiring ------------------------------------------------------------- */

  function buildPickers() {
    var now = new Date();
    var mSel = $("mMonth"), ySel = $("mYear");

    MONTHS.forEach(function (name, i) {
      var o = document.createElement("option");
      o.value = String(i + 1); o.textContent = name;
      mSel.appendChild(o);
    });
    for (var y = now.getFullYear() - 1; y <= now.getFullYear() + 1; y++) {
      var o = document.createElement("option");
      o.value = String(y); o.textContent = String(y);
      ySel.appendChild(o);
    }
    mSel.value = String(now.getMonth() + 1);
    ySel.value = String(now.getFullYear());
  }

  function refreshExcelUi() {
    var configured = window.Excel.configured();
    var signedIn = window.Excel.isSignedIn();
    var book = window.Excel.bookInfo();

    $("setupCard").hidden = configured;
    if (!configured) {
      $("setupMsg").textContent =
        "The Excel connection is not configured yet, so entries are saving to this " +
        "device only. Fill in Settings below and they will upload, nothing is lost.";
    }

    $("signInBtn").hidden = !configured || signedIn;
    $("syncBtn").hidden = !signedIn;
    $("signOutBtn").hidden = !signedIn;

    var open = $("openBookBtn");
    if (book && book.webUrl) {
      open.hidden = false;
      open.href = book.webUrl;
      open.textContent = "Open workbook";
    } else {
      open.hidden = true;
    }

    if (configured && signedIn) {
      var n = pending().length;
      setExcelState(n ? n + " entr" + (n === 1 ? "y" : "ies") + " waiting to upload."
                      : "Connected. Workbook is up to date.", n ? "" : "ok");
    } else if (configured) {
      setExcelState("Connect your Microsoft account to write into the workbook.");
    }
    renderSyncState();
  }

  function init() {
    // Runs before anything reads settings, so a setup link is already applied
    // by the time sign-in or a redirect callback needs the client id.
    var imported = window.Excel.importFromHash();
    var importedEntries = importEntriesFromHash();

    buildPickers();
    $("fDate").value = todayISO();

    $("entryForm").addEventListener("submit", submitForm);
    $("fIn").addEventListener("change", previewForm);
    $("fOut").addEventListener("change", previewForm);
    $("mMonth").addEventListener("change", renderMonth);
    $("mYear").addEventListener("change", renderMonth);

    $("clockBtn").addEventListener("click", function () {
      if (clockState()) stopClock(); else startClock();
    });
    $("clockCancel").addEventListener("click", function () {
      localStorage.removeItem(CLOCK);
      renderClock();
    });

    $("entryList").addEventListener("click", function (ev) {
      var btn = ev.target.closest(".del");
      if (btn) deleteEntry(btn.getAttribute("data-id"));
    });

    $("signInBtn").addEventListener("click", function () { window.Excel.signIn(); });
    $("signOutBtn").addEventListener("click", function () {
      window.Excel.signOut();
      refreshExcelUi();
    });
    $("syncBtn").addEventListener("click", function () { syncNow(false); });
    $("syncBadge").addEventListener("click", function () { syncNow(false); });
    $("csvBtn").addEventListener("click", downloadCsv);
    $("saveSettings").addEventListener("click", saveSettingsForm);
    $("sRate").addEventListener("change", function () {
      setPayRate($("sRate").value);
      $("sRate").value = getPayRate();
      renderPeriod();
      renderMonth();
    });
    loadSettingsForm();

    if (imported) applyImported();
    if (importedEntries) showImportResult(importedEntries);

    // Tapping the setup link while the app is already open changes only the
    // fragment, which does not reload the page. Without this the link would
    // appear to do nothing.
    window.addEventListener("hashchange", function () {
      if (window.Excel.importFromHash()) applyImported();
      var batch = importEntriesFromHash();
      if (batch) {
        showImportResult(batch);
        renderAll();
        syncNow(true);
      }
    });

    window.Excel.onChange(refreshExcelUi);

    reconcileWorkbookTarget();
    renderClock();
    renderAll();

    window.Excel.handleRedirect().then(function (signedIn) {
      refreshExcelUi();
      if (window.Excel.isSignedIn()) {
        window.Excel.resolveBook()
          .then(function () { refreshExcelUi(); return syncNow(true); })
          .catch(function (err) {
            setExcelState("Cannot open the workbook: " + err.message, "bad");
          });
      }
    }).catch(function (err) {
      setExcelState("Sign-in failed: " + err.message, "bad");
      refreshExcelUi();
    });

    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("sw.js").catch(function () {});
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
