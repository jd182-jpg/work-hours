/* ---------------------------------------------------------------------------
 * excel.js — Microsoft sign-in (OAuth 2.0 auth code + PKCE, no library) and
 * the handful of Graph calls needed to append rows to the workbook table.
 *
 * Why hand-rolled instead of MSAL: this is four endpoints and ~150 lines, and
 * it means the phone loads no third-party script to log a day's hours.
 *
 * The Azure app registration MUST list the redirect URI under the "Single-page
 * application" platform. Registered as "Web" instead, the token endpoint
 * refuses the browser's CORS request and sign-in fails at the last step.
 * ------------------------------------------------------------------------- */
(function () {
  "use strict";

  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = "openid profile offline_access Files.ReadWrite.All";
  var TOKEN_KEY = "ih.token";
  var BOOK_KEY = "ih.book";
  var VERIFIER_KEY = "ih.pkce";
  var STATE_KEY = "ih.state";

  var SETTINGS_KEY = "ih.settings";
  var listeners = [];

  // config.js supplies defaults; anything saved on this device wins. Keeping the
  // tenant and workbook URL in localStorage means the published repo carries no
  // detail about where the file actually lives.
  function S() {
    var base = window.CONFIG || {};
    var over = {};
    try { over = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}"); } catch (e) {}
    var pick = function (a, b) {
      if (a && String(a).indexOf("PASTE") !== 0) return a;
      if (b && String(b).indexOf("PASTE") !== 0) return b;
      return "";
    };
    return {
      clientId: pick(over.clientId, base.clientId),
      tenantId: pick(over.tenantId, base.tenantId),
      workbookUrl: pick(over.workbookUrl, base.workbookUrl),
      tableName: base.tableName || "HoursLog",
      decimals: base.decimals || 2
    };
  }

  function getSettings() { return S(); }

  function saveSettings(next) {
    var prev = S();
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({
      clientId: (next.clientId || "").trim(),
      tenantId: (next.tenantId || "").trim(),
      workbookUrl: (next.workbookUrl || "").trim()
    }));
    var now = S();
    // A different workbook means the cached drive/item ids are stale; a
    // different app or tenant means the stored token is no longer valid.
    if (now.workbookUrl !== prev.workbookUrl) localStorage.removeItem(BOOK_KEY);
    if (now.clientId !== prev.clientId || now.tenantId !== prev.tenantId) clearToken();
    emit();
  }

  function authority(path) {
    return "https://login.microsoftonline.com/" + encodeURIComponent(S().tenantId) +
      "/oauth2/v2.0/" + path;
  }

  // The page URL with no query or hash, which is what Azure must have on file.
  // "index.html" is trimmed so opening the bare folder and opening the file
  // itself produce the same value; Azure matches the redirect URI exactly and
  // one registered entry has to cover both.
  function redirectUri() {
    return location.origin + location.pathname.replace(/index\.html$/, "");
  }

  function configured() {
    var c = S();
    return !!(c.clientId && c.tenantId && c.workbookUrl);
  }

  /* -- small helpers ------------------------------------------------------ */

  function b64url(bytes) {
    var s = "";
    for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  }

  function randomString(len) {
    var a = new Uint8Array(len);
    crypto.getRandomValues(a);
    return b64url(a).slice(0, len);
  }

  function sha256(text) {
    return crypto.subtle.digest("SHA-256", new TextEncoder().encode(text))
      .then(function (buf) { return b64url(new Uint8Array(buf)); });
  }

  function readToken() {
    try { return JSON.parse(localStorage.getItem(TOKEN_KEY) || "null"); }
    catch (e) { return null; }
  }

  function writeToken(data) {
    // expires_in is seconds from now; keep a 60s safety margin.
    data.expiresAt = Date.now() + (data.expires_in - 60) * 1000;
    localStorage.setItem(TOKEN_KEY, JSON.stringify(data));
  }

  function clearToken() {
    localStorage.removeItem(TOKEN_KEY);
    localStorage.removeItem(BOOK_KEY);
  }

  function emit() {
    listeners.forEach(function (fn) { try { fn(); } catch (e) {} });
  }

  /* -- sign-in ------------------------------------------------------------ */

  function signIn() {
    var verifier = randomString(64);
    var state = randomString(16);
    sessionStorage.setItem(VERIFIER_KEY, verifier);
    sessionStorage.setItem(STATE_KEY, state);
    return sha256(verifier).then(function (challenge) {
      var q = new URLSearchParams({
        client_id: S().clientId,
        response_type: "code",
        redirect_uri: redirectUri(),
        response_mode: "query",
        scope: SCOPES,
        state: state,
        code_challenge: challenge,
        code_challenge_method: "S256"
      });
      location.assign(authority("authorize") + "?" + q.toString());
    });
  }

  function signOut() {
    clearToken();
    emit();
  }

  // Called once on load. Completes the redirect if we came back with a code.
  function handleRedirect() {
    var params = new URLSearchParams(location.search);
    var code = params.get("code");
    var err = params.get("error");

    if (!code && !err) return Promise.resolve(null);

    // Strip the code out of the address bar either way.
    history.replaceState({}, "", redirectUri());

    if (err) {
      return Promise.reject(new Error(
        params.get("error_description") || err));
    }
    if (params.get("state") !== sessionStorage.getItem(STATE_KEY)) {
      return Promise.reject(new Error("Sign-in state mismatch. Try again."));
    }

    var verifier = sessionStorage.getItem(VERIFIER_KEY);
    sessionStorage.removeItem(VERIFIER_KEY);
    sessionStorage.removeItem(STATE_KEY);

    return postToken({
      client_id: S().clientId,
      grant_type: "authorization_code",
      code: code,
      redirect_uri: redirectUri(),
      code_verifier: verifier
    }).then(function (data) { writeToken(data); emit(); return data; });
  }

  function postToken(body) {
    return fetch(authority("token"), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(body).toString()
    }).then(function (r) {
      return r.json().then(function (j) {
        if (!r.ok) {
          throw new Error(j.error_description || j.error || ("HTTP " + r.status));
        }
        return j;
      });
    });
  }

  // Returns a valid access token, refreshing silently when it has aged out.
  function getToken() {
    var t = readToken();
    if (!t) return Promise.reject(new Error("Not signed in"));
    if (Date.now() < t.expiresAt) return Promise.resolve(t.access_token);
    if (!t.refresh_token) { clearToken(); emit(); return Promise.reject(new Error("Session expired")); }

    return postToken({
      client_id: S().clientId,
      grant_type: "refresh_token",
      refresh_token: t.refresh_token,
      scope: SCOPES
    }).then(function (data) {
      writeToken(data);
      return data.access_token;
    }).catch(function (e) {
      clearToken(); emit();
      throw new Error("Session expired, sign in again");
    });
  }

  function isSignedIn() { return !!readToken(); }

  /* -- Graph -------------------------------------------------------------- */

  function graph(path, opts) {
    opts = opts || {};
    return getToken().then(function (token) {
      var headers = { Authorization: "Bearer " + token };
      if (opts.body) headers["Content-Type"] = "application/json";
      return fetch(GRAPH + path, {
        method: opts.method || "GET",
        headers: headers,
        body: opts.body ? JSON.stringify(opts.body) : undefined
      });
    }).then(function (r) {
      if (r.status === 204) return null;
      return r.text().then(function (text) {
        var json = text ? JSON.parse(text) : null;
        if (!r.ok) {
          var msg = (json && json.error && json.error.message) || ("HTTP " + r.status);
          var e = new Error(msg);
          e.status = r.status;
          throw e;
        }
        return json;
      });
    });
  }

  // Turn the workbook's share/browse URL into the driveId + itemId Graph needs.
  // Cached, because this never changes for a given file.
  function resolveBook() {
    var cached = null;
    try { cached = JSON.parse(localStorage.getItem(BOOK_KEY) || "null"); } catch (e) {}
    if (cached && cached.driveId && cached.itemId) return Promise.resolve(cached);

    var encoded = "u!" + btoa(S().workbookUrl)
      .replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");

    return graph("/shares/" + encoded + "/driveItem?$select=id,name,webUrl,parentReference")
      .then(function (item) {
        var book = {
          driveId: item.parentReference.driveId,
          itemId: item.id,
          name: item.name,
          webUrl: item.webUrl
        };
        localStorage.setItem(BOOK_KEY, JSON.stringify(book));
        return book;
      });
  }

  function bookInfo() {
    try { return JSON.parse(localStorage.getItem(BOOK_KEY) || "null"); }
    catch (e) { return null; }
  }

  function tablePath(book) {
    return "/drives/" + book.driveId + "/items/" + book.itemId +
      "/workbook/tables/" + encodeURIComponent(S().tableName || "HoursLog");
  }

  /* -- the one operation the app needs ------------------------------------ */

  // Append one row. The generated workbook ships with a single blank row so the
  // table is valid; the first real entry overwrites it rather than landing
  // underneath and leaving a gap in the log.
  function appendRow(values) {
    return resolveBook().then(function (book) {
      var base = tablePath(book);
      return graph(base + "/rows?$top=1").then(function (res) {
        var rows = (res && res.value) || [];
        var firstIsBlank = rows.length === 1 &&
          (rows[0].values[0][0] === "" || rows[0].values[0][0] === null);

        if (firstIsBlank) {
          return graph(base + "/rows/itemAt(index=0)", {
            method: "PATCH",
            body: { values: [values] }
          });
        }
        return graph(base + "/rows/add", {
          method: "POST",
          body: { values: [values] }
        });
      });
    });
  }

  // Pull every Date cell already in the workbook, so a re-sync can skip rows
  // that are present and avoid double-counting a day.
  function existingKeys() {
    return resolveBook().then(function (book) {
      return graph(tablePath(book) + "/rows?$select=values");
    }).then(function (res) {
      var keys = {};
      ((res && res.value) || []).forEach(function (row) {
        var v = row.values[0];
        if (v && v[0] !== "" && v[0] !== null) {
          keys[String(v[0]) + "|" + String(v[2])] = true;
        }
      });
      return keys;
    });
  }


  // Remove a row from the workbook by matching its date + time-in, so deleting
  // an entry in the app does not leave a ghost day in payroll's copy.
  function deleteRow(dateSerial, inFrac) {
    return resolveBook().then(function (book) {
      var base = tablePath(book);
      return graph(base + "/rows?$select=values").then(function (res) {
        var rows = (res && res.value) || [];
        for (var i = 0; i < rows.length; i++) {
          var v = rows[i].values[0];
          if (!v) continue;
          if (near(v[0], dateSerial) && near(v[2], inFrac)) {
            return graph(base + "/rows/itemAt(index=" + i + ")", { method: "DELETE" });
          }
        }
        return null; // already gone
      });
    });
  }

  // Excel round-trips serials as floats; compare with a tolerance.
  function near(a, b) {
    if (a === null || a === "" || a === undefined) return false;
    return Math.abs(Number(a) - Number(b)) < 1e-6;
  }

  window.Excel = {
    configured: configured,
    getSettings: getSettings,
    saveSettings: saveSettings,
    redirectUri: redirectUri,
    isSignedIn: isSignedIn,
    signIn: signIn,
    signOut: signOut,
    handleRedirect: handleRedirect,
    appendRow: appendRow,
    existingKeys: existingKeys,
    deleteRow: deleteRow,
    resolveBook: resolveBook,
    bookInfo: bookInfo,
    onChange: function (fn) { listeners.push(fn); }
  };
})();
