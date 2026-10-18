/* ─────────────────────────────────────────────────────────────────────────────
   RetireBlueprint Pro — Google sign-in connection

   The customer signs in with Google and picks their own planning sheet in Google's
   file picker. From then on, THEIR browser talks directly to THEIR sheet through
   Google's Sheets API. Nothing passes through RetireBlueprint Pro, and the business
   account never needs access to the sheet.

   Permission requested: drive.file only — "only the files you pick with this app".

   window.RBPGoogle:
     signIn(forceConsent)       → Promise<token>  (Google popup; must follow a click)
     hasToken()                 → true if a usable sign-in is held for this browser tab
     pickSheet()                → Promise<{id,name}> (Google's own file picker)
     getValues(id, ranges[])    → Promise<{range: [[...]]}>
     setValues(id, {range: [[...]]})
     getSpreadsheet(id, query)  → raw spreadsheets.get (grid data etc.)
     batchUpdate(id, requests)  → raw spreadsheets.batchUpdate
     valuesBatchUpdate(id, data)→ raw spreadsheets.values.batchUpdate (USER_ENTERED)
     sheetId() / setSheetId(id) remembered in this browser only
     signOut()
   ───────────────────────────────────────────────────────────────────────────── */
(function (root) {
  var CLIENT_ID = '524649270870-krd7ev3knt9b7i14fth2nggimj3brc31.apps.googleusercontent.com';
  var API_KEY   = 'AIzaSyAAtkUeJpxvpwUHOxb2k7wCr_Y1PFMBEnM'; // browser key: Picker only, retireblueprintpro.com only
  var APP_ID    = '524649270870';                              // Cloud project number (Picker needs it)
  var SCOPE     = 'https://www.googleapis.com/auth/drive.file';
  var KEY_SHEET = 'rbp_g_sheet_id';
  var KEY_TOKEN = 'rbp_g_token';   // sessionStorage: survives page changes, cleared when the tab closes
  var SHEETS    = 'https://sheets.googleapis.com/v4/spreadsheets/';

  var token = null, tokenExp = 0, tokenClient = null;

  // Re-use a sign-in from an earlier page in this same browser tab (tokens last ~1 hour).
  try {
    var saved = JSON.parse(sessionStorage.getItem(KEY_TOKEN) || 'null');
    if (saved && saved.t && saved.e > Date.now() + 60000) { token = saved.t; tokenExp = saved.e; }
  } catch (e) {}

  function loadScript(src) {
    return new Promise(function (res, rej) {
      if (document.querySelector('script[data-rbp-src="' + src + '"]')) { res(); return; }
      var s = document.createElement('script'); s.async = true; s.setAttribute('data-rbp-src', src);
      s.onload = function () { res(); }; s.onerror = function () { rej(new Error('Could not load Google sign-in. Check your connection.')); };
      s.src = src;
      document.head.appendChild(s);
    });
  }

  function ready() {
    return Promise.all([
      loadScript('https://accounts.google.com/gsi/client'),
      loadScript('https://apis.google.com/js/api.js')
    ]).then(function () {
      return new Promise(function (res) { root.gapi.load('picker', res); });
    });
  }

  function hasToken() { return !!(token && Date.now() < tokenExp - 60000); }

  // The Google window must open in the SAME moment as the click, or browsers block it as a pop-up.
  // So Google's sign-in tools are loaded ahead of time (preload), and the click opens the window
  // right away. Results come back through `pending`.
  var pending = null;
  function finish(fn, v) { var p = pending; pending = null; if (p) p[fn](v); }
  function initClient() {
    tokenClient = root.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPE,
      callback: function (r) {
        if (!r || r.error) { finish('rej', new Error((r && (r.error_description || r.error)) || 'Sign-in was cancelled')); return; }
        // Google lets people untick the permission box. Without it the app can't open their sheet.
        if (!root.google.accounts.oauth2.hasGrantedAllScopes(r, SCOPE)) {
          var e = new Error('The permission box was left unticked. Click Sign in again and tick the box so the app can open the sheet you pick.');
          e.code = 'scope_missing'; finish('rej', e); return;
        }
        token = r.access_token; tokenExp = Date.now() + (Number(r.expires_in) || 3600) * 1000;
        try { sessionStorage.setItem(KEY_TOKEN, JSON.stringify({ t: token, e: tokenExp })); } catch (e2) {}
        finish('res', token);
      },
      error_callback: function (e) {
        var t = (e && e.type) || '';
        finish('rej', new Error(t === 'popup_failed_to_open' ? 'popup blocked' : ((e && e.message) || 'Sign-in window was closed')));
      }
    });
  }
  var preloading = null;
  function preload() {
    if (!preloading) preloading = ready().then(function () { if (!tokenClient) initClient(); });
    return preloading;
  }
  function openGoogle(forceConsent) {
    return new Promise(function (res, rej) {
      if (pending) pending.rej(new Error('Sign-in was cancelled'));
      pending = { res: res, rej: rej };
      tokenClient.requestAccessToken({ prompt: forceConsent ? 'consent' : '' });
    });
  }

  // forceConsent: show Google's permission screen again (used when the box was left unticked).
  function signIn(forceConsent) {
    if (!forceConsent && hasToken()) return Promise.resolve(token);
    if (tokenClient) return openGoogle(forceConsent);          // ready: open in the same click
    return preload().then(function () { return openGoogle(forceConsent); });
  }

  function pickSheet(title) {
    return signIn().then(function (tok) {
      return ready().then(function () {
        return new Promise(function (res, rej) {
          var g = root.google.picker;
          var view = new g.DocsView(g.ViewId.SPREADSHEETS).setIncludeFolders(true).setMode(g.DocsViewMode.LIST);
          var picker = new g.PickerBuilder()
            .addView(view)
            .setOAuthToken(tok)
            .setDeveloperKey(API_KEY)
            .setAppId(APP_ID)
            .setTitle(title || 'Pick your RetireBlueprint Pro planning sheet')
            .setCallback(function (d) {
              if (d.action === g.Action.PICKED) {
                var doc = d.docs && d.docs[0];
                if (!doc) { rej(new Error('Nothing was picked')); return; }
                res({ id: doc.id, name: doc.name });
              } else if (d.action === g.Action.CANCEL) {
                rej(new Error('Picker closed'));
              }
            })
            .build();
          picker.setVisible(true);
        });
      });
    });
  }

  function api(method, url, body, _retried) {
    if (!hasToken()) {
      var e = new Error('Please sign in with Google again to continue.'); e.code = 'need_signin';
      return Promise.reject(e);
    }
    return fetch(url, {
      method: method,
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: body ? JSON.stringify(body) : undefined
    }).then(function (r) {
      if (r.status === 401) {               // token expired or revoked
        token = null; tokenExp = 0;
        try { sessionStorage.removeItem(KEY_TOKEN); } catch (e) {}
        var e2 = new Error('Please sign in with Google again to continue.'); e2.code = 'need_signin'; throw e2;
      }
      if ((r.status === 429 || r.status >= 500) && !_retried) {   // Google busy: wait and try once more
        return new Promise(function (res) { setTimeout(res, 1500); }).then(function () { return api(method, url, body, true); });
      }
      return r.json().then(function (j) {
        if (!r.ok) {
          var msg = (j && j.error && j.error.message) || ('Google said ' + r.status);
          var e3 = new Error(msg); e3.status = r.status; throw e3;
        }
        return j;
      });
    });
  }

  function getValues(id, ranges) {
    var q = ranges.map(function (r) { return 'ranges=' + encodeURIComponent(r); }).join('&');
    return api('GET', SHEETS + id + '/values:batchGet?valueRenderOption=UNFORMATTED_VALUE&' + q)
      .then(function (j) {
        var out = {};
        (j.valueRanges || []).forEach(function (vr, i) { out[ranges[i]] = vr.values || []; });
        return out;
      });
  }

  function valuesBatchUpdate(id, data) {
    if (!data.length) return Promise.resolve({});
    return api('POST', SHEETS + id + '/values:batchUpdate', { valueInputOption: 'USER_ENTERED', data: data });
  }

  function setValues(id, map) {
    return valuesBatchUpdate(id, Object.keys(map).map(function (r) { return { range: r, values: map[r] }; }));
  }

  function getSpreadsheet(id, query) {
    return api('GET', SHEETS + id + (query ? '?' + query : ''));
  }

  function batchUpdate(id, requests) {
    if (!requests.length) return Promise.resolve({});
    return api('POST', SHEETS + id + ':batchUpdate', { requests: requests });
  }

  function sheetId() { try { return localStorage.getItem(KEY_SHEET) || ''; } catch (e) { return ''; } }
  function setSheetId(id) { try { localStorage.setItem(KEY_SHEET, id); } catch (e) {} }

  function signOut() {
    try { if (token && root.google && root.google.accounts) root.google.accounts.oauth2.revoke(token, function () {}); } catch (e) {}
    token = null; tokenExp = 0;
    try { sessionStorage.removeItem(KEY_TOKEN); } catch (e) {}
  }

  root.RBPGoogle = {
    signIn: signIn, hasToken: hasToken, pickSheet: pickSheet,
    getValues: getValues, setValues: setValues,
    preload: preload, getSpreadsheet: getSpreadsheet, batchUpdate: batchUpdate, valuesBatchUpdate: valuesBatchUpdate,
    sheetId: sheetId, setSheetId: setSheetId, signOut: signOut
  };
  // Get Google's sign-in tools ready in the background so a click opens the window instantly.
  try { if (root.document) setTimeout(function () { preload().catch(function () { preloading = null; }); }, 0); } catch (e) {}
})(typeof window !== 'undefined' ? window : this);
