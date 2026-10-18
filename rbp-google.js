/* ─────────────────────────────────────────────────────────────────────────────
   RetireBlueprint Pro — Google sign-in connection (NEW, in testing)

   The customer signs in with Google and picks their own planning sheet in Google's
   file picker. From then on, THEIR browser talks directly to THEIR sheet through
   Google's Sheets API. Nothing passes through RetireBlueprint Pro, and the business
   account never needs access to the sheet.

   Permission requested: drive.file only — "only the files you pick with this app".

   window.RBPGoogle:
     signIn()                 → Promise<token>   (shows Google's sign-in popup if needed)
     pickSheet()              → Promise<{id,name}> (Google's own file picker)
     getValues(id, ranges[])  → Promise<{range: [[...]]}> read cells
     setValues(id, {range: [[...]]}) → Promise<result> write cells
     sheetId() / setSheetId(id)  remembered in this browser only
     signOut()
   ───────────────────────────────────────────────────────────────────────────── */
(function (root) {
  var CLIENT_ID = '524649270870-krd7ev3knt9b7i14fth2nggimj3brc31.apps.googleusercontent.com';
  var API_KEY   = 'AIzaSyAAtkUeJpxvpwUHOxb2k7wCr_Y1PFMBEnM'; // browser key: Picker only, retireblueprintpro.com only
  var APP_ID    = '524649270870';                              // Cloud project number (Picker needs it)
  var SCOPE     = 'https://www.googleapis.com/auth/drive.file';
  var KEY_SHEET = 'rbp_g_sheet_id';

  var token = null, tokenExp = 0, tokenClient = null;

  function loadScript(src) {
    return new Promise(function (res, rej) {
      if (document.querySelector('script[src="' + src + '"]')) { res(); return; }
      var s = document.createElement('script'); s.src = src; s.async = true;
      s.onload = function () { res(); }; s.onerror = function () { rej(new Error('Could not load ' + src)); };
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

  // forceConsent: show Google's permission screen again (used when the box was left unticked).
  function signIn(forceConsent) {
    if (!forceConsent && token && Date.now() < tokenExp - 60000) return Promise.resolve(token);
    return ready().then(function () {
      return new Promise(function (res, rej) {
        tokenClient = root.google.accounts.oauth2.initTokenClient({
          client_id: CLIENT_ID,
          scope: SCOPE,
          callback: function (r) {
            if (!r || r.error) { rej(new Error((r && (r.error_description || r.error)) || 'Sign-in was cancelled')); return; }
            // Google lets people untick the permission box. Without it the app can't open their sheet.
            if (!root.google.accounts.oauth2.hasGrantedAllScopes(r, SCOPE)) {
              var e = new Error('The permission box was left unticked. Click Sign in again and tick the box so the app can open the sheet you pick.');
              e.code = 'scope_missing'; rej(e); return;
            }
            token = r.access_token; tokenExp = Date.now() + (Number(r.expires_in) || 3600) * 1000;
            res(token);
          },
          error_callback: function (e) { rej(new Error((e && e.message) || 'Sign-in window was closed')); }
        });
        tokenClient.requestAccessToken({ prompt: forceConsent ? 'consent' : '' });
      });
    });
  }

  function pickSheet() {
    return signIn().then(function (tok) {
      return new Promise(function (res, rej) {
        var g = root.google.picker;
        var view = new g.DocsView(g.ViewId.SPREADSHEETS).setIncludeFolders(true).setMode(g.DocsViewMode.LIST);
        var picker = new g.PickerBuilder()
          .addView(view)
          .setOAuthToken(tok)
          .setDeveloperKey(API_KEY)
          .setAppId(APP_ID)
          .setTitle('Pick your RetireBlueprint Pro planning sheet')
          .setCallback(function (d) {
            if (d.action === g.Action.PICKED) {
              var doc = d.docs && d.docs[0];
              if (!doc) { rej(new Error('Nothing was picked')); return; }
              setSheetId(doc.id);
              res({ id: doc.id, name: doc.name });
            } else if (d.action === g.Action.CANCEL) {
              rej(new Error('Picker closed'));
            }
          })
          .build();
        picker.setVisible(true);
      });
    });
  }

  function api(method, url, body) {
    return signIn().then(function (tok) {
      return fetch(url, {
        method: method,
        headers: { 'Authorization': 'Bearer ' + tok, 'Content-Type': 'application/json' },
        body: body ? JSON.stringify(body) : undefined
      }).then(function (r) {
        return r.json().then(function (j) {
          if (!r.ok) throw new Error((j && j.error && j.error.message) || ('Google said ' + r.status));
          return j;
        });
      });
    });
  }

  function getValues(id, ranges) {
    var q = ranges.map(function (r) { return 'ranges=' + encodeURIComponent(r); }).join('&');
    return api('GET', 'https://sheets.googleapis.com/v4/spreadsheets/' + id + '/values:batchGet?valueRenderOption=UNFORMATTED_VALUE&' + q)
      .then(function (j) {
        var out = {};
        (j.valueRanges || []).forEach(function (vr, i) { out[ranges[i]] = vr.values || []; });
        return out;
      });
  }

  function setValues(id, map) {
    var data = Object.keys(map).map(function (r) { return { range: r, values: map[r] }; });
    return api('POST', 'https://sheets.googleapis.com/v4/spreadsheets/' + id + '/values:batchUpdate',
      { valueInputOption: 'USER_ENTERED', data: data });
  }

  function sheetId() { try { return localStorage.getItem(KEY_SHEET) || ''; } catch (e) { return ''; } }
  function setSheetId(id) { try { localStorage.setItem(KEY_SHEET, id); } catch (e) {} }

  function signOut() {
    try { if (token && root.google && root.google.accounts) root.google.accounts.oauth2.revoke(token, function () {}); } catch (e) {}
    token = null; tokenExp = 0;
  }

  root.RBPGoogle = {
    signIn: signIn, pickSheet: pickSheet, getValues: getValues, setValues: setValues,
    sheetId: sheetId, setSheetId: setSheetId, signOut: signOut
  };
})(typeof window !== 'undefined' ? window : this);
