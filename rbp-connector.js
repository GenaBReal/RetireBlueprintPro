/* ============================================================================
   RetireBlueprint Pro — SHARED CONNECTOR  (the one file every page talks through)
   ----------------------------------------------------------------------------
   WHAT THIS IS:
   Every page used to carry its own private copy of "talk to the Google Sheet"
   code. They drifted apart — some had timeouts and error messages, some failed
   silently — which is why loading worked on one page but not another.

   This file replaces all of those copies with ONE reliable version:
       RBP.load(onOk, onErr)             -> reads the plan from the sheet
       RBP.save(action, payload, onOk, onErr)  -> saves data to the sheet

   It has a built-in timeout AND one automatic retry, so a single hiccup no
   longer leaves a page blank. Fix a connection bug here once, and every page
   that includes this file is fixed.

   HOW TO USE (already wired into the pages — you don't normally touch this):
   In the <head> of each page:
       <script src="https://genabreal.github.io/RetireBlueprintPro/rbp-connector.js"></script>

   You should rarely need to edit this file. The only line that could ever
   change is the deployment URL below, and only if you ever re-deploy the
   Apps Script to a brand-new URL (you normally won't — re-deploying the
   existing deployment keeps the same URL).
   ============================================================================ */

window.RBP = window.RBP || {};

(function () {

  /* ---- The one place the connector URL lives ---- */
  RBP.SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxR40oJlUa6p8sCANeoHNfXU5XCnlwOudWfkgKWgorlzrXiqXAT8NbPuJU1x0QbfUg7/exec';

  /* ---- Settings ---- */
  var READ_TIMEOUT_MS = 60000;   // give a cold sheet time to wake up
  var MAX_ATTEMPTS    = 2;       // 1 retry on failure before we report an error

  /* ---- Where the customer's Sheet ID is stored ---- */
  RBP.getSheetId = function () {
    try { return localStorage.getItem('rbp_sheet_id') || ''; }
    catch (e) { return ''; }
  };

  /* ========================================================================
     RBP.load(onOk, onErr)
       onOk(data)        -> called with the full plan object on success
       onErr(message, kind) -> called with a friendly message if it fails
                               kind is one of: 'no-sheet','timeout','network','server'
     Uses JSONP (a <script> tag) because a normal fetch is blocked by CORS on
     a hard refresh. Times out, and retries once automatically before failing.
     ======================================================================== */
  RBP.load = function (onOk, onErr, _attempt) {
    onOk  = onOk  || function () {};
    onErr = onErr || function () {};
    _attempt = _attempt || 1;

    var sheetId = RBP.getSheetId();
    if (!sheetId) { onErr('No sheet connected. Please complete Setup first.', 'no-sheet'); return; }

    var cb = 'rbpLoad_' + Date.now() + '_' + Math.floor(Math.random() * 1e6);
    var s  = document.createElement('script');
    var settled = false;

    var timer = setTimeout(function () { finish(null, 'timeout'); }, READ_TIMEOUT_MS);

    function cleanup() {
      clearTimeout(timer);
      try { delete window[cb]; } catch (e) { window[cb] = undefined; }
      if (s && s.parentNode) s.parentNode.removeChild(s);
    }

    function finish(data, kind) {
      if (settled) return;
      settled = true;
      cleanup();

      if (data && !data.error) { onOk(data); return; }   // success

      // failure: retry once, then report
      if (_attempt < MAX_ATTEMPTS) { RBP.load(onOk, onErr, _attempt + 1); return; }

      var msg = (data && data.error) ? data.error
              : kind === 'timeout' ? 'Timeout — make sure your sheet is shared with retireblueprintpro@gmail.com as Editor, then refresh.'
              : 'Network error. Check your connection and refresh.';
      onErr(msg, (data && data.error) ? 'server' : kind);
    }

    window[cb] = function (data) { finish(data, (data && data.error) ? 'server' : null); };

    s.crossOrigin = 'anonymous';
    s.src = RBP.SCRIPT_URL +
            '?authuser=0&action=read&sheetId=' + encodeURIComponent(sheetId) +
            '&callback=' + cb + '&_=' + Date.now();
    s.onerror = function () { finish(null, 'network'); };
    document.head.appendChild(s);
  };

  /* ========================================================================
     RBP.save(action, payload, onOk, onErr)
       action  -> 'save' (Inputs) or 'saveCheckIn' (Annual Check-In), etc.
       payload -> the data object to store
       onOk()  -> called once the save request has been sent
       onErr(message) -> called if the request itself failed
     Sends a POST (base64-encoded, same as before) and stamps rbp_last_saved
     so other open tabs know to refresh. Page-side "verify by reading back"
     can still be layered on top by calling RBP.load after onOk.
     ======================================================================== */
  RBP.save = function (action, payload, onOk, onErr) {
    onOk  = onOk  || function () {};
    onErr = onErr || function () {};

    var sheetId = RBP.getSheetId();
    if (!sheetId) { onErr('No sheet connected. Please complete Setup first.'); return; }

    var b64;
    try { b64 = btoa(unescape(encodeURIComponent(JSON.stringify(payload)))); }
    catch (e) { onErr('Could not encode data: ' + (e.message || e)); return; }

    fetch(RBP.SCRIPT_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: 'action=' + encodeURIComponent(action) +
            '&enc=b64&sheetId=' + encodeURIComponent(sheetId) +
            '&data=' + encodeURIComponent(b64)
    })
    .then(function () {
      try { localStorage.setItem('rbp_last_saved', Date.now().toString()); } catch (e) {}
      onOk();
    })
    .catch(function (err) { onErr((err && err.message) || String(err)); });
  };

})();
/* ============================================================================ */
