/* ─────────────────────────────────────────────────────────────────────────────
   RetireBlueprint Pro — Google sign-in bridge

   Every page talks to the old shared connector (Apps Script) with JSONP reads and
   POST saves. For a customer who connected with "Sign in with Google", this file
   quietly answers those same calls from THEIR browser, straight from THEIR sheet,
   using the connector's own logic (rbp-gs-core.js) on top of the adapter (rbp-gs-shim.js).

   • Customers on the old connection: this file does NOTHING. Every request goes out
     exactly as before.
   • Google-connected customers: requests for their own sheet never leave the browser
     for RetireBlueprint Pro. The sample plan and the setup access-code check still use
     the old connector (they don't involve customer data).

   Must load in <head> BEFORE the page's own scripts, after rbp-google.js, rbp-gs-shim.js
   and rbp-gs-core.js.
   ───────────────────────────────────────────────────────────────────────────── */
(function (root) {
  var SCRIPT_BASE = 'https://script.google.com/macros/s/AKfycbxR40oJlUa6p8sCANeoHNfXU5XCnlwOudWfkgKWgorlzrXiqXAT8NbPuJU1x0QbfUg7/exec';
  var MODE_KEY = 'rbp_conn';

  function mode() { try { return localStorage.getItem(MODE_KEY) || ''; } catch (e) { return ''; } }
  function gSheet() { return root.RBPGoogle ? root.RBPGoogle.sheetId() : ''; }
  function isGoogleMode() { return mode() === 'google' && !!gSheet(); }

  function parseParams(str) {
    var out = {};
    String(str || '').replace(/^\?/, '').split('&').forEach(function (pair) {
      if (!pair) return;
      var i = pair.indexOf('='), k = i < 0 ? pair : pair.slice(0, i), v = i < 0 ? '' : pair.slice(i + 1);
      try { out[decodeURIComponent(k)] = decodeURIComponent(v.replace(/\+/g, ' ')); } catch (e) { out[k] = v; }
    });
    return out;
  }
  function decodePayload(p) {
    var raw = p.data || '';
    if (p.enc === 'b64') raw = decodeURIComponent(escape(atob(raw)));
    return raw ? JSON.parse(raw) : {};
  }
  // Should this request be answered locally?
  function claims(url, params) {
    if (!isGoogleMode()) return false;
    if (String(url).indexOf(SCRIPT_BASE) !== 0) return false;
    if ((params.action || 'read') === 'unlockSetup') return false;   // no sheet involved
    return params.sheetId === gSheet();
  }

  // ── Sign-in prompt (Google requires a click to open its sign-in window) ────
  var waiting = [], overlay = null;
  function needSignIn() {
    return new Promise(function (res, rej) {
      waiting.push({ res: res, rej: rej });
      showOverlay();
    });
  }
  function showOverlay(errMsg) {
    if (!document.body) { document.addEventListener('DOMContentLoaded', function () { showOverlay(errMsg); }); return; }
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.setAttribute('role', 'dialog');
      overlay.style.cssText = 'position:fixed;inset:0;z-index:2147483646;background:rgba(4,20,38,.82);display:flex;align-items:center;justify-content:center;padding:16px;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif';
      overlay.innerHTML =
        '<div style="background:#051628;border:1px solid rgba(28,199,208,.45);border-radius:14px;max-width:420px;width:100%;padding:22px 20px;color:#E8F1F5;text-align:center;box-shadow:0 10px 40px rgba(0,0,0,.5)">' +
        '<div style="font-weight:800;font-size:1.15rem;margin-bottom:6px">Continue with Google</div>' +
        '<div style="color:#9FB6C3;font-size:.95rem;line-height:1.5;margin-bottom:16px">To open your plan, confirm it\'s you. Your browser reads your own Google Sheet directly — RetireBlueprint Pro never sees it.</div>' +
        '<button type="button" id="rbpg-go" style="background:#1CC7D0;color:#041E2F;border:0;border-radius:8px;padding:11px 18px;font-weight:800;font-size:1rem;cursor:pointer">Continue with Google</button>' +
        '<div id="rbpg-err" style="color:#ff9b9b;font-size:.9rem;margin-top:12px;display:none"></div></div>';
      document.body.appendChild(overlay);
      overlay.querySelector('#rbpg-go').onclick = function () {
        var btn = this; btn.disabled = true; btn.textContent = 'Opening Google…';
        root.RBPGoogle.signIn().then(function () {
          hideOverlay();
          var w = waiting; waiting = [];
          w.forEach(function (x) { x.res(); });
        }).catch(function (e) {
          btn.disabled = false; btn.textContent = 'Continue with Google';
          var el = overlay.querySelector('#rbpg-err'); el.style.display = 'block'; el.textContent = e.message;
        });
      };
    }
    overlay.style.display = 'flex';
  }
  function hideOverlay() { if (overlay) overlay.style.display = 'none'; }

  function withSignIn(fn) {
    var run = function () {
      return fn().catch(function (e) {
        if (e && e.code === 'need_signin') return needSignIn().then(fn);
        throw e;
      });
    };
    if (root.RBPGoogle.hasToken()) return run();
    return needSignIn().then(run);
  }

  // ── Running the connector logic locally ───────────────────────────────────
  var S = function () { return root.RBPShim; };
  function core(book, others) { return root.RBPCoreFactory(S().env(book, others)); }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  // Port of rbpWaitForReadSettle: wait until the safe-spend cell stops moving (max ~21s).
  function settle(id) {
    var prev = null, stable = 0, w = 0;
    function step() {
      if (w >= 15) return Promise.resolve();
      return sleep(w === 0 ? 400 : 1500).then(function () {
        w++;
        return root.RBPGoogle.getValues(id, ['Inputs!B119']).then(function (v) {
          var cur = Number(v['Inputs!B119'] && v['Inputs!B119'][0] && v['Inputs!B119'][0][0]) || 0;
          if (prev !== null && Math.abs(cur - prev) < 1) { if (++stable >= 2) return; } else stable = 0;
          prev = cur;
          return step();
        });
      });
    }
    return step();
  }

  var readCache = {};   // same 8-second sharing the connector does, per page
  function readTabs(fast) {
    var t = { 'Inputs': null, '_StressTests': null, 'Annual Check-In': 'A1:F130' };
    if (!fast) t['Master'] = 'A8:CD200';
    return t;
  }

  function doRead(id, fast) {
    var key = id + (fast ? '_fast' : '_full'), hit = readCache[key];
    if (hit && Date.now() - hit.t < 8000) return Promise.resolve(JSON.parse(hit.j));
    return S().load(id, { tabs: readTabs(fast) }).then(function (book) {
      var inp = book.getSheetByName('Inputs');
      if (!inp) return { error: 'No Inputs sheet' };
      var c = core(book);
      // Same self-repairs the connector runs on every page load.
      c.rbpHealStress(book, inp, false);
      try { c.rbpTrimMasterBloat(book); } catch (e) {}
      try { c.rbpEnsureCalcSettings(book); } catch (e) {}
      try { c.rbpHealSolver(inp); } catch (e) {}
      var changed = book.hasChanges();
      return (changed ? book.flush() : Promise.resolve()).then(function () {
        return fast ? null : settle(id);
      }).then(function () {
        return (changed || !fast) ? S().load(id, { tabs: readTabs(fast) }) : book;
      }).then(function (b2) {
        var result = core(b2).readAll(b2, b2.getSheetByName('Inputs'), fast);
        if (!fast) {
          var has = result && ((result.portfolio && result.portfolio.accounts && result.portfolio.accounts.length) ||
                               (result.projections && result.projections.years && result.projections.years.length));
          if (has) readCache[key] = { t: Date.now(), j: JSON.stringify(result) };
        }
        return result;
      });
    });
  }

  function withBook(id, tabs, fn) {
    return S().load(id, { tabs: tabs }).then(function (book) {
      var out = fn(book, core(book));
      return book.flush().then(function () { readCache = {}; return out; });
    });
  }

  // Tiller: the customer's own Tiller sheet is opened READ-ONLY with their Google sign-in
  // (they pick it once in Google's file list), then the connector's own rbpTiller() runs
  // exactly as on the server. Only the plan sheet is ever written to.
  function runTiller(id, action, p) {
    return S().load(id, { tabs: { 'Inputs': null } }).then(function (book) {
      var c0 = core(book);
      var unlink = action === 'tillerLink' && p.unlink === '1';
      var tid = unlink ? '' : c0.rbpTillerId(p.tillerId || ((c0.rbpTillerCfg(book) || {}).tillerId));
      var needTx = /^tillerSpend/.test(action);
      var tabs = needTx ? { 'Transactions': null, 'Categories': null } : { 'Accounts': null };
      var loadTiller = tid
        ? S().load(tid, { tabs: tabs }).then(function (tb) { var o = {}; o[tid] = tb; return o; },
            function (e) { if (e && e.code === 'need_signin') throw e; return null; })
        : Promise.resolve(null);
      return loadTiller.then(function (others) {
        if (tid && !others) return { error: 'tiller_access' };
        var res;
        try {
          res = core(book, others || {}).rbpTiller(action, book, book.getSheetByName('Inputs'), id, p);
        } catch (e) { res = { error: (e && e.message) || String(e) }; }
        if (res && /tiller_access/.test(String(res.error || ''))) return { error: 'tiller_access' };
        return book.flush().then(function () { readCache = {}; return res; });
      });
    });
  }

  function handle(p) {
    var id = p.sheetId, action = p.action || 'read';
    switch (action) {
      case 'read':
        return doRead(id, p.fast === '1');
      case 'save':
        return withBook(id, { 'Inputs': null }, function (book, c) {
          return c.writeInputs(book, book.getSheetByName('Inputs'), decodePayload(p));
        });
      case 'saveCheckIn':
        return withBook(id, { 'Annual Check-In': 'A1:F130' }, function (book, c) {
          return c.writeCheckIn(book, decodePayload(p));
        });
      case 'solveWriteB119':
        var v = Number(p.b119), ok = isFinite(v) && v > 0 && v <= 200000;
        return (ok ? root.RBPGoogle.valuesBatchUpdate(id, [{ range: 'Inputs!B119', values: [[v]] }]) : Promise.resolve())
          .then(function () { readCache = {}; return { ok: ok, b119: v }; });
      case 'loadStressTests': return withBook(id, { '_StressTests': null }, function (b, c) { return c.loadStressTests(b); });
      case 'saveStressTests': return withBook(id, { '_StressTests': null }, function (b, c) { return c.saveStressTests(b, decodePayload(p)); });
      case 'loadSweep':       return withBook(id, { '_StressTests': null }, function (b, c) { return c.loadSweep(b); });
      case 'saveSweep':       return withBook(id, { '_StressTests': null }, function (b, c) { return c.saveSweep(b, decodePayload(p)); });
      case 'loadMonteCarlo':  return withBook(id, { '_StressTests': null }, function (b, c) { return c.loadMonteCarlo(b); });
      case 'saveMonteCarlo':  return withBook(id, { '_StressTests': null }, function (b, c) { return c.saveMonteCarlo(b, decodePayload(p)); });
      case 'runStressTest':
      case 'runSweep':
      case 'runSweepOne':
        // Same as the connector: heal any orphaned stress state first, run, then restore.
        return S().load(id, { tabs: { 'Inputs': null, '_StressTests': null } }).then(function (book) {
          var c = core(book), inp = book.getSheetByName('Inputs');
          c.rbpHealStress(book, inp, true);
          var run = action === 'runStressTest' ? c.runStressTest : action === 'runSweep' ? c.runSweepBatch : c.runSweepOne;
          return Promise.resolve(run(book, inp, p)).then(function (res) {
            return book.flush().then(function () { readCache = {}; return res; });
          });
        });
      case 'tillerGet':
        return S().load(id, { tabs: {} }).then(function (b) { return { ok: true, cfg: core(b).rbpTillerCfg(b) }; });
      case 'tillerList': case 'tillerLink': case 'tillerPull':
      case 'tillerSpend': case 'tillerSpendLink': case 'tillerSpendApply': case 'tillerSpendGap':
        return runTiller(id, action, p);
      default:
        if (root.RBPGBridgeExtra && root.RBPGBridgeExtra[action]) return root.RBPGBridgeExtra[action](p);
        return Promise.resolve({ error: 'This feature isn’t available yet with Google sign-in (' + action + ').' });
    }
  }

  function answer(params) {
    return withSignIn(function () { return handle(params); })
      .catch(function (e) { return { error: (e && e.message) || String(e) }; });
  }

  root.RBPGBridge = {
    isGoogleMode: isGoogleMode,
    handle: function (params) { return answer(params); },
    // Switch this browser to the Google sign-in connection for the sheet the customer picked.
    useGoogle: function (sheetId) {
      try {
        localStorage.setItem(MODE_KEY, 'google');
        localStorage.setItem('rbp_sheet_id', sheetId);
        root.RBPGoogle.setSheetId(sheetId);
      } catch (e) {}
    },
    // Go back to the old connection on this browser.
    useLegacy: function () { try { localStorage.removeItem(MODE_KEY); } catch (e) {} }
  };

  // Only Google-connected browsers get the request interception installed at all.
  if (!isGoogleMode()) return;

  // ── Interception: JSONP <script src=...> ──────────────────────────────────
  function runJsonp(el, url) {
    var params = parseParams(url.split('?')[1] || '');
    var cb = params.callback;
    answer(params).then(function (data) {
      var tries = 0;
      (function deliver() {
        var fn = cb && root[cb];
        if (typeof fn === 'function') { try { fn(data); } catch (e) { setTimeout(function () { throw e; }); } return; }
        if (++tries < 50) setTimeout(deliver, 100);   // page may define its callback a moment later
      })();
    });
  }
  var desc = Object.getOwnPropertyDescriptor(HTMLScriptElement.prototype, 'src');
  if (desc && desc.set) {
    Object.defineProperty(HTMLScriptElement.prototype, 'src', {
      configurable: true, enumerable: desc.enumerable,
      get: function () { return this.__rbpSrc || desc.get.call(this); },
      set: function (v) {
        var url = String(v);
        if (url.indexOf(SCRIPT_BASE) === 0 && claims(url, parseParams(url.split('?')[1] || ''))) {
          this.__rbpSrc = url; runJsonp(this, url); return;     // answered locally — no network request
        }
        desc.set.call(this, v);
      }
    });
  }
  var origSetAttr = Element.prototype.setAttribute;
  Element.prototype.setAttribute = function (name, value) {
    if (this instanceof HTMLScriptElement && String(name).toLowerCase() === 'src') { this.src = value; return; }
    return origSetAttr.apply(this, arguments);
  };

  // ── Interception: fetch() (POST saves and the Inputs page's GET read) ─────
  var origFetch = root.fetch;
  root.fetch = function (input, init) {
    var url = typeof input === 'string' ? input : (input && input.url) || '';
    if (url.indexOf(SCRIPT_BASE) === 0 && isGoogleMode()) {
      var params = parseParams(url.split('?')[1] || '');
      if (init && init.body && typeof init.body === 'string') {
        var b = parseParams(init.body); for (var k in b) params[k] = b[k];
      }
      if (claims(url, params)) {
        return answer(params).then(function (data) {
          return new Response(JSON.stringify(data), { status: 200, headers: { 'Content-Type': 'application/json' } });
        });
      }
    }
    return origFetch.apply(this, arguments);
  };
})(window);
