/* ============================================================================
   RetireBlueprint Pro — SAMPLE PLAN DEMO MODE  (rbp-demo.js)
   ----------------------------------------------------------------------------
   WHAT THIS DOES
   Lets shoppers open the real app on a real sample plan (Michael & Linda)
   without being able to change it.

   - Turned ON by visiting any app page with  ?demo=1  in the address
     (the homepage "Explore the sample plan" button does this).
     It stays on for that browser tab until the visitor taps "Exit demo".
   - While ON, every page reads the DEMO sheet instead of the visitor's own
     sheet. A real customer's saved connection is never touched or erased.
   - Scrolling, charts, tabs and what-if sliders all work normally.
   - Changing a plan input, or tapping any Save / Apply button, shows the
     "Get your own plan" pop-up instead.
   - Any request that would write to the sheet is stopped here. The connector
     also refuses writes to the demo sheet, so it stays safe even if someone
     bypasses this file.

   HOW IT IS WIRED IN
   One line, the FIRST script in the <head> of each app page:
       <script src="rbp-demo.js"></script>
   With no ?demo=1 it does nothing at all, so customers never see it.
   ============================================================================ */
(function () {
  'use strict';

  /* ---- Settings ---- */
  var DEMO_SHEET_ID = '1eWe0KMgVXUsTpELyv-612Qa6ijXlLjFEl8PSuyQEn20';
  var BUY_URL       = 'https://www.etsy.com/listing/4544156492';
  var NAMES         = { rbp_p1_name: 'Michael', rbp_p2_name: 'Linda' };
  var FLAG          = 'rbp_demo';

  /* Connector actions that only READ. Everything else is a write and is blocked. */
  var READ_ACTIONS = { read: 1, loadSweep: 1, loadStressTests: 1, loadMonteCarlo: 1, runStressTest: 1, runSweep: 1, runSweepOne: 1, tillerGet: 1 };

  /* Pages where every form field is a saved plan input */
  var INPUT_PAGES = /RetireBlueprintPro_(Inputs|AnnualCheckIn)\.html/i;

  /* ---- Turn on / off ---- */
  var ss;
  try { ss = window.sessionStorage; } catch (e) { ss = null; }
  function getFlag() { try { return ss && ss.getItem(FLAG) === '1'; } catch (e) { return false; } }
  function setFlag(on) { try { if (!ss) return; on ? ss.setItem(FLAG, '1') : ss.removeItem(FLAG); } catch (e) {} }

  var params = new URLSearchParams(location.search);
  if (params.get('demo') === '1') setFlag(true);
  if (params.get('demo') === '0') setFlag(false);
  if (!getFlag()) return;                     // not in demo: do nothing
  if (DEMO_SHEET_ID.indexOf('PASTE_') === 0) { setFlag(false); return; }  // not configured yet

  window.RBP_DEMO = true;
  var inFrame = false;
  try { inFrame = window.self !== window.top; } catch (e) { inFrame = true; }

  /* ---- 1. Point every page at the demo sheet, without touching the visitor's own saved link ---- */
  var proto = Storage.prototype;
  var _get = proto.getItem, _set = proto.setItem, _remove = proto.removeItem;
  var GUARDED = { rbp_sheet_id: DEMO_SHEET_ID, rbp_p1_name: NAMES.rbp_p1_name, rbp_p2_name: NAMES.rbp_p2_name };

  proto.getItem = function (k) {
    if (this === window.localStorage && Object.prototype.hasOwnProperty.call(GUARDED, k)) return GUARDED[k];
    return _get.apply(this, arguments);
  };
  proto.setItem = function (k) {
    if (this === window.localStorage && Object.prototype.hasOwnProperty.call(GUARDED, k)) return;
    return _set.apply(this, arguments);
  };
  proto.removeItem = function (k) {
    if (this === window.localStorage && Object.prototype.hasOwnProperty.call(GUARDED, k)) return;
    return _remove.apply(this, arguments);
  };

  /* ---- 2. Stop anything that would write to the sheet ---- */
  function actionOf(text) {
    var m = /(?:^|[?&])action=([A-Za-z0-9_]+)/.exec(text || '');
    if (m) return m[1];
    try { var j = JSON.parse(text); if (j && j.action) return String(j.action); } catch (e) {}
    return null;
  }
  function isConnector(url) { return /script\.google(usercontent)?\.com/.test(String(url || '')); }
  function isWrite(action) { return action && !READ_ACTIONS[action]; }

  var _fetch = window.fetch;
  if (_fetch) {
    window.fetch = function (input, init) {
      var url = typeof input === 'string' ? input : (input && input.url) || '';
      if (isConnector(url)) {
        var body = init && typeof init.body === 'string' ? init.body : '';
        var act = actionOf(body) || actionOf(url.split('?')[1] || '');
        if (isWrite(act) || (init && /post/i.test(init.method || '') && !READ_ACTIONS[act])) {
          showBuy('save');
          return Promise.reject(new Error('Sample plan: changes are turned off.'));
        }
      }
      return _fetch.apply(this, arguments);
    };
  }

  var srcDesc = Object.getOwnPropertyDescriptor(HTMLScriptElement.prototype, 'src');
  function blockedSrc(v) {
    if (!isConnector(v)) return null;
    var act = actionOf(String(v).split('?')[1] || '');
    if (!isWrite(act)) return null;
    var cb = /[?&]callback=([A-Za-z0-9_$.]+)/.exec(v);
    showBuy('save');
    var js = cb ? cb[1] + '({error:"Sample plan: changes are turned off."});' : '';
    return 'data:text/javascript,' + encodeURIComponent(js);
  }
  if (srcDesc && srcDesc.set) {
    Object.defineProperty(HTMLScriptElement.prototype, 'src', {
      configurable: true, enumerable: srcDesc.enumerable,
      get: function () { return srcDesc.get.call(this); },
      set: function (v) { var b = blockedSrc(v); srcDesc.set.call(this, b || v); }
    });
  }
  var _setAttr = Element.prototype.setAttribute;
  Element.prototype.setAttribute = function (name, value) {
    if (this instanceof HTMLScriptElement && String(name).toLowerCase() === 'src') {
      var b = blockedSrc(value); if (b) value = b;
    }
    return _setAttr.call(this, name, value);
  };

  /* ---- 3. Pop-up when someone tries to change the plan ---- */
  var SAVE_WORDS = /\b(save|apply|link|connect|import|reset|delete)\b/i;
  var onInputPage = INPUT_PAGES.test(location.pathname);

  document.addEventListener('click', function (e) {
    var el = e.target && e.target.closest ? e.target.closest('button, a, input[type=button], input[type=submit], [role=button]') : null;
    if (!el || el.closest('#rbpDemoBar, #rbpDemoModal')) return;
    var label = (el.textContent || el.value || el.getAttribute('aria-label') || el.title || '').trim();
    var isSave = el.id === 'save-btn' || /save/i.test(el.className || '') || SAVE_WORDS.test(label);
    var isSetup = el.tagName === 'A' && /setup\.html/i.test(el.getAttribute('href') || '');
    if (isSave || isSetup) { e.preventDefault(); e.stopImmediatePropagation(); showBuy('save'); }
  }, true);

  if (onInputPage) {
    var guard = function (e) {
      var f = e.target;
      if (!f || !f.closest || f.closest('#rbpDemoBar, #rbpDemoModal')) return;
      if (/^(INPUT|SELECT|TEXTAREA)$/.test(f.tagName) || f.isContentEditable) {
        e.preventDefault(); e.stopImmediatePropagation();
        if (f.blur) f.blur();
        showBuy('input');
      }
    };
    document.addEventListener('mousedown', guard, true);
    document.addEventListener('touchstart', guard, { capture: true, passive: false });
    document.addEventListener('focusin', guard, true);
    document.addEventListener('keydown', function (e) {
      var f = e.target; if (f && /^(INPUT|SELECT|TEXTAREA)$/.test(f.tagName) && !f.closest('#rbpDemoModal')) { e.preventDefault(); showBuy('input'); }
    }, true);
  }

  /* ---- 4. The sample-plan bar and the pop-up ---- */
  var css =
    '#rbpDemoBar{position:fixed;left:0;right:0;bottom:0;z-index:2147483000;background:#041E2F;border-top:2px solid #f4b86a;' +
    'color:#fff;font:600 14px/1.35 Montserrat,system-ui,-apple-system,"Segoe UI",sans-serif;' +
    'padding:10px 16px calc(10px + env(safe-area-inset-bottom,0px));display:flex;gap:10px 16px;align-items:center;justify-content:center;flex-wrap:wrap;text-align:center}' +
    '#rbpDemoBar b{color:#1CC7D0}' +
    '#rbpDemoBar a{color:#041E2F;background:#f4b86a;border-radius:999px;padding:8px 16px;text-decoration:none;font-weight:800;white-space:nowrap}' +
    '#rbpDemoBar button{background:none;border:1px solid rgba(255,255,255,.35);color:#cfe0e8;border-radius:999px;padding:7px 14px;font:inherit;cursor:pointer}' +
    'body.rbp-demo{padding-bottom:84px!important}' +
    '#rbpDemoModal{position:fixed;inset:0;z-index:2147483001;background:rgba(2,14,24,.72);display:flex;align-items:center;justify-content:center;padding:16px}' +
    '#rbpDemoModal[hidden]{display:none!important}' +
    '#rbpDemoModal .box{background:#fff;color:#14232E;max-width:420px;width:100%;border-radius:18px;overflow:hidden;box-shadow:0 30px 80px rgba(0,0,0,.4);font:16px/1.5 "Source Sans 3",system-ui,-apple-system,"Segoe UI",sans-serif}' +
    '#rbpDemoModal .top{background:#041E2F;color:#fff;padding:22px 24px}' +
    '#rbpDemoModal h2{margin:0;font:800 22px/1.2 Montserrat,system-ui,sans-serif}' +
    '#rbpDemoModal .top p{margin:8px 0 0;color:#c9d9e1;font-size:15px}' +
    '#rbpDemoModal .body{padding:20px 24px 24px;display:grid;gap:12px}' +
    '#rbpDemoModal .price{font:800 34px/1 Montserrat,system-ui,sans-serif;color:#9A6420}' +
    '#rbpDemoModal .price small{font-size:14px;color:#4A5D69;font-weight:600;margin-left:6px}' +
    '#rbpDemoModal .buy{display:block;text-align:center;background:#f4b86a;color:#041E2F;font:800 16px Montserrat,system-ui,sans-serif;text-decoration:none;border-radius:999px;padding:14px}' +
    '#rbpDemoModal .stay{background:none;border:0;color:#0E7C83;font:700 15px Montserrat,system-ui,sans-serif;cursor:pointer;padding:6px}' +
    '#rbpDemoModal .fine{font-size:13px;color:#4A5D69;text-align:center}';

  var MESSAGES = {
    input: ['This is Michael & Linda\'s sample plan', 'Want to see your own numbers here? Get your own private copy and enter your real ages, savings and Social Security.'],
    save:  ['Saving is for your own plan', 'This sample plan is read-only so every visitor sees the same example. Your own copy saves everything you change.']
  };

  var modal;
  function build() {
    if (document.getElementById('rbpDemoBar')) return;
    var st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);

    var bar = document.createElement('div');
    bar.id = 'rbpDemoBar';
    bar.innerHTML = '<span>You\'re exploring <b>Michael &amp; Linda\'s sample plan</b>. Look around, nothing here is real money.</span>' +
      '<a href="' + BUY_URL + '" target="_blank" rel="noopener">Get your own plan · $149</a>' +
      '<button type="button" id="rbpDemoExit">Exit sample</button>';
    document.body.appendChild(bar);
    document.body.classList.add('rbp-demo');

    modal = document.createElement('div');
    modal.id = 'rbpDemoModal';
    modal.hidden = true;
    modal.setAttribute('role', 'dialog');
    modal.setAttribute('aria-modal', 'true');
    modal.setAttribute('aria-labelledby', 'rbpDemoTitle');
    modal.innerHTML = '<div class="box"><div class="top"><h2 id="rbpDemoTitle"></h2><p id="rbpDemoMsg"></p></div>' +
      '<div class="body"><div class="price">$149<small>one time · no subscription</small></div>' +
      '<a class="buy" href="' + BUY_URL + '" target="_blank" rel="noopener">Get RetireBlueprint Pro on Etsy</a>' +
      '<button type="button" class="stay" id="rbpDemoStay">Keep exploring the sample</button>' +
      '<div class="fine">Instant digital delivery. Your plan lives in your own Google Sheet.</div></div></div>';
    document.body.appendChild(modal);

    document.getElementById('rbpDemoStay').addEventListener('click', hide);
    modal.addEventListener('click', function (e) { if (e.target === modal) hide(); });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') hide(); });
    document.getElementById('rbpDemoExit').addEventListener('click', function () {
      setFlag(false);
      location.href = '/';
    });
  }
  var lastShown = 0;
  function showBuy(kind) {
    if (inFrame) {                                     // Stress Test tools run in frames: use the page's pop-up
      try { if (window.parent.RBP_DEMO_SHOW) { window.parent.RBP_DEMO_SHOW(kind); return; } } catch (e) {}
    }
    if (!modal) { if (document.body) build(); else return; }
    if (!modal.hidden || Date.now() - lastShown < 600) return;
    var m = MESSAGES[kind] || MESSAGES.save;
    document.getElementById('rbpDemoTitle').textContent = m[0];
    document.getElementById('rbpDemoMsg').textContent = m[1];
    modal.hidden = false;
    lastShown = Date.now();
    var stay = document.getElementById('rbpDemoStay'); if (stay) stay.focus();
  }
  function hide() { if (modal) { modal.hidden = true; lastShown = Date.now(); } }

  if (!inFrame) {
    window.RBP_DEMO_SHOW = showBuy;
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', build);
    else build();
  }
})();
