/* ─────────────────────────────────────────────────────────────────────────────
   RetireBlueprint Pro — Tiller bridge (optional)
   The customer keeps their OWN Tiller subscription. Tiller (via Plaid) holds the bank /
   403(b) logins and refreshes the customer's own Tiller Google Sheet. The customer shares
   that sheet with retireblueprintpro@gmail.com, and the connector READS the latest balances
   from it (never writes to it). This file only talks to our connector.

   window.RBPTiller:
     list(tillerUrlOrId, cb)       → {ok, tillerId, accounts:[{id,name,num,institution,type,balance,lastUpdate}]}
     link(tillerUrlOrId, map, cb)  → save matching {tillerAccountId: planSlot 0-11} in the plan sheet
     unlink(cb)
     pull(cb)                      → write any changed balances into the plan → {ok, linked, updated:[], asOf}
     refreshPlan(cb)               → after a pull: fresh read, re-solve safe spending, store it → cb(freshData)
     autoPull(data, onUpdated)     → Dashboard: pull at most every 6 hours per browser; if anything
                                     changed, re-solve and hand back fresh data
   ───────────────────────────────────────────────────────────────────────────── */
(function (root) {
  var SCRIPT = 'https://script.google.com/macros/s/AKfycbxR40oJlUa6p8sCANeoHNfXU5XCnlwOudWfkgKWgorlzrXiqXAT8NbPuJU1x0QbfUg7/exec';

  function sheetId() { try { return localStorage.getItem('rbp_sheet_id') || ''; } catch (e) { return ''; } }

  function jsonp(params, cb, timeoutMs) {
    var name = 'rbpTiller_' + Date.now() + '_' + Math.floor(Math.random() * 1e6), done = false;
    var t = setTimeout(function () { if (!done) { done = true; cleanup(); cb({ error: 'timeout' }); } }, timeoutMs || 90000);
    function cleanup() { try { delete root[name]; } catch (e) { root[name] = undefined; } if (s && s.parentNode) s.parentNode.removeChild(s); }
    root[name] = function (d) { if (done) return; done = true; clearTimeout(t); cleanup(); cb(d || {}); };
    var q = ['authuser=0', 'sheetId=' + encodeURIComponent(sheetId())];
    Object.keys(params).forEach(function (k) { q.push(k + '=' + encodeURIComponent(params[k])); });
    var s = document.createElement('script');
    s.crossOrigin = 'anonymous';
    s.src = SCRIPT + '?' + q.join('&') + '&callback=' + name + '&_=' + Date.now();
    s.onerror = function () { if (!done) { done = true; clearTimeout(t); cleanup(); cb({ error: 'network' }); } };
    document.head.appendChild(s);
  }

  function tillerIdFrom(x) { var m = String(x || '').match(/[-\w]{25,}/); return m ? m[0] : ''; }

  function money(n) { n = Math.round(Number(n) || 0); return '$' + n.toLocaleString(); }

  var api = {
    tillerIdFrom: tillerIdFrom,
    money: money,

    list: function (tiller, cb) {
      var id = tillerIdFrom(tiller);
      if (!id) { cb({ error: 'Paste the full link to your Tiller sheet.' }); return; }
      jsonp({ action: 'tillerList', tillerId: id }, cb);
    },

    link: function (tiller, map, cb) {
      jsonp({ action: 'tillerLink', tillerId: tillerIdFrom(tiller), map: JSON.stringify(map || {}) }, cb);
    },

    unlink: function (cb) { jsonp({ action: 'tillerLink', unlink: '1' }, cb); },

    pull: function (cb) { jsonp({ action: 'tillerPull' }, cb, 120000); },

    // Fresh read → re-solve the safe-extra with the shared engine → store it (B119) the same
    // lean way the Inputs page does after a save. cb(freshData or null).
    refreshPlan: function (cb) {
      jsonp({ action: 'read' }, function (d) {
        if (!d || d.error) { cb(null); return; }
        try {
          var sp = d.spending || {};
          var allPinned = (+sp.phase1Override > 0 && +sp.phase2Override > 0 && +sp.phase3Override > 0);
          if (!allPinned && root.RBP && RBP.buildI && RBP.solveSafeExtra) {
            var res = RBP.solveSafeExtra(RBP.buildI(d));
            var solved = res && Number(res.safeExtra);
            if (isFinite(solved) && solved > 0 && solved <= 200000 && Math.abs(solved - (Number(sp.safeExtra) || 0)) >= 0.5) {
              var w1 = Number(sp.phase1Weight) || 1, w2 = Number(sp.phase2Weight) || 1, w3 = Number(sp.phase3Weight) || 1, avg = (w1 + w2 + w3) / 3 || 1;
              d.spending.safeExtra = solved;
              d.spending.phase1Extra = solved * w1 / avg;
              d.spending.phase2Extra = solved * w2 / avg;
              d.spending.phase3Extra = solved * w3 / avg;
              jsonp({ action: 'solveWriteB119', b119: solved }, function () { cb(d); }, 30000);
              return;
            }
          }
        } catch (e) {}
        cb(d);
      }, 120000);
    },

    // Dashboard helper. onUpdated(freshData, result) fires only when a balance actually changed.
    autoPull: function (data, onUpdated) {
      try {
        var cfg = data && data.tiller;
        if (!cfg || !cfg.tillerId || !cfg.map || !Object.keys(cfg.map).length) return;
        var key = 'rbp_tiller_pulled_' + sheetId(), last = 0;
        try { last = Number(localStorage.getItem(key)) || 0; } catch (e) {}
        if (Date.now() - last < 6 * 3600 * 1000) return; // at most every 6 hours per browser
        try { localStorage.setItem(key, String(Date.now())); } catch (e) {}
        api.pull(function (res) {
          if (!res || res.error || !res.updated || !res.updated.length) return;
          api.refreshPlan(function (fresh) { if (fresh) onUpdated(fresh, res); });
        });
      } catch (e) {}
    },

    // Plain-English one-liner for a pull result.
    describe: function (res) {
      if (!res || !res.updated || !res.updated.length) return 'Your balances already match Tiller.';
      return 'Updated from Tiller' + (res.asOf ? ' (as of ' + res.asOf + ')' : '') + ': ' +
        res.updated.map(function (u) { return (u.account || 'Account') + ' ' + money(u.from) + ' → ' + money(u.to); }).join(' · ');
    }
  };

  root.RBPTiller = api;
})(typeof window !== 'undefined' ? window : this);
