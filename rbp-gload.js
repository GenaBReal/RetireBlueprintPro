/* RetireBlueprint Pro — page loader (first script in <head>).
   1) Loads the Google sign-in bridge ONLY for customers who connected with "Connect Privately
      with Google". Everyone else: loads nothing extra.
   2) For visitors with no plan connected (people looking around), shows a friendly banner
      pointing them to buy on Etsy or try the free demo. Buyers can tap "Connect" instead. */
(function () {
  var ETSY = 'https://www.etsy.com/listing/4544156492';
  var conn = '', gid = '', sid = '';
  try {
    conn = localStorage.getItem('rbp_conn') || '';
    gid = localStorage.getItem('rbp_g_sheet_id') || '';
    sid = localStorage.getItem('rbp_sheet_id') || '';
  } catch (e) {}

  try {
    if (conn === 'google' && gid) {
      var v = '?v=20261017c';
      ['rbp-google.js', 'rbp-gs-shim.js', 'rbp-gs-core.js', 'rbp-gbridge.js'].forEach(function (f) {
        document.write('<script src="/' + f + v + '"><\/script>');
      });
    }
  } catch (e) {}

  // ── Visitor banner ──────────────────────────────────────────────────────────
  var OLD_TEMPLATE = '1D_vUyoSusb7iPX1z4c3jOv5TVGpFJqo2YZfvHfHOonA';
  var noPlan = !sid || sid === OLD_TEMPLATE;
  var inDemo = /[?&]demo=1\b/.test(location.search);
  try { if (sessionStorage.getItem('rbp_demo') === '1' && !/[?&]demo=0\b/.test(location.search)) inDemo = true; } catch (e) {}
  if (!noPlan || inDemo || /Demo\.html$/i.test(location.pathname)) return;
  function show() {
    if (document.getElementById('rbp-visitor')) return;
    var d = document.createElement('div');
    d.id = 'rbp-visitor';
    d.setAttribute('role', 'region');
    d.setAttribute('aria-label', 'Get RetireBlueprint Pro');
    d.style.cssText = 'background:#0c2438;border-bottom:2px solid #f4b86a;color:#E8F1F5;padding:12px 16px;' +
      'font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;text-align:center;position:relative;z-index:50;';
    var btn = 'display:inline-block;margin:6px 4px 0;padding:8px 14px;border-radius:6px;font-weight:700;font-size:13px;text-decoration:none;';
    d.innerHTML =
      '<strong style="color:#f4b86a;">You’re looking around RetireBlueprint Pro.</strong> ' +
      'These pages fill in with your own numbers once you have your plan.<br>' +
      '<a href="' + ETSY + '" target="_blank" rel="noopener" style="' + btn + 'background:#1CC7D0;color:#041E2F;">Get RetireBlueprint Pro on Etsy</a>' +
      '<a href="/RetireBlueprintPro_Dashboard.html?demo=1" style="' + btn + 'border:1px solid #1CC7D0;color:#1CC7D0;">Try the Free Demo</a>' +
      '<a href="/setup.html" style="' + btn + 'color:#9FB6C3;font-weight:400;text-decoration:underline;">Already bought? Connect your plan</a>';
    document.body.insertBefore(d, document.body.firstChild);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', show);
  else show();
})();
