/* RetireBlueprint Pro — Google Analytics + Etsy button tracking.
   One file, loaded by every public page. Does not touch the planner or any numbers.
   To change the Google Analytics ID, edit GA_ID below only. */
(function () {
  var GA_ID = 'G-YP10J1Z3DR';
  if (!GA_ID || GA_ID.indexOf('XXXX') !== -1) return;           // not set up yet: do nothing

  // "Don't count me" switch for Gena's own devices.
  // Open any page with ?rbp_notrack=1 once on a device to stop counting that browser;
  // ?rbp_notrack=0 turns counting back on. Stored only in that browser.
  var KEY = 'rbp_notrack', m = /[?&]rbp_notrack=([01])\b/.exec(location.search), off = false;
  try {
    if (m) { if (m[1] === '1') localStorage.setItem(KEY, '1'); else localStorage.removeItem(KEY); }
    off = localStorage.getItem(KEY) === '1';
  } catch (e) {}
  if (m) {
    var note = function () {
      var d = document.createElement('div');
      d.textContent = off ? '\u2713 This browser will no longer be counted in website visits.'
                          : '\u2713 This browser is counted in website visits again.';
      d.style.cssText = 'position:fixed;left:50%;bottom:24px;transform:translateX(-50%);z-index:99999;' +
        'background:#041E2F;color:#fff;border:2px solid #1CC7D0;border-radius:10px;padding:12px 18px;' +
        'font:15px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,.3);';
      document.body.appendChild(d);
      setTimeout(function () { d.remove(); }, 6000);
    };
    if (document.body) note(); else document.addEventListener('DOMContentLoaded', note);
  }
  if (off) { window['ga-disable-' + GA_ID] = true; return; }      // not counted: load nothing
  var s = document.createElement('script');
  s.async = true;
  s.src = 'https://www.googletagmanager.com/gtag/js?id=' + GA_ID;
  document.head.appendChild(s);
  window.dataLayer = window.dataLayer || [];
  window.gtag = function () { dataLayer.push(arguments); };
  gtag('js', new Date());
  gtag('config', GA_ID, {
    anonymize_ip: true,
    // never send query strings that could hold anything private; keep only demo=1
    page_location: location.origin + location.pathname + (/[?&]demo=1/.test(location.search) ? '?demo=1' : '')
  });

  // Etsy buttons: every click on any link to etsy.com becomes an "etsy_click" event,
  // including buttons added later by the sample-plan popup.
  document.addEventListener('click', function (e) {
    var a = e.target && e.target.closest ? e.target.closest('a[href*="etsy.com"]') : null;
    if (!a) return;
    var where = a.getAttribute('data-placement') ||
      (a.closest('header') ? 'header' : a.closest('footer') ? 'footer' :
       a.closest('.rbp-demo-modal, .rbp-demo-bar, [class*="demo"]') ? 'sample-plan-popup' : 'page');
    gtag('event', 'etsy_click', {
      link_text: (a.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 80),
      placement: where,
      page_path: location.pathname,
      transport_type: 'beacon'
    });
  }, true);
})();
