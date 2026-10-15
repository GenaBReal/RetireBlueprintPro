/* RetireBlueprint Pro — Google Analytics + Etsy button tracking.
   One file, loaded by every public page. Does not touch the planner or any numbers.
   To change the Google Analytics ID, edit GA_ID below only. */
(function () {
  var GA_ID = 'G-YP10J1Z3DR';
  if (!GA_ID || GA_ID.indexOf('XXXX') !== -1) return;           // not set up yet: do nothing
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
