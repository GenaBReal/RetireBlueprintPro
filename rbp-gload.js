/* RetireBlueprint Pro — loads the Google sign-in bridge ONLY for customers who connected
   with "Sign in with Google". Everyone else: this does nothing and loads nothing extra.
   Must be the first script in <head> so the bridge is ready before the page talks to its sheet. */
(function () {
  try {
    if (localStorage.getItem('rbp_conn') === 'google' && localStorage.getItem('rbp_g_sheet_id')) {
      var v = '?v=20261017b';
      ['rbp-google.js', 'rbp-gs-shim.js', 'rbp-gs-core.js', 'rbp-gbridge.js'].forEach(function (f) {
        document.write('<script src="/' + f + v + '"><\/script>');
      });
    }
  } catch (e) {}
})();
