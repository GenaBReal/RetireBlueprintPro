/**
 * RetireBlueprint Pro — Daily Parity Watchdog
 * ----------------------------------------------------------------------------
 * Runs once a day inside Apps Script. It compares your JS engine's projection
 * against the sheet's A8-computed projection. If they ever drift apart by more
 * than $1, it emails you the exact year(s) and amounts. If they match, it stays
 * silent (unless you flip sendAllClear to true).
 *
 * This is the automated version of the parity harness you already built — same
 * idea, just on a timer with an alert, so you can't forget to run it after
 * touching A8 or the engine.
 *
 * SETUP (one time):
 *   1) Paste this whole file into your Apps Script project (a new .gs file).
 *   2) Fill in the two functions marked  // TODO  below (your engine call and
 *      your master-sheet range). These are the only parts specific to you.
 *   3) Set masterSheetId and alertEmail in PARITY_CONFIG.
 *   4) Run installDailyParityTrigger() ONCE (it schedules the 6am daily check
 *      and will ask you to authorize MailApp the first time).
 * ----------------------------------------------------------------------------
 */

var PARITY_CONFIG = {
  toleranceDollars: 1,                              // flag any year off by more than this
  alertEmail: 'retireblueprintpro@gmail.com',       // where warnings go
  masterSheetId: 'PASTE_YOUR_PRIVATE_MASTER_SHEET_ID', // your master sheet (A8 intact)
  sendAllClear: false                               // true = also email on days everything matches
};

function runDailyParityCheck() {
  try {
    var engine = getEngineProjection_();   // engine numbers  (array, one per year)
    var sheet  = getSheetProjection_();    // sheet/A8 numbers (same shape)
    var diffs  = compareProjections_(engine, sheet, PARITY_CONFIG.toleranceDollars);

    if (diffs.length > 0) {
      sendParityAlert_(diffs);
    } else if (PARITY_CONFIG.sendAllClear) {
      MailApp.sendEmail(PARITY_CONFIG.alertEmail,
        '✅ RBP parity OK',
        'Engine and sheet matched to the dollar across all ' + engine.length + ' years today.');
    }
  } catch (err) {
    MailApp.sendEmail(PARITY_CONFIG.alertEmail,
      '⚠️ RBP parity check ERRORED',
      'The watchdog itself failed to run:\n\n' + String((err && err.stack) || err));
  }
}

/* ───────────────────────── PLUG IN YOUR SETUP (2 functions) ───────────────── */

// 1) Return the ENGINE's projected balances as a flat array of numbers, one per
//    year. Use the SAME validated engine call your parity harness already uses
//    (your rbp-engine / rbpComputeSweep function). Replace the body below.
function getEngineProjection_() {
  // TODO: call your engine and return e.g. [925000, 931200, 938050, ...]
  throw new Error('getEngineProjection_ not wired yet — paste your engine call here.');
}

// 2) Return the SHEET's A8-computed balances for the same years, from your
//    private master sheet. Adjust the tab name and range to wherever the A8
//    projection output lives (one number per row/year).
function getSheetProjection_() {
  var ss = SpreadsheetApp.openById(PARITY_CONFIG.masterSheetId);
  var rng = ss.getSheetByName('Master').getRange('B2:B41');  // TODO: adjust tab + range
  return rng.getValues()
            .map(function (r) { return Number(r[0]); })
            .filter(function (n) { return !isNaN(n); });
}

/* ───────────────────────── comparison + alert (done) ─────────────────────── */

function compareProjections_(engine, sheet, tol) {
  var diffs = [];
  if (engine.length !== sheet.length) {
    diffs.push({ note: 'Length mismatch: engine has ' + engine.length +
                       ' years, sheet has ' + sheet.length + '.' });
  }
  var n = Math.min(engine.length, sheet.length);
  for (var i = 0; i < n; i++) {
    var gap = Math.round((engine[i] - sheet[i]) * 100) / 100;
    if (Math.abs(gap) > tol) {
      diffs.push({ year: i + 1, engine: engine[i], sheet: sheet[i], gap: gap });
    }
  }
  return diffs;
}

function sendParityAlert_(diffs) {
  var lines = diffs.map(function (d) {
    return d.note
      ? '• ' + d.note
      : '• Year ' + d.year + ':   sheet $' + fmt_(d.sheet) +
        '    engine $' + fmt_(d.engine) + '    gap $' + fmt_(d.gap);
  }).join('\n');

  var body =
    'Heads up — the RetireBlueprint engine no longer matches the sheet (A8) ' +
    'within $' + PARITY_CONFIG.toleranceDollars + '.\n\n' +
    diffs.length + ' mismatch(es) found:\n\n' + lines + '\n\n' +
    'Most likely cause: A8 (or the engine) was changed and the other side ' +
    'wasn\u2019t updated to match.\n\n' +
    'Next step: re-run your parity harness, reconcile the two, and don\u2019t ship ' +
    'until this check is clean again.';

  MailApp.sendEmail(PARITY_CONFIG.alertEmail,
    '⚠️ RBP PARITY DRIFT — ' + diffs.length + ' mismatch(es)', body);
}

function fmt_(n) { return Number(n).toLocaleString('en-US'); }

/* ───────────────────────── one-time scheduler ────────────────────────────── */

// Run this ONCE to schedule the daily 6am check. Safe to re-run (it replaces
// any existing watchdog trigger rather than stacking duplicates).
function installDailyParityTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runDailyParityCheck') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyParityCheck').timeBased().everyDays(1).atHour(6).create();
}

// Optional: run this anytime to test the alert email end-to-end without waiting.
function testParityAlertEmail() {
  sendParityAlert_([{ year: 7, engine: 612345, sheet: 612300, gap: 45 }]);
}
