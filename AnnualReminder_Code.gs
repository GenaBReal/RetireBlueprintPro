/****************************************************************************
 * RetireBlueprint Pro — Annual Tax & Medicare Update Reminder
 * --------------------------------------------------------------------------
 * Emails you a checklist + official source links each year when the new
 * government numbers are released, so you can keep the tool current:
 *   • mid-October  → IRS brackets / standard deduction / LTCG + Social Security COLA
 *   • mid-November → Medicare Part B premium + IRMAA (CMS)
 *
 * SETUP (one time):
 *   1. In your Master sheet: Extensions → Apps Script.
 *   2. Paste this file in (as a new script file or alongside your existing code).
 *   3. Run  setupAnnualReminders  once (authorize when prompted).
 *   4. Optional: run  sendTestReminders  to email yourself both reminders now,
 *      just to confirm it works.
 *
 * That's it. Nothing else runs automatically except the two yearly emails.
 *
 * BONUS — helping existing customers:
 *   After you finish your annual updates, run  emailCustomerUpdateNote  once.
 *   It reads this year's numbers from your Tax Engine tab and emails you a
 *   friendly, ready-to-send message (pre-filled) that you can paste to your
 *   buyers (e.g. an Etsy Message to Buyers) so they can refresh their own copy.
 ****************************************************************************/

var REMINDER_EMAIL = 'gbphilpott@gmail.com';
var TAX_TAB_NAME   = 'Tax Engine';   // change this only if your tab has a different name
var CHEATSHEET_URL = 'https://genabreal.github.io/RetireBlueprintPro/RetireBlueprintPro_TaxUpdate_CheatSheet.pdf';

/**
 * Run this ONCE to install the schedule. Safe to re-run (it clears old copies first).
 */
function setupAnnualReminders() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'monthlyReminderCheck') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Fires on the 20th of every month ~9am; the function below only acts in Oct & Nov.
  ScriptApp.newTrigger('monthlyReminderCheck')
    .timeBased()
    .onMonthDay(20)
    .atHour(9)
    .create();
  Logger.log('Reminders installed → ' + REMINDER_EMAIL +
             '. IRS reminder each October, Medicare reminder each November.');
}

/**
 * Trigger target. Sends the right reminder depending on the month.
 */
function monthlyReminderCheck() {
  var month = new Date().getMonth(); // 0=Jan ... 9=Oct, 10=Nov
  if (month === 9) {
    sendIrsReminder();
  } else if (month === 10) {
    sendCmsReminder();
  }
}

/** Run manually any time to preview BOTH emails immediately. */
function sendTestReminders() {
  sendIrsReminder();
  sendCmsReminder();
}

/* ---------------------------------------------------------------------- */

function sendIrsReminder() {
  var nextYear = new Date().getFullYear() + 1;
  var subject = 'RetireBlueprint Pro — Update IRS tax tables for ' + nextYear;
  var att = getCheatSheetAttachment_();
  var html =
    '<div style="font-family:Segoe UI,Arial,sans-serif;color:#1f2937;max-width:640px;line-height:1.5">' +
    '<h2 style="color:#0A3145;margin-bottom:4px">Annual IRS update — ' + nextYear + '</h2>' +
    '<p style="color:#5f6368;margin-top:0">The IRS releases next year\'s inflation-adjusted numbers in mid-to-late October. ' +
    'Verify each one against the official source, then update these cells in your Master sheet.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Federal income tax brackets (MFJ)</h3>' +
    '<p style="margin-top:0">Paste the new <b>lower-bound thresholds</b> into <b>Tax&nbsp;Engine!C8:C14</b> ' +
    '(rates in D8:D14 almost never change). Remember C8 stays <b>0</b>.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Standard deduction (MFJ)</h3>' +
    '<p style="margin-top:0">Update <b>Inputs!B25</b> to the new MFJ standard deduction. ' +
    'The engine halves it automatically for single/survivor years.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Social Security COLA</h3>' +
    '<p style="margin-top:0">Update the <b>Social Security COLA %</b> field on the Inputs tab with the announced COLA.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Long-term capital gains brackets</h3>' +
    '<p style="margin-top:0">Update the displayed 0% / 15% / 20% thresholds on your <b>Tax Engine</b> and ' +
    '<b>Roth Ladder</b> web pages so the reference tables stay current.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Update the website (one file)</h3>' +
    '<p style="margin-top:0">Open <b>rbp-tax-config.js</b> in your GitHub repo, edit the bracket and standard-deduction ' +
    'numbers in the \u201cEDIT THESE NUMBERS\u201d block, and Commit. Every web page updates automatically \u2014 ' +
    'no page-by-page edits, and customers never reconnect.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Help existing customers (optional)</h3>' +
    '<p style="margin-top:0">Their website tools update automatically, but the projection inside each customer\u2019s own ' +
    'sheet keeps the year they bought with. To offer them a refresh: finish the updates above, then run the ' +
    '<b>emailCustomerUpdateNote</b> function once \u2014 it emails you a ready-to-send note (pre-filled with this ' +
    'year\u2019s numbers) you can paste into an Etsy Message to Buyers.</p>' +

    '<p style="margin-top:0"><b>No change needed:</b> RMD start age (Tax&nbsp;Engine!F32/F33) is a formula tied to date of birth — it updates itself.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Official sources</h3>' +
    '<ul>' +
    '<li>IRS inflation adjustments (Revenue Procedure): <a href="https://www.irs.gov/newsroom">irs.gov/newsroom</a> — search "inflation adjustments ' + nextYear + '"</li>' +
    '<li>Social Security COLA: <a href="https://www.ssa.gov/cola/">ssa.gov/cola</a></li>' +
    '</ul>' +
    (att.length ? '<p style="margin-top:0;color:#0A3145"><b>Your one-page cheat sheet is attached</b> to this email for quick reference.</p>' : '') +
    '<p style="color:#9ca3af;font-size:12px">Sent automatically by your RetireBlueprint Pro reminder script.</p>' +
    '</div>';

  MailApp.sendEmail({ to: REMINDER_EMAIL, subject: subject, htmlBody: html, attachments: att });
}

function sendCmsReminder() {
  var nextYear = new Date().getFullYear() + 1;
  var subject = 'RetireBlueprint Pro — Update Medicare / IRMAA for ' + nextYear;
  var att = getCheatSheetAttachment_();
  var html =
    '<div style="font-family:Segoe UI,Arial,sans-serif;color:#1f2937;max-width:640px;line-height:1.5">' +
    '<h2 style="color:#0A3145;margin-bottom:4px">Annual Medicare update — ' + nextYear + '</h2>' +
    '<p style="color:#5f6368;margin-top:0">CMS announces next year\'s Part B premium and IRMAA brackets in early-to-mid November. ' +
    'Verify against the official source, then update these cells.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Part B base premium</h3>' +
    '<p style="margin-top:0">Reflect the new monthly base premium in your post-Medicare healthcare assumption ' +
    '<b>Inputs!B53</b> and <b>Inputs!D53</b>.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">IRMAA income thresholds</h3>' +
    '<p style="margin-top:0">Single filers → <b>Tax&nbsp;Engine!F8:F13</b>; married filing jointly → <b>G8:G13</b>. ' +
    'Use the <b>first dollar that triggers each tier</b> (e.g. 218001, not 218000).</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">IRMAA surcharges (monthly)</h3>' +
    '<p style="margin-top:0">Part B surcharge → <b>Tax&nbsp;Engine!H8:H13</b>; Part D surcharge → <b>I8:I13</b>. ' +
    'These are the surcharge amounts only — the base premium lives in Inputs!B53/D53.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Update the website (one file)</h3>' +
    '<p style="margin-top:0">Open <b>rbp-tax-config.js</b> in your GitHub repo, update the <b>irmaa</b> section ' +
    '(Part&nbsp;B base premium + the tier thresholds and surcharges) in the \u201cEDIT THESE NUMBERS\u201d block, and Commit. ' +
    'Every web page updates automatically.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Help existing customers (optional)</h3>' +
    '<p style="margin-top:0">After you finish updating, run the <b>emailCustomerUpdateNote</b> function once to get a ' +
    'ready-to-send note (pre-filled with this year\u2019s numbers) you can paste to your buyers.</p>' +

    '<h3 style="color:#0A3145;margin-bottom:4px">Official sources</h3>' +
    '<ul>' +
    '<li>Medicare costs (Part B + IRMAA): <a href="https://www.medicare.gov/basics/costs/medicare-costs">medicare.gov/basics/costs</a></li>' +
    '<li>CMS announcements: <a href="https://www.cms.gov/newsroom">cms.gov/newsroom</a></li>' +
    '</ul>' +
    (att.length ? '<p style="margin-top:0;color:#0A3145"><b>Your one-page cheat sheet is attached</b> to this email for quick reference.</p>' : '') +
    '<p style="color:#9ca3af;font-size:12px">Sent automatically by your RetireBlueprint Pro reminder script.</p>' +
    '</div>';

  MailApp.sendEmail({ to: REMINDER_EMAIL, subject: subject, htmlBody: html, attachments: att });
}

/* ======================================================================
 * CUSTOMER UPDATE NOTE  (run manually, after your annual updates)
 * Reads this year's numbers from your Tax Engine tab and emails YOU a
 * ready-to-send, non-technical message you can paste to your buyers.
 * If it can't read the tab, it still sends a blank template you can fill.
 * ==================================================================== */
function emailCustomerUpdateNote() {
  var year = new Date().getFullYear() + 1;   // the numbers you're entering are for next year
  var brk = null, irmaa = null, cg = null;
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TAX_TAB_NAME);
    if (sh) {
      brk   = sh.getRange('C8:C14').getValues().map(function(r){ return r[0]; });
      irmaa = sh.getRange('F8:I13').getValues();           // [single, couple, partB, partD] x6
      cg    = sh.getRange('L8:L10').getValues().map(function(r){ return r[0]; });
    }
  } catch (e) { /* fall through to blank template */ }

  var plain = buildCustomerNote_(year, brk, irmaa, cg);
  var inner = buildCustomerNoteHtml_(year, brk, irmaa, cg);
  var wrap =
    '<div style="font-family:Segoe UI,Arial,sans-serif;color:#1f2937;max-width:660px;line-height:1.5">' +
    '<h2 style="color:#0A3145;margin-bottom:4px">Ready-to-send customer note \u2014 ' + year + '</h2>' +
    '<p style="color:#5f6368;margin-top:0">Copy the message below and paste it to your buyers (e.g. Etsy \u2192 ' +
    'Shop&nbsp;Manager \u2192 Message&nbsp;to&nbsp;Buyers, or any channel you use). It\u2019s written for a ' +
    'non-technical reader and framed as optional. ' +
    (brk ? 'The numbers are pre-filled from your Tax Engine tab.' :
           '\u26a0\ufe0f Could not read your Tax Engine tab, so the numbers are blank \u2014 fill them in, or check that the tab is named \u201c' + TAX_TAB_NAME + '\u201d.') +
    '</p>' +
    '<div style="border:1px solid #d4dde3;border-radius:8px;padding:14px 16px;background:#f7fafb">' + inner + '</div>' +
    '<p style="color:#9ca3af;font-size:12px;margin-top:14px">A plain-text version is in this email body too, ' +
    'in case you\u2019d rather paste without formatting.</p></div>';

  MailApp.sendEmail({
    to: REMINDER_EMAIL,
    subject: 'RetireBlueprint Pro \u2014 customer update note for ' + year + ' (ready to send)',
    htmlBody: wrap,
    body: plain
  });
}

function fmtMoney_(v) {
  if (v === '' || v === null || v === undefined) return '___';
  var n = Math.round(Number(v));
  if (isNaN(n)) return String(v);
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function buildCustomerNote_(year, brk, irmaa, cg) {
  var L = [];
  L.push('Keeping your RetireBlueprint Pro current for ' + year + ' (optional, ~a few minutes)');
  L.push('');
  L.push('Hi! Each year the IRS and Medicare release new tax numbers. Good news: the calculators and');
  L.push('reference tables on the RetireBlueprint Pro website update automatically -- you do not need to');
  L.push('do anything for those.');
  L.push('');
  L.push('If you would like your personal long-range projection to use the latest figures too, you can');
  L.push('update a few cells on the "Tax Engine" tab inside your own copy of the sheet. This is optional');
  L.push('-- a one-year change barely moves a multi-decade estimate -- but here is how, if you like:');
  L.push('');
  L.push('1) Open your RetireBlueprint Pro Google Sheet and click the "Tax Engine" tab at the bottom.');
  L.push('');
  L.push('MOST IMPORTANT -- Tax bracket thresholds (the "Threshold" column, cells C8 down to C14):');
  L.push('   ' + (brk ? brk.map(fmtMoney_).join(',  ') : 'C8=0,  C9=___,  C10=___,  C11=___,  C12=___,  C13=___,  C14=___'));
  L.push('');
  L.push('OPTIONAL -- if you would like to be thorough:');
  L.push('  Medicare/IRMAA (Section IV). Type each column straight down, starting at row 8:');
  if (irmaa) {
    L.push('    Individual MAGI (F8:F13):   ' + irmaa.map(function(r){ return fmtMoney_(r[0]); }).join(',  '));
    L.push('    Couple MAGI     (G8:G13):   ' + irmaa.map(function(r){ return fmtMoney_(r[1]); }).join(',  '));
    L.push('    Part B surcharge (H8:H13):  ' + irmaa.map(function(r){ return fmtMoney_(r[2]); }).join(',  '));
    L.push('    Part D surcharge (I8:I13):  ' + irmaa.map(function(r){ return fmtMoney_(r[3]); }).join(',  '));
  } else {
    L.push('    Individual MAGI (F8:F13), Couple MAGI (G8:G13), Part B (H8:H13), Part D (I8:I13)');
  }
  L.push('  Capital-gains thresholds (Section V, cells L8 down to L10):');
  L.push('    ' + (cg ? cg.map(fmtMoney_).join(',  ') : 'L8=0,  L9=___,  L10=___'));
  L.push('');
  L.push('That is it -- the sheet saves automatically. If anything looks off, just undo (Ctrl/Cmd + Z).');
  L.push('Questions? Reply any time and I am happy to help.');
  return L.join('\n');
}

function buildCustomerNoteHtml_(year, brk, irmaa, cg) {
  function pre(s) {
    return '<pre style="font-family:Consolas,Menlo,monospace;font-size:12px;background:#eef3f5;' +
           'padding:8px 10px;border-radius:6px;white-space:pre-wrap;margin:6px 0">' + s + '</pre>';
  }
  var brkLine = brk ? brk.map(fmtMoney_).join(',  ')
                    : 'C8=0,  C9=___,  C10=___,  C11=___,  C12=___,  C13=___,  C14=___';
  var ir;
  if (irmaa) {
    ir = 'Individual MAGI (F8:F13):   ' + irmaa.map(function(r){ return fmtMoney_(r[0]); }).join(',  ') + '\n' +
         'Couple MAGI     (G8:G13):   ' + irmaa.map(function(r){ return fmtMoney_(r[1]); }).join(',  ') + '\n' +
         'Part B surcharge (H8:H13):  ' + irmaa.map(function(r){ return fmtMoney_(r[2]); }).join(',  ') + '\n' +
         'Part D surcharge (I8:I13):  ' + irmaa.map(function(r){ return fmtMoney_(r[3]); }).join(',  ');
  } else {
    ir = 'Individual MAGI (F8:F13), Couple MAGI (G8:G13), Part B (H8:H13), Part D (I8:I13)';
  }
  var cgLine = cg ? cg.map(fmtMoney_).join(',  ') : 'L8=0,  L9=___,  L10=___';

  var h = '';
  h += '<p style="margin:0 0 8px"><b>Keeping your RetireBlueprint Pro current for ' + year + '</b> (optional, ~a few minutes)</p>';
  h += '<p style="margin:0 0 8px">Hi! Each year the IRS and Medicare release new tax numbers. The calculators and reference tables on the RetireBlueprint&nbsp;Pro website update <b>automatically</b> &mdash; nothing to do there.</p>';
  h += '<p style="margin:0 0 8px">If you\u2019d like your personal long-range projection to use the latest figures too, you can update a few cells on the <b>\u201cTax Engine\u201d tab</b> inside your own copy of the sheet. This is entirely optional &mdash; a one-year change barely moves a multi-decade estimate.</p>';
  h += '<p style="margin:0 0 4px"><b>1)</b> Open your sheet and click the <b>\u201cTax Engine\u201d</b> tab at the bottom.</p>';
  h += '<p style="margin:10px 0 2px"><b>Most important &mdash; tax bracket thresholds</b> (the \u201cThreshold\u201d column, cells <b>C8\u2192C14</b>):</p>' + pre(brkLine);
  h += '<p style="margin:10px 0 2px"><b>Optional, if you\u2019d like to be thorough:</b></p>';
  h += '<p style="margin:4px 0 2px">Medicare / IRMAA (Section IV) &mdash; type each column straight down from row 8:</p>' + pre(ir);
  h += '<p style="margin:4px 0 2px">Capital-gains thresholds (Section V, cells L8:L10):</p>' + pre(cgLine);
  h += '<p style="margin:10px 0 0">That\u2019s it &mdash; the sheet saves automatically. If anything looks off, just undo (Ctrl/Cmd&nbsp;+&nbsp;Z). Questions? Reply any time.</p>';
  return h;
}

/* ----------------------------------------------------------------------
 * Fetches your cheat-sheet PDF from your website so it can be attached to
 * the reminder emails. Upload RetireBlueprintPro_TaxUpdate_CheatSheet.pdf
 * to your GitHub repo (same place as your images) and it will be picked up
 * automatically. If it isn't there yet, the email simply sends without it.
 * -------------------------------------------------------------------- */
function getCheatSheetAttachment_() {
  try {
    var resp = UrlFetchApp.fetch(CHEATSHEET_URL, { muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      return [ resp.getBlob().setName('RetireBlueprintPro_TaxUpdate_CheatSheet.pdf') ];
    }
  } catch (e) { /* unreachable / not uploaded yet — send without attachment */ }
  return [];
}
