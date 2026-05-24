// RetireBlueprint Pro — Apps Script
// Built from ground-truth sheet audit. Every cell mapped correctly.
// NEVER write to formula cells: B33,B36,B42,B43,B45,B48,B51(formula?),
//   B86,B87,B119,B120,B128-B130,B133,B134,B140,B142,B143,B144,B146,G122-G124

var SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxCPNzU1A9jT8eiOrc-fg2kWSo9aZmSHYBFfPILZzdZD-IEhDCr0uYhGeOqUSysaCpP/exec';

function doGet(e) {
  var action   = e.parameter.action;
  var sheetId  = e.parameter.sheetId;
  var callback = e.parameter.callback; // JSONP support

  function respond(obj) {
    var jsonStr = JSON.stringify(obj);
    if (callback) {
      // JSONP — wrap in callback function
      return ContentService
        .createTextOutput(callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    // Plain JSON with CORS headers
    return ContentService
      .createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (!sheetId) return respond({error:'No sheetId'});

  try {
    var ss  = SpreadsheetApp.openById(sheetId);
    var inp = ss.getSheetByName('Inputs');
    if (!inp) return respond({error:'No Inputs sheet'});

    if (action === 'read') {
      addWelcomeBanner(ss, inp);
      return respond(readAll(ss, inp));
    }

    if (action === 'save') {
      var enc = e.parameter.enc;
      var raw = enc === 'b64'
        ? Utilities.newBlob(Utilities.base64Decode(e.parameter.data)).getDataAsString()
        : e.parameter.data;
      var data = JSON.parse(raw);
      return respond(writeInputs(ss, inp, data));
    }

    return respond({error:'Unknown action: ' + action});
  } catch(err) {
    return respond({error: err.toString()});
  }
}

// ── HELPERS ───────────────────────────────────────────────────
function doPost(e) {
  // Called by HTML fetch POST (no-cors). Can't return readable CORS response,
  // but the function executes fully and writes to the sheet.
  try {
    var params = {};
    // Parse URL-encoded body
    if (e.postData && e.postData.contents) {
      e.postData.contents.split('&').forEach(function(pair) {
        var kv = pair.split('=');
        params[decodeURIComponent(kv[0])] = decodeURIComponent((kv[1]||'').replace(/\+/g,' '));
      });
    }
    var sheetId = params.sheetId;
    if (!sheetId) return ContentService.createTextOutput('{"error":"No sheetId"}').setMimeType(ContentService.MimeType.JSON);
    var ss  = SpreadsheetApp.openById(sheetId);
    var inp = ss.getSheetByName('Inputs');
    if (!inp) return ContentService.createTextOutput('{"error":"No Inputs sheet"}').setMimeType(ContentService.MimeType.JSON);
    var enc = params.enc;
    var raw = enc === 'b64'
      ? Utilities.newBlob(Utilities.base64Decode(params.data)).getDataAsString()
      : params.data;
    var data = JSON.parse(raw);
    // Route to correct handler
    var action2 = params.action || 'save';
    if (action2 === 'saveCheckIn') {
      var result2 = writeCheckIn(ss, data);
      return ContentService.createTextOutput(JSON.stringify(result2)).setMimeType(ContentService.MimeType.JSON);
    }
    var result = writeInputs(ss, inp, data);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}


function readAll(ss, inp) {
  // Read entire sheet in ONE batch call — avoids timeout from individual getValue() calls
  var allData = inp.getDataRange().getValues();

  function r(row, col) {
    // row/col are 1-based (sheet coordinates)
    var rowArr = allData[row-1];
    if (!rowArr) return '';
    var val = rowArr[col-1];
    return (val === undefined || val === null) ? '' : val;
  }
  function pct(row, col) { return Number(r(row,col))||0; }
  function dt(row, col) {
    var val = r(row,col);
    if (!val) return '';
    try {
      var d = new Date(val);
      return isNaN(d) ? '' : (d.getMonth()+1)+'/'+d.getDate()+'/'+d.getFullYear();
    } catch(e) { return ''; }
  }
  function num(row, col) { return Number(r(row,col))||0; }
  function str(row, col) { return String(r(row,col)||''); }

  // Partner 1 (col B = col 2)
  var currentYear = new Date().getFullYear();
  var p1Age = num(33,2);
  var p1DeathAge = num(35,2);
  var p1 = {
    name:str(31,2), dob:dt(32,2), age:p1Age,
    deathAge:p1DeathAge, theme:str(34,2),
    projectedDeathYear: p1Age > 0 && p1DeathAge > 0 ? currentYear + (p1DeathAge - p1Age) : 0,
    salary:num(37,2), salaryStart:dt(38,2), salaryEnd:dt(39,2),
    ssBase:num(40,2), ssFra:num(41,2),
    ssMonthly:num(43,2), ssStartDate:dt(44,2),
    ssAnnual: num(43,2) * 12,
    ssYear1: (function() {
      // Prorate SS for the year it starts: count months from start month to Dec
      // e.g. July start (month=6) → 12-6 = 6 months
      var raw = r(44,2);
      if (!raw) return 0;
      try { var d=new Date(raw); if(isNaN(d)) return 0;
        return Math.round(num(43,2) * (12 - d.getMonth())); } catch(e) { return 0; }
    })(),
    pension:num(46,2), pensionStart:dt(47,2),
    otherIncome:num(49,2), otherStart:dt(50,2), otherEnd:dt(51,2),
    healthPreMedicare:num(52,2), healthMedicare:num(53,2), medicareAge:num(54,2),
  };

  // Partner 2 (col D = col 4)
  var p2Age = num(33,4);
  var p2DeathAge = num(35,4);
  var p2 = {
    name:str(31,4), dob:dt(32,4), age:p2Age,
    deathAge:p2DeathAge, theme:str(34,4),
    projectedDeathYear: p2Age > 0 && p2DeathAge > 0 ? currentYear + (p2DeathAge - p2Age) : 0,
    salary:num(37,4), salaryStart:dt(38,4), salaryEnd:dt(39,4),
    ssBase:num(40,4), ssFra:num(41,4),
    ssMonthly:num(43,4), ssStartDate:dt(44,4),
    ssAnnual: num(43,4) * 12,
    ssYear1: (function() {
      var raw = r(44,4);
      if (!raw) return 0;
      try { var d=new Date(raw); if(isNaN(d)) return 0;
        return Math.round(num(43,4) * (12 - d.getMonth())); } catch(e) { return 0; }
    })(),
    pension:num(46,4), pensionStart:dt(47,4),
    otherIncome:num(49,4), otherStart:dt(50,4), otherEnd:dt(51,4),
    healthPreMedicare:num(52,4), healthMedicare:num(53,4), medicareAge:num(54,4),
  };

  // Global (col B = col 2)
  var gl = {
    planStartYear:dt(6,2), planYear:(function(){ try{ return new Date(allData[5][1]).getFullYear(); } catch(e){ return new Date().getFullYear(); }})(), projectionYears:num(7,2),
    legacyGoal:num(8,2), safetyFloor:num(9,2),
    survivorReduction:pct(10,2),
    assumeExtraSpend:str(11,2),
    stressEarly:str(12,2), stressEarlyStart:num(13,2),
    stressEarlyDur:num(14,2), stressEarlyDrag:pct(15,2),
    stressLate:str(16,2), stressLateStart:num(17,2),
    stressLateDur:num(18,2), stressLateDrag:pct(19,2),
    filingStatus:str(24,2), standardDeduction:num(25,2),
  };

  // Tax
  var tax = {
    inflation:pct(20,2), ssCola:pct(21,2),
    baseTaxAdj:pct(22,2), marginalExtra:pct(23,2),
    stateRate:pct(26,2), stdDedInflation:pct(27,2),
    healthcareInflation:pct(28,2),
    rothConversionSavings:num(144,2),
    standardDeduction:num(25,2),
    // Effective federal rate read directly from Master sheet (AJ/O for year 1)
    // Populated after Master sheet is read — set as placeholder here, overwritten below
    effectiveRate: 0
  };

  // Accounts (rows 105-116, cols A-H = 1-8)
  var accounts = [], portfolioTotal = 0;
  for (var ar=105; ar<=116; ar++) {
    var aName = str(ar,2);
    if (!aName) continue;
    var bal = num(ar,5);
    var ret = pct(ar,6); // stored as decimal
    if (str(ar,1).toLowerCase()==='yes') portfolioTotal += bal;
    accounts.push({
      showInCalc:str(ar,1), name:aName, owner:str(ar,3),
      type:str(ar,4), balance:bal,
      expectedReturn:ret*100, // *100 for display
      contrib:num(ar,12),    // L = Annual Contribution
      match:num(ar,13),      // M = Employer Match
      status:str(ar,7), showOnDashboard:str(ar,8)
    });
  }

  // Expenses — A=label(col1), B=monthly(col2), D=notes(col4)
  var expRowNums = [58,59,60,61,62,63,64,65,66,67,68,69,71,72,73,74,75,76,77,78,79,80,81,83,84,85];
  var expenses = expRowNums.map(function(row) {
    var mo = num(row,2);
    return {
      row:row, name:str(row,1),
      monthly:mo, annual:mo*12, note:str(row,4)
    };
  });

  // Debts rows 92-101 (rows 88-91 are expenses/header area — debt table starts at 92)
  var debtRows = [92,93,94,95,96,97,98,99,100,101];
  var debts = debtRows.map(function(row) {
    // Smart read: detect old format (C=monthly,D=annual) vs new format (C=pp,D=monthly,E=annual)
    var colC = num(row,3), colD = num(row,4), colE = num(row,5);
    var monthly, pp;
    if (colE > 0) {
      // New format: D=monthly, E=annual
      monthly = colD; pp = colC;
    } else if (colC > 0 && colD > 0 && Math.abs(colD - colC*12) < 10) {
      // Old format: C=monthly, D=annual(=C*12)
      monthly = colC; pp = 0;
    } else if (colD > 0 && colC === 0) {
      // Ambiguous: only D has value - check if it's divisible by 12 (likely annual)
      // If D is a round multiple of 12, it was probably written as annual (old format)
      var isLikelyAnnual = colD > 1000 && (colD % 12 === 0 || colD % 100 === 0);
      monthly = isLikelyAnnual ? Math.round(colD/12) : colD;
      pp = 0;
    } else {
      monthly = colD || colC; pp = 0;
    }
    return {
      inc:str(row,1), name:str(row,2), pp:pp,
      mo:monthly, ann:monthly*12,
      start:dt(row,6), end:dt(row,7),
      bal:num(row,8),
      inflate:str(row,9),
      rate:(num(row,10)||0)*100,
      curval:num(row,11)
    };
  });

  // Spending/solver
  var baseSpend = num(87,2);
  var safeExtra = num(119,2);
  function phaseYear(row, col) {
    var val = r(row,col);
    if (!val) return null;
    try { var d=new Date(val); return isNaN(d)?null:d.getFullYear(); } catch(e){return null;}
  }
  var spending = {
    baseAnnual:baseSpend, safeExtra:safeExtra, surplus:num(120,2),
    phase1Start:phaseYear(122,3), phase1End:phaseYear(122,4),
    phase2Start:phaseYear(123,3), phase2End:phaseYear(123,4),
    phase3Start:phaseYear(124,3), phase3End:phaseYear(124,4),
    phase1Weight:num(122,5), phase2Weight:num(123,5), phase3Weight:num(124,5),
    phase1Override:num(122,6), phase2Override:num(123,6), phase3Override:num(124,6),
    phase1Extra:num(122,7), phase2Extra:num(123,7), phase3Extra:num(124,7),
  };

  // Legacy
  var endLiquid = num(133,2);
  var legacyGoal = num(8,2);
  var legacy = {
    goal:legacyGoal, safetyFloor:num(9,2),
    projectedEnding:endLiquid,
    variance:endLiquid-legacyGoal,
    achievedPct:legacyGoal>0?Math.round(endLiquid/legacyGoal*100):100,
    status:str(134,2),
  };

  // Roth
  var roth = {
    year:num(139,2), bracket:str(141,2),
    taxableIncome:num(140,2), bracketThreshold:num(142,2),
    optimalAmount:num(143,2), marginalTax:num(144,2),
    assumedRate:pct(145,2), taxSavings:num(146,2),
  };

  // Roth plan (rows 150-169, B=col2, C=col3)
  var rothPlan = [];
  for (var i=0; i<20; i++) {
    var row = 150+i;
    rothPlan.push({year:2026+i, p1:num(row,2), p2:num(row,3)});
  }

  // Read Master sheet for year-by-year projection data (charts)
  var projections = {years:[], income:[], withdrawals:[], endLiquid:[],
                     preTax:[], roth:[], taxable:[], hsa:[],
                     federalTaxes:[], taxableIncome:[]};
  var masterYear1FedTax = 0, masterYear1GrossIncome = 0, masterYear1TaxableIncome = 0;
  try {
    var masterSheet = ss.getSheetByName('Master');
    if (masterSheet) {
      var masterData = masterSheet.getRange('A8:BX200').getValues();
      var isFirstRow = true;
      masterData.forEach(function(row) {
        var yr = Number(row[0]); // col A = year
        if (!yr || yr < 2020 || yr > 2200) return;
        projections.years.push(yr);
        projections.income.push(Math.round(Number(row[14])||0));      // O  = Total Gross Guaranteed Income
        projections.withdrawals.push(Math.round(Number(row[39])||0)); // AN = Final Gross Withdrawal
        projections.endLiquid.push(Math.round(Number(row[75])||0));   // BX = Total End of Year Liquid Assets
        projections.federalTaxes.push(Math.round(Number(row[35])||0));  // AJ = Total Federal Taxes
        projections.taxableIncome.push(Math.round(Number(row[34])||0)); // AI = Total Taxable Income
        // Capture year-1 values for effective rate calculation
        if (isFirstRow) {
          masterYear1FedTax       = Number(row[35])||0; // AJ = Total Federal Taxes
          masterYear1GrossIncome  = Number(row[14])||0; // O  = Total Gross Guaranteed Income
          masterYear1TaxableIncome= Number(row[34])||0; // AI = Total Taxable Income
          isFirstRow = false;
        }
        // Account balances
        var p1_401k = Number(row[54])||0; // BC
        var p1_roth = Number(row[55])||0; // BD
        var p1_brok = Number(row[56])||0; // BE
        var p2_401k = Number(row[57])||0; // BF
        var p2_roth = Number(row[58])||0; // BG
        var p2_brok = Number(row[59])||0; // BH
        var hsa     = Number(row[60])||0; // BI
        projections.preTax.push(Math.round(p1_401k + p2_401k));
        projections.roth.push(Math.round(p1_roth + p2_roth));
        projections.taxable.push(Math.round(p1_brok + p2_brok));
        projections.hsa.push(Math.round(hsa));
      });
    }
  } catch(e) {
    Logger.log('Master read error: ' + e);
  }

  // Set effective federal rate from Master sheet year-1 data
  if (masterYear1GrossIncome > 0) {
    tax.effectiveRate = Math.round((masterYear1FedTax / masterYear1TaxableIncome) * 10000) / 10000; // AJ÷AI: fed taxes / taxable income = standard effective rate
    tax.taxableIncome = masterYear1TaxableIncome;
    tax.federalTaxes  = masterYear1FedTax;
  }

  // Read Annual Check-In data from existing 'Annual Check-In' tab
  // Cols: A=Year, B=Projected, C=Actual, D=Variance, E=Status, F=Notes
  // Data starts at row 8
  var checkInData = [];
  try {
    var ciSheet = ss.getSheetByName('Annual Check-In');
    if (ciSheet) {
      var ciLastRow = ciSheet.getLastRow();
      if (ciLastRow >= 8) {
        var ciData = ciSheet.getRange(8, 1, ciLastRow-7, 6).getValues();
        ciData.forEach(function(ciRow) {
          if (ciRow[0]) {
            checkInData.push({
              year:      Number(ciRow[0]),
              projected: Number(ciRow[1])||0,   // B = Projected (Master formula)
              actual:    ciRow[2]!==''?Number(ciRow[2]):null, // C = Actual (we write)
              variance:  ciRow[3]!==''?Number(ciRow[3]):null, // D = Variance (formula)
              status:    String(ciRow[4]||''),                // E = Status (formula)
              notes:     String(ciRow[5]||'')                 // F = Notes (we write)
            });
          }
        });
      }
    }
  } catch(e) { Logger.log('CheckIn read error: '+e); }

  return {
    meta:{planYear:new Date().getFullYear()},
    people:{craig:p1, gena:p2},
    checkIn:checkInData,
    global:gl, tax:tax,
    portfolio:{accounts:accounts, total:portfolioTotal},
    expenses:expenses, debts:debts,
    spending:spending, legacy:legacy,
    roth:roth, rothPlan:rothPlan,
    projections:projections,
  };
}


// ── WRITE ────────────────────────────────────────────────────
function writeInputs(ss, inp, data) {
  // Helper: set a cell value (string)
  function set(cell, val) {
    if (val!==undefined && val!==null && val!=='') inp.getRange(cell).setValue(String(val));
  }
  // Helper: set a numeric value
  function setN(cell, val) {
    var n = Number(val);
    if (!isNaN(n)) inp.getRange(cell).setValue(n);
  }
  // Helper: set a percent — val is whole number (7), store as decimal (0.07)
  function setPct(cell, val) {
    var n = Number(val);
    if (!isNaN(n)) inp.getRange(cell).setValue(n/100);
  }
  // Helper: set a date from YYYY-MM-DD or M/D/YYYY string
  function setD(cell, val) {
    if (!val || val==='' || val==='0') return;
    try {
      var s = String(val).trim();
      var yr, mo, dy;
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        var p=s.split('-'); yr=parseInt(p[0]); mo=parseInt(p[1])-1; dy=parseInt(p[2]);
      } else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
        var p=s.split('/'); yr=parseInt(p[2]); mo=parseInt(p[0])-1; dy=parseInt(p[1]);
      } else { return; }
      if (yr>1900 && yr<2200) inp.getRange(cell).setValue(new Date(yr,mo,dy,12,0,0));
    } catch(e) { Logger.log('setD error '+cell+': '+e); }
  }

  try {
    // ── GLOBAL SETTINGS ──────────────────────────────────────
    if (data.global) {
      var g = data.global;
      // B6 = Plan Start Year — write as Jan 2 to avoid timezone rollback
      if (g.B6) {
        var yr = parseInt(String(g.B6).replace(/[^0-9]/g,''));
        if (!yr || yr < 2000) yr = new Date().getFullYear();
        inp.getRange('B6').setValue(new Date(yr, 0, 2, 12, 0, 0)); // Jan 2, noon
      }
      if (g.B7)  setN('B7', g.B7);
      if (g.B8)  setN('B8', g.B8);
      if (g.B9)  setN('B9', g.B9);
      if (g.B10 !== undefined) setPct('B10', g.B10); // whole number input → /100
      if (g.B11) set('B11', g.B11);
      if (g.B12) set('B12', g.B12);
      if (g.B13) setN('B13', g.B13);
      if (g.B14) setN('B14', g.B14);
      if (g.B15 !== undefined) setPct('B15', g.B15);
      if (g.B16) set('B16', g.B16);
      if (g.B17) setN('B17', g.B17);
      if (g.B18) setN('B18', g.B18);
      if (g.B19 !== undefined) setPct('B19', g.B19);
      if (g.B20 !== undefined) setPct('B20', g.B20);
      if (g.B21 !== undefined) setPct('B21', g.B21);
      if (g.B22 !== undefined) setPct('B22', g.B22);
      if (g.B23 !== undefined) setPct('B23', g.B23);
      if (g.B24) set('B24', g.B24);
      if (g.B25) setN('B25', g.B25);
      if (g.B26 !== undefined) setPct('B26', g.B26);
      if (g.B27 !== undefined) setPct('B27', g.B27);
      if (g.B28 !== undefined) setPct('B28', g.B28);
      // NEVER WRITE: B33, B36
    }

    // ── PARTNER 1 ─────────────────────────────────────────────
    if (data.partner1) {
      var p = data.partner1;
      if (p.B31) set('B31', p.B31);
      if (p.B32) setD('B32', p.B32);
      if (p.B34) set('B34', p.B34);
      if (p.B35) setN('B35', p.B35);
      if (p.B37) setN('B37', p.B37);
      if (p.B38) setD('B38', p.B38);
      if (p.B39) setD('B39', p.B39);
      if (p.B40) setN('B40', p.B40);
      if (p.B41) setN('B41', p.B41);
      // NEVER WRITE: B33,B36,B42,B43,B45,B48
      if (p.B44) setD('B44', p.B44);
      if (p.B46) setN('B46', p.B46);
      if (p.B47) setD('B47', p.B47);
      if (p.B49) setN('B49', p.B49);
      if (p.B50) setD('B50', p.B50);
      if (p.B51) setD('B51', p.B51);
      if (p.B52) setN('B52', p.B52);
      if (p.B53) setN('B53', p.B53);
      if (p.B54) setN('B54', p.B54);
    }

    // ── PARTNER 2 ─────────────────────────────────────────────
    if (data.partner2) {
      var p = data.partner2;
      if (p.D31) set('D31', p.D31);
      if (p.D32) setD('D32', p.D32);
      if (p.D34) set('D34', p.D34);
      if (p.D35) setN('D35', p.D35);
      if (p.D37) setN('D37', p.D37);
      if (p.D38) setD('D38', p.D38);
      if (p.D39) setD('D39', p.D39);
      if (p.D40) setN('D40', p.D40);
      if (p.D41) setN('D41', p.D41);
      // NEVER WRITE: D33,D36,D42,D43,D45,D48
      if (p.D44) setD('D44', p.D44);
      if (p.D46) setN('D46', p.D46);
      if (p.D47) setD('D47', p.D47);
      if (p.D49) setN('D49', p.D49);
      if (p.D50) setD('D50', p.D50);
      if (p.D51) setD('D51', p.D51);
      if (p.D52) setN('D52', p.D52);
      if (p.D53) setN('D53', p.D53);
      if (p.D54) setN('D54', p.D54);
    }

    // ── EXPENSES ──────────────────────────────────────────────
    // A=editable label, B=monthly($), C=yearly(FORMULA-NEVER WRITE), D=notes
    if (data.expenses && data.expenses.length) {
      var skip = {70:true, 82:true, 86:true, 87:true}; // 70 and 82 are section headers; 86,87 are SUM rows
      data.expenses.forEach(function(exp) {
        var r = Number(exp.row);
        if (!r || r<58 || r>85 || skip[r]) return;
        if (exp.name && String(exp.name).trim() !== '') inp.getRange('A'+r).setValue(String(exp.name).trim());
        var mo = Number(exp.monthly)||0;
        inp.getRange('B'+r).setValue(mo);
        inp.getRange('C'+r).setValue(mo * 12); // yearly = monthly * 12
        if (exp.note !== undefined) inp.getRange('D'+r).setValue(String(exp.note||''));
      });
    }

    // ── DEBTS ─────────────────────────────────────────────────
    // A=Include, B=Name, C=PurchasePrice, D=Monthly, E=Annual, F=Start, G=End, H=Balance
    // Clear rows 88-89 in case old code wrote debt data there by mistake
    ['A88','B88','C88','D88','E88','F88','G88','H88','I88','J88','K88',
     'A89','B89','C89','D89','E89','F89','G89','H89','I89','J89','K89'].forEach(function(cell){
      try { inp.getRange(cell).clearContent(); } catch(e){}
    });
    if (data.debts && data.debts.length) {
      var debtRows = [92,93,94,95,96,97,98,99,100,101];
      data.debts.forEach(function(d, i) {
        if (i >= debtRows.length) return;
        var r = debtRows[i];
        var mo = Number(d.mo)||0;
        // Only write if row has meaningful data (name OR monthly > 0)
        // This prevents wiping existing sheet data with blank web rows
        var hasData = (d.name && String(d.name).trim() !== '') || mo > 0 || Number(d.bal) > 0;
        if (!hasData && i >= 2) return; // Skip empty extra rows (keep sheet data intact)

        // A=Include, B=Name, C=PurchasePrice, D=Monthly, E=Annual, F=Start, G=End, H=Balance
        inp.getRange('A'+r).setValue(d.inc||'No');
        if (d.name && String(d.name).trim()) inp.getRange('B'+r).setValue(String(d.name));
        if (d.pp) inp.getRange('C'+r).setValue(Number(d.pp));
        if (mo > 0 || hasData) {
          inp.getRange('D'+r).setValue(mo);      // D = Monthly
          inp.getRange('E'+r).setValue(mo*12);   // E = Annual (Master reads this!)
        }
        if (d.start) setD('F'+r, d.start);
        if (d.end)   setD('G'+r, d.end);
        var bal = Number(d.bal)||0;
        if (bal) inp.getRange('H'+r).setValue(bal);
        if (d.inflate) inp.getRange('I'+r).setValue(String(d.inflate));
        if (d.rate) inp.getRange('J'+r).setValue(Number(d.rate)/100);
        if (d.curval) inp.getRange('K'+r).setValue(Number(d.curval));
      });
    }

    // ── ACCOUNTS ──────────────────────────────────────────────
    // A=include, B=name, C=owner, D=type, E=balance, F=return(decimal), G=status, H=dashboard
    if (data.accounts && data.accounts.length) {
      var p1name = data.partner1 && data.partner1.B31 ? String(data.partner1.B31).trim() : '';
      var p2name = data.partner2 && data.partner2.D31 ? String(data.partner2.D31).trim() : '';

      // Clear owner validation temporarily
      try { inp.getRange('C105:C116').clearDataValidations(); } catch(e){}

      var acctData = [];
      for (var i=0; i<12; i++) {
        var a = data.accounts[i] || {};
        var owner = String(a.owner||'Joint').trim();
        if (owner==='Partner 1'||owner==='Craig'||owner==='Kent') owner = p1name||'Joint';
        if (owner==='Partner 2'||owner==='Gena'||owner==='Georgia') owner = p2name||'Joint';
        if (!owner || owner==='Joint') owner = 'Joint';
        // ret comes in as whole number (7.0), store as decimal (0.07)
        var ret = Number(a.ret||a.expectedReturn||0);
        if (ret > 1) ret = ret/100;
        acctData.push([
          a.inc||a.showInCalc||'No',  // A
          a.name||'',                  // B
          owner,                       // C
          a.type||'',                  // D
          Number(a.bal||a.balance)||0, // E = balance
          ret,                         // F = return as decimal
          a.status||'Use for Withdrawals', // G
          a.dash||a.showOnDashboard||'No', // H
          '', '', '',                    // I, J, K (unused)
          Number(a.contrib||0),            // L = Annual Contribution
          Number(a.match||0)               // M = Employer Match
        ]);
      }
      inp.getRange('A105:M116').setValues(acctData);

      // Restore owner validation
      try {
        var tsr = ss.getSheetByName('Technical Style Reference');
        if (tsr) {
          var rule = SpreadsheetApp.newDataValidation()
            .requireValueInRange(tsr.getRange('F6:F8'), true)
            .setAllowInvalid(false).build();
          inp.getRange('C105:C116').setDataValidation(rule);
        }
      } catch(e){}
    }

    // ── SMILE CURVE ───────────────────────────────────────────
    // C122-D124 = Date objects, E122-E124 = weights, F122-F124 = overrides
    // G122-G124 = FORMULA — NEVER WRITE
    if (data.phases) {
      var ph = data.phases;
      function yearToDate(yr) {
        var n = parseInt(yr);
        return (n > 2000 && n < 2200) ? new Date(n, 0, 1) : null;
      }
      var d; // reuse var
      if (ph.C122) { d=yearToDate(ph.C122); if(d) inp.getRange('C122').setValue(d); }
      if (ph.D122) { d=yearToDate(ph.D122); if(d) inp.getRange('D122').setValue(d); }
      if (ph.C123) { d=yearToDate(ph.C123); if(d) inp.getRange('C123').setValue(d); }
      if (ph.D123) { d=yearToDate(ph.D123); if(d) inp.getRange('D123').setValue(d); }
      if (ph.C124) { d=yearToDate(ph.C124); if(d) inp.getRange('C124').setValue(d); }
      if (ph.D124) { d=yearToDate(ph.D124); if(d) inp.getRange('D124').setValue(d); }
      if (ph.E122!==undefined) inp.getRange('E122').setValue(Number(ph.E122)||1.15);
      if (ph.F122!==undefined) inp.getRange('F122').setValue(Number(ph.F122)||0);
      if (ph.E123!==undefined) inp.getRange('E123').setValue(Number(ph.E123)||1.0);
      if (ph.F123!==undefined) inp.getRange('F123').setValue(Number(ph.F123)||0);
      if (ph.E124!==undefined) inp.getRange('E124').setValue(Number(ph.E124)||0.9);
      if (ph.F124!==undefined) inp.getRange('F124').setValue(Number(ph.F124)||0);
    }

    // ── ROTH SOLVER SETTINGS ─────────────────────────────────
    // B139=year(writable), B141=bracket(writable), B145=rate(writable as decimal)
    // NEVER WRITE: B140,B142,B143,B144,B146
    if (data.roth) {
      var yr = Number(data.roth.B139);
      if (yr > 2020) setN('B139', yr);
      if (data.roth.B141) set('B141', data.roth.B141);
      var rate = Number(data.roth.B145);
      if (rate > 0) setPct('B145', rate); // whole number in, /100 stored
    }

    // ── ROTH CONVERSION PLAN ──────────────────────────────────
    // B150:B169 = partner1 amounts, C150:C169 = partner2 amounts (rows = 2026-2045)
    if (data.rothPlan && data.rothPlan.length) {
      // Write year, p1 amount, p2 amount to A:C rows 150-169
      // Master VLOOKUP uses col A for year lookup: VLOOKUP(year, A136:C157, 2, FALSE)
      var allArr = [];
      for (var i=0; i<20; i++) {
        var yr = 2026 + i;
        var p1 = 0, p2 = 0;
        data.rothPlan.forEach(function(row) {
          if (Number(row.year) === yr) { p1=Number(row.p1)||0; p2=Number(row.p2)||0; }
        });
        allArr.push([yr, p1, p2]); // A=year, B=p1, C=p2
      }
      inp.getRange('A150:C169').setValues(allArr);
    }

    SpreadsheetApp.flush();

    // Hide and protect all sheets except Master after every save
    try { setupAllSheets(ss); } catch(e) { Logger.log('Sheet setup skipped: '+e); }

    return {success:true, timestamp:new Date().toISOString()};
  } catch(err) {
    return {success:false, error:err.toString()};
  }
}

// ── WELCOME MESSAGE ───────────────────────────────────────────────────────────
// Called once when a new customer first connects their sheet
// Adds a friendly banner so they know the blank sheet is normal
function addWelcomeBanner(ss, inp) {
  try {
    // Check if welcome has already been shown
    var props = PropertiesService.getDocumentProperties();
    if (props.getProperty('welcomeShown')) return;
    
    // Add a note to the sheet letting them know it's blank on purpose
    var note = inp.getRange('C5');
    note.setValue('👋 Welcome! This sheet will fill in automatically after you enter your data on the Inputs page of your dashboard and click Save All Changes. The blank cells are completely normal — nothing is missing!');
    note.setFontColor('#1CC7D0');
    note.setFontWeight('bold');
    note.setFontSize(11);
    
    props.setProperty('welcomeShown', 'true');
  } catch(e) {
    Logger.log('Welcome banner error: ' + e);
  }
}

// ── ANNUAL CHECK-IN WRITE ─────────────────────────────────────────────
// Writes ONLY to the existing "Annual Check-In" sheet tab
// Col A = Year (formula/static — never write)
// Col B = Projected balance (Master formula — NEVER WRITE)
// Col C = Actual End-of-Year Balance ← WE WRITE THIS
// Col D = Variance (formula — NEVER WRITE)
// Col E = Review Status (formula — NEVER WRITE)
// Col F = Notes / Life Events ← WE WRITE THIS
// Data rows start at row 8

// ── SHEET VISIBILITY + FORMULA PROTECTION SETUP ──────────────────────
// 1. Hides all tabs except Master (hidden state copies to customer sheets)
// 2. Adds WARNING-ONLY range protection on formula cells in Inputs
//    — scripts bypass warnings entirely so Apps Script writes freely
//    — customers see "Are you sure?" if they unhide and try to edit formulas
//    — data-entry cells are left completely open (no warning, no block)
// 3. Annual Check-In: warning on everything except C and F (actuals/notes)
// Safe to run multiple times — clears old range protections before adding new ones
function setupAllSheets(ss) {
  try {
    var sheets = ss.getSheets();
    sheets.forEach(function(sheet) {
      var name = sheet.getName();

      // Master stays visible and untouched
      if (name === 'Master') return;

      // Hide every other sheet
      try { sheet.hideSheet(); } catch(e) {}

      // Clear any existing RANGE protections on this sheet
      var rangeProt = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      rangeProt.forEach(function(p) { try { p.remove(); } catch(e){} });

      if (name === 'Inputs') {
        // Protect formula/structural cells with WARNING only
        // Scripts bypass warnings — customers see a caution dialog
        // NEVER-WRITE formula cells (from handoff doc):
        var formulaRanges = [
          'B33','B36','B42','B43','B45','B48',   // Partner formula cells
          'B86','B87',                             // Expense SUM rows
          'B119','B120',                           // Solver cells
          'B128:B130','B133','B134',               // Legacy cells
          'B140','B142:B144','B146',               // Roth formula cells
          'G122:G124',                             // Smile curve formulas
          'C58:C85',                               // Expense yearly = monthly*12 (formula)
        ];
        formulaRanges.forEach(function(r) {
          try {
            var p = sheet.getRange(r).protect()
              .setDescription('Formula — do not edit directly')
              .setWarningOnly(true); // Warning dialog, not a hard block — scripts bypass
          } catch(e) {}
        });

      } else if (name === 'Annual Check-In') {
        // Protect everything EXCEPT col C (actuals) and col F (notes) with warning
        // Rows 8+ are data rows
        var ciWarningRanges = ['A8:B200', 'D8:E200', 'G8:ZZ200'];
        ciWarningRanges.forEach(function(r) {
          try {
            sheet.getRange(r).protect()
              .setDescription('Calculated — do not edit directly')
              .setWarningOnly(true);
          } catch(e) {}
        });

      } else {
        // All other hidden sheets — warning on everything
        try {
          sheet.getDataRange().protect()
            .setDescription(name + ' — do not edit directly')
            .setWarningOnly(true);
        } catch(e) {}
      }
    });
  } catch(e) {
    Logger.log('setupAllSheets error: ' + e);
  }
}

function setupCheckInSheet(ss) {
  // Just hides the Annual Check-In tab — no protection needed
  // setupAllSheets() handles hiding all non-Master tabs anyway
  try {
    var ciSheet = ss.getSheetByName('Annual Check-In');
    if (ciSheet) ciSheet.hideSheet();
  } catch(e) {
    Logger.log('setupCheckInSheet error: ' + e);
  }
}

function writeCheckIn(ss, data) {
  try {
    var checkIns = data.checkIn;
    if (!checkIns || !checkIns.length) return {status:'ok', message:'No data'};

    var ciSheet = ss.getSheetByName('Annual Check-In');
    if (!ciSheet) return {status:'error', error:'Annual Check-In sheet tab not found'};

    // Ensure sheet is hidden and protected on every write (idempotent — safe to repeat)
    try { setupCheckInSheet(ss); } catch(e) { Logger.log('Setup skipped: '+e); }

    var written = 0;
    checkIns.forEach(function(row) {
      if (!row.year) return;
      // Find the row by matching year in col A (rows 8+)
      var startRow = 8;
      var lastRow = ciSheet.getLastRow();
      for (var r = startRow; r <= lastRow; r++) {
        var yr = ciSheet.getRange(r, 1).getValue();
        if (Number(yr) === Number(row.year)) {
          // Write actual to col C only if provided
          if (row.actual !== null && row.actual !== undefined && row.actual !== '') {
            ciSheet.getRange(r, 3).setValue(Number(row.actual));
          } else if (row.actual === null || row.actual === '') {
            ciSheet.getRange(r, 3).clearContent();
          }
          // Write notes to col F
          ciSheet.getRange(r, 6).setValue(String(row.notes||''));
          written++;
          break;
        }
      }
    });

    return {status:'ok', message:'CheckIn saved', written: written};
  } catch(e) {
    return {status:'error', error: e.toString()};
  }
}
