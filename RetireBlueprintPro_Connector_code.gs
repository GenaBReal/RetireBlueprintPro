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

    if (action === 'runStressTest') {
      // Write bear-market params -> recalc -> read stressed results -> RESTORE originals.
      return respond(runStressTest(ss, inp, e.parameter));
    }

    if (action === 'loadStressTests') {
      // Return the saved Test 1-5 slots from the hidden _StressTests tab.
      return respond(loadStressTests(ss));
    }

    if (action === 'saveStressTests') {
      // Persist the Test 1-5 slots to the hidden _StressTests tab.
      var stRaw = e.parameter.enc === 'b64'
        ? Utilities.newBlob(Utilities.base64Decode(e.parameter.data)).getDataAsString()
        : e.parameter.data;
      return respond(saveStressTests(ss, JSON.parse(stRaw)));
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


function doPost(e) {
  try {
    var params = {};
    if (e.postData && e.postData.contents) {
      e.postData.contents.split('&').forEach(function(pair) {
        var kv = pair.split('=');
        if (kv.length >= 2) {
          params[decodeURIComponent(kv[0])] = decodeURIComponent(kv.slice(1).join('=').replace(/\+/g,' '));
        }
      });
    }
    var sheetId = params.sheetId;
    if (!sheetId) return ContentService.createTextOutput(JSON.stringify({error:'No sheetId'})).setMimeType(ContentService.MimeType.JSON);
    var ss  = SpreadsheetApp.openById(sheetId);
    var inp = ss.getSheetByName('Inputs');
    if (!inp) return ContentService.createTextOutput(JSON.stringify({error:'No Inputs sheet'})).setMimeType(ContentService.MimeType.JSON);
    var enc = params.enc;
    var raw = enc === 'b64'
      ? Utilities.newBlob(Utilities.base64Decode(params.data)).getDataAsString()
      : params.data;
    var data = JSON.parse(raw);
    var action = params.action || 'save';
    if (action === 'saveCheckIn') {
      var result2 = writeCheckIn(ss, data);
      return ContentService.createTextOutput(JSON.stringify(result2)).setMimeType(ContentService.MimeType.JSON);
    }
    var result = writeInputs(ss, inp, data);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── HELPERS ───────────────────────────────────────────────────
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
    ssAnnual:(function(){
      var mo=num(43,2); return mo*12;
    })(),
    ssYear1:(function(){
      // Prorate SS for start year based on start month
      var mo=num(43,2); if(!mo) return 0;
      var startDt=dt(44,2); if(!startDt) return mo*12;
      var parts=String(startDt).split('/');
      if(parts.length<3) return mo*12;
      var planYr=(function(){ try{ return new Date(allData[5][1]).getFullYear(); } catch(e){ return new Date().getFullYear(); }})();
      var ssYr=parseInt(parts[2]);
      if(ssYr>planYr) return 0; // SS hasn't started yet in plan year
      if(ssYr<planYr) return mo*12; // Full year
      // Same year — prorate by months remaining
      var ssMo=parseInt(parts[0]);
      var monthsInYear=13-ssMo; // e.g. July start = 6 months (Jul-Dec)
      return Math.round(mo*monthsInYear);
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
    ssAnnual:(function(){
      var mo=num(43,4); return mo*12;
    })(),
    ssYear1:(function(){
      var mo=num(43,4); if(!mo) return 0;
      var startDt=dt(44,4); if(!startDt) return mo*12;
      var parts=String(startDt).split('/');
      if(parts.length<3) return mo*12;
      var planYr=(function(){ try{ return new Date(allData[5][1]).getFullYear(); } catch(e){ return new Date().getFullYear(); }})();
      var ssYr=parseInt(parts[2]);
      if(ssYr>planYr) return 0;
      if(ssYr<planYr) return mo*12;
      var ssMo=parseInt(parts[0]);
      var monthsInYear=13-ssMo;
      return Math.round(mo*monthsInYear);
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
    effectiveRate: 0  // set below from Master sheet data
  };

  // Accounts (rows 105-116, cols A-H = 1-8)
  var accounts = [], portfolioTotal = 0;
  var acctSlots = []; // FIXED 12-slot order (rows 105-116) to align with per-account withdrawals
  for (var ar=105; ar<=116; ar++) {
    var aNameSlot = str(ar,2);
    acctSlots.push({ name:aNameSlot, owner:str(ar,3), type:str(ar,4), included:(str(ar,1).toLowerCase()==='yes') });
  }
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
      contrib: (function(){ var v=num(ar,10); return (v>0 && v<500000) ? v : 0; })(),  // J = Annual Contribution
      match:   (function(){ var v=num(ar,11); return (v>0 && v<500000) ? v : 0; })(),  // K = Employer Match
      contribStart:dt(ar,12),  // L = Contrib Start Date
      contribEnd:dt(ar,13),    // M = Contrib End Date
      withdrawStart:dt(ar,14), // N = Withdrawal Start Date
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

  // Debts rows 88-89 + 92-101 (A=1,B=2,C=3,D=4,E=5,F=6,H=8)
  var debtRows = [92,93,94,95,96,97,98,99,100,101]; // rows 88,89 are expenses area
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
      curval:num(row,11),
      appr:(num(row,12)||0)*100  // L = Appreciation rate, stored as decimal, returned as %
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

  // Roth plan (rows 150-199, B=col2, C=col3) — 50 years 2026-2075
  var rothPlan = [];
  for (var i=0; i<50; i++) {
    var row = 150+i;
    rothPlan.push({year:2026+i, p1:num(row,2), p2:num(row,3)});
  }

  // Read Master sheet for year-by-year projection data (charts)
  var projections = {years:[], income:[], withdrawals:[], endLiquid:[],
                     preTax:[], roth:[], taxable:[], hsa:[],
                     federalTaxes:[], taxableIncome:[], stateIncome:[],
                     acctWithdrawals:[], totalRMD:[], byYear:[]};
  var masterYear1FedTax = 0, masterYear1TaxableIncome = 0, masterYear1Set = false;
  try {
    var masterSheet = ss.getSheetByName('Master');
    if (masterSheet) {
      var masterData = masterSheet.getRange('A8:CC200').getValues();
      masterData.forEach(function(row) {
        // col A = year — handle both plain number and date object
        var yr;
        var rawYr = row[0];
        if (rawYr instanceof Date) {
          yr = rawYr.getFullYear();
        } else {
          yr = Number(rawYr);
          // If it looks like a date serial (large number), convert
          if (yr > 40000 && yr < 100000) {
            // Google Sheets date serial — convert to year
            var d = new Date((yr - 25569) * 86400 * 1000);
            yr = d.getFullYear();
          }
        }
        if (!yr || yr < 2020 || yr > 2200) return;
        projections.years.push(yr);
        projections.income.push(Math.round(Number(row[14])||0));     // O  = Total Gross Income (unchanged)
        projections.withdrawals.push(Math.round(Number(row[44])||0)); // AN = Final Gross Withdrawal (+5)
        projections.endLiquid.push(Math.round(Number(row[80])||0));   // BX = End of Year Liquid (+5)
        // Federal taxes for effective rate calculation
        var fedTaxes     = Number(row[40])||0;  // AJ = Total Federal Taxes (+5)
        var taxableIncome= Number(row[39])||0;  // AI = Total Taxable Income (+5)
        projections.federalTaxes.push(Math.round(fedTaxes));
        projections.taxableIncome.push(Math.round(taxableIncome));
        projections.stateIncome.push(Math.round(Number(row[42])||0)); // AP = State tax
        projections.totalRMD.push(Math.round(Number(row[35])||0));     // Total RMD (formula output index 35)
        // FULL per-year snapshot — lets the dashboard render every card for any chosen year.
        function N(i){ return Math.round(Number(row[i])||0); }
        projections.byYear.push({
          year:yr, craigAge:N(2), genaAge:N(3),
          gSal:N(4), gSS:N(5), gProjSS:N(6), gPen:N(7), gOth:N(8),
          cSal:N(9), cSS:N(10), cProjSS:N(11), cPen:N(12), cOth:N(13),
          totalGross:N(14),
          craigContrib:N(15), craigMatch:N(16), genaContrib:N(17), genaMatch:N(18), totalContrib:N(19),
          stdDed:N(20), baseLiving:N(23), healthcare:N(26), debt:N(27), totalNeed:N(28),
          gap:N(29), surplus:N(30), chosenExtra:N(31), fundingRequired:N(32),
          rmdC:N(33), rmdG:N(34), totalRMD:N(35),
          convC:N(36), convG:N(37), totalConv:N(38),
          totalTaxable:N(39), federalTaxes:N(40), stateTax:N(41), totalTaxes:N(43),
          acctW:(function(){ var a=[]; for(var k=46;k<=57;k++) a.push(N(k)); return a; })(),
          monthlyWithdraw:N(58),
          acctEnd:(function(){ var a=[]; for(var k=59;k<=70;k++) a.push(N(k)); return a; })(),
          preTaxTotal:N(72), rothTotal:N(73), taxableTotal:N(74), hsaTotal:N(75),
          endLiquid:N(80)
        });
        // Per-account WITHDRAWALS for this year — 12 account slots (var_D 1..12),
        // indices 46..57 on the SAME basis as the verified row[59]=Craig 401k ending.
        // Slot order matches the account order used for ending balances / account names.
        var acctW = [];
        for (var wi = 46; wi <= 57; wi++) { acctW.push(Math.round(Number(row[wi])||0)); }
        projections.acctWithdrawals.push(acctW);
        // Account type TOTALS — correct column indices
        var p1_401k   = Number(row[59])||0;  // BH = Craig 401k ending balance
        var p1_roth   = Number(row[60])||0;  // BI = Craig Roth ending balance
        var p1_brok   = Number(row[61])||0;  // BJ = Craig Brokerage ending balance
        var p2_401k   = Number(row[62])||0;  // BK = Gena 401k ending balance
        var p2_roth   = Number(row[63])||0;  // BL = Gena Roth ending balance
        var p2_brok   = Number(row[64])||0;  // BM = Gena Brokerage ending balance
        var hsa       = Number(row[65])||0;  // BN = HSA ending balance
        // Portfolio type SUMMARIES for chart breakdowns
        var totalPreTax = Number(row[72])||0;  // BU = total pre-tax
        var totalRoth   = Number(row[73])||0;  // BV = total roth
        var totalTaxable= Number(row[74])||0;  // BW = total taxable/liquid
        var totalHSA    = Number(row[75])||0;  // BX = total HSA
        // Capture year 1 values for effective rate
        if (!masterYear1Set && yr >= 2020) {
          masterYear1FedTax        = fedTaxes;
          masterYear1TaxableIncome = taxableIncome;
          masterYear1Set = true;
        }
        projections.preTax.push(Math.round(totalPreTax||0));
        projections.roth.push(Math.round(totalRoth||0));
        projections.taxable.push(Math.round(totalTaxable||0));
        projections.hsa.push(Math.round(totalHSA||0));
      });

      // ── EXTEND PROJECTIONS TO 100 YEARS ─────────────────────────────
      // The sheet formula stops at the last death year. For MC and charts,
      // we extend the projection to 100 years from plan start, simulating
      // portfolio growth with zero income/withdrawals after both partners die.
      if (projections.years.length > 0) {
        var lastSheetYr  = projections.years[projections.years.length - 1];
        var firstYr      = projections.years[0];
        var targetEndYr  = firstYr + 99; // 100 years from plan start year
        
        // Derive avg return from last known portfolio balance trend
        // Use a simple 7% default — same as base assumption
        var baseReturn = 0.07;
        try {
          // Try to get actual avg return from account data
          if (port && port.accounts && port.accounts.length) {
            var rs = 0, rc = 0;
            port.accounts.forEach(function(a) {
              if ((a.expectedReturn||0) > 0) { rs += a.expectedReturn/100; rc++; }
            });
            if (rc > 0) baseReturn = rs / rc;
          }
        } catch(e2) {}
        
        var lastBal = projections.endLiquid[projections.endLiquid.length - 1] || 0;
        var lastPre = projections.preTax[projections.preTax.length - 1] || 0;
        var lastRoth= projections.roth[projections.roth.length - 1] || 0;
        var lastTax = projections.taxable[projections.taxable.length - 1] || 0;
        var lastHsa = projections.hsa[projections.hsa.length - 1] || 0;
        
        for (var extYr = lastSheetYr + 1; extYr <= targetEndYr; extYr++) {
          // Portfolio grows at base return, no income, no withdrawals
          lastBal  = Math.round(Math.max(0, lastBal  * (1 + baseReturn)));
          lastPre  = Math.round(Math.max(0, lastPre  * (1 + baseReturn)));
          lastRoth = Math.round(Math.max(0, lastRoth * (1 + baseReturn)));
          lastTax  = Math.round(Math.max(0, lastTax  * (1 + baseReturn)));
          lastHsa  = Math.round(Math.max(0, lastHsa  * (1 + baseReturn)));
          
          projections.years.push(extYr);
          projections.income.push(0);         // no income after both die
          projections.withdrawals.push(0);    // no withdrawals needed
          projections.endLiquid.push(lastBal);
          projections.federalTaxes.push(0);
          projections.taxableIncome.push(0);
          projections.stateIncome.push(0);
          projections.preTax.push(lastPre);
          projections.roth.push(lastRoth);
          projections.taxable.push(lastTax);
          projections.hsa.push(lastHsa);
          projections.acctWithdrawals.push([0,0,0,0,0,0,0,0,0,0,0,0]); // no withdrawals after both die
          projections.totalRMD.push(0);
          projections.byYear.push({
            year:extYr, craigAge:0, genaAge:0,
            gSal:0,gSS:0,gProjSS:0,gPen:0,gOth:0, cSal:0,cSS:0,cProjSS:0,cPen:0,cOth:0,
            totalGross:0, craigContrib:0,craigMatch:0,genaContrib:0,genaMatch:0,totalContrib:0,
            stdDed:0, baseLiving:0, healthcare:0, debt:0, totalNeed:0,
            gap:0, surplus:0, chosenExtra:0, fundingRequired:0,
            rmdC:0, rmdG:0, totalRMD:0, convC:0, convG:0, totalConv:0,
            totalTaxable:0, federalTaxes:0, stateTax:0, totalTaxes:0,
            acctW:[0,0,0,0,0,0,0,0,0,0,0,0], monthlyWithdraw:0,
            acctEnd:[0,0,0,0,0,0,0,0,0,0,0,0],
            preTaxTotal:lastPre, rothTotal:lastRoth, taxableTotal:lastTax, hsaTotal:lastHsa,
            endLiquid:lastBal
          });
        }
      }
      // ── END EXTENSION ────────────────────────────────────────────────

    }
  } catch(e) {
    Logger.log('Master read error: ' + e);
  }

  // Set effective federal rate from Master year-1 data (AJ/AI)
  if (masterYear1TaxableIncome > 0) {
    tax.effectiveRate = Math.round((masterYear1FedTax / masterYear1TaxableIncome) * 10000) / 10000;
    tax.federalTaxes  = masterYear1FedTax;
    tax.taxableIncome = masterYear1TaxableIncome;
  }

  return {
    meta:{planYear:new Date().getFullYear()},
    people:{craig:p1, gena:p2},
    global:gl, tax:tax,
    portfolio:{accounts:accounts, total:portfolioTotal, slots:acctSlots},
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
      if (g.B13 !== undefined && g.B13 !== null && g.B13 !== '') setN('B13', g.B13);
      if (g.B14 !== undefined && g.B14 !== null && g.B14 !== '') setN('B14', g.B14);
      if (g.B15 !== undefined) setPct('B15', g.B15);
      if (g.B16) set('B16', g.B16);
      if (g.B17 !== undefined && g.B17 !== null && g.B17 !== '') setN('B17', g.B17);
      if (g.B18 !== undefined && g.B18 !== null && g.B18 !== '') setN('B18', g.B18);
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

    // ── AUTO-UPDATE ACCOUNT OWNER NAMES ──────────────────────
    // When partner names change, update account owner cells that
    // don't match current valid names (p1New, p2New, Joint)
    try {
      var p1New = data.partner1 && data.partner1.B31 ? String(data.partner1.B31).trim() : '';
      var p2New = data.partner2 && data.partner2.D31 ? String(data.partner2.D31).trim() : '';
      if (p1New && p2New) {
        var ownerRange = inp.getRange('C105:C116').getValues();
        var accountData = inp.getRange('A105:N116').getValues();
        var validNames = [p1New.toLowerCase(), p2New.toLowerCase(), 'joint', ''];
        var updated = false;
        for (var oi=0; oi<ownerRange.length; oi++) {
          var currentOwner = String(ownerRange[oi][0]||'').trim();
          var currentLower = currentOwner.toLowerCase();
          // If current owner is not a valid name, figure out which partner it belongs to
          if (validNames.indexOf(currentLower) < 0 && currentOwner !== '') {
            // Check accounts data: if account has contribution data, 
            // it's likely a personal account
            // Use position heuristic: accounts 0-2 tend to be P1, 3-5 tend to be P2
            // But better: check if the account has contributions matching P1 or P2
            var hasP1Contrib = Number(accountData[oi][9]) > 0; // col J contrib
            // Default: if first few accounts, assign to P1; later ones to P2
            // Actually safest: keep as-is but update if it looks like an old test name
            // The account write will have set the owner based on web app data
            // so just log what we find
            Logger.log('Found non-valid owner: "' + currentOwner + '" in row ' + (105+oi));
          }
        }
        // Better approach: the accounts array from data already has correct owners
        // (set by the web app). The sheet C105:C116 should match data.accounts[i].owner
        // Let's use that directly:
        if (data.accounts && data.accounts.length) {
          for (var ai=0; ai<Math.min(data.accounts.length, 12); ai++) {
            var webOwner = String(data.accounts[ai].owner||'').trim();
            var sheetOwner = String(ownerRange[ai][0]||'').trim();
            if (webOwner && webOwner !== sheetOwner) {
              ownerRange[ai][0] = webOwner;
              updated = true;
              Logger.log('Updating owner row ' + (105+ai) + ': "' + sheetOwner + '" → "' + webOwner + '"');
            }
          }
        }
        if (updated) inp.getRange('C105:C116').setValues(ownerRange);
      }
    } catch(e) { Logger.log('Auto-update owner names error: ' + e); }

    // ── EXPENSES ──────────────────────────────────────────────
    // A=editable label, B=monthly($), C=yearly(FORMULA-NEVER WRITE), D=notes
    if (data.expenses && data.expenses.length) {
      var skip = {62:true, 70:true, 82:true, 86:true, 87:true};
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
    // A=include, B=name, C=monthly, D=annual(C*12), E=start, F=end, H=balance
    if (data.debts && data.debts.length) {
      var debtRows = [92,93,94,95,96,97,98,99,100,101]; // rows 88,89 are expenses area
      data.debts.forEach(function(d, i) {
        if (i >= debtRows.length) return;
        var r = debtRows[i];
        var mo = Number(d.mo)||0;
        // Only write if row has meaningful data (name OR monthly > 0)
        // This prevents wiping existing sheet data with blank web rows
        var hasData = (d.name && String(d.name).trim() !== '') || mo > 0 || Number(d.bal) > 0;
        if (!hasData && i >= 2) return; // Skip empty extra rows (keep sheet data intact)

        // A=Include, B=Name, C=PurchasePrice, D=Monthly, E=Annual, F=Start, G=End, H=Balance
        inp.getRange('A'+r).setValue(d.inc||'Yes');
        if (d.name && String(d.name).trim()) inp.getRange('B'+r).setValue(String(d.name));
        if (d.pp > 0) inp.getRange('C'+r).setValue(Number(d.pp));  // C = Purchase Price
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
        if (d.appr)   inp.getRange('L'+r).setValue(Number(d.appr)/100); // store as decimal
      });
    }

    // ── ACCOUNTS ──────────────────────────────────────────────
    // A=include, B=name, C=owner, D=type, E=balance, F=return(decimal), G=status, H=dashboard
    if (data.accounts && data.accounts.length) {
      Logger.log('Accounts received: ' + data.accounts.length);
      Logger.log('First account owner: ' + data.accounts[0].owner);
      Logger.log('p1name: ' + (data.partner1 ? data.partner1.B31 : 'MISSING'));
      Logger.log('p2name: ' + (data.partner2 ? data.partner2.D31 : 'MISSING'));
      var p1name = data.partner1 && data.partner1.B31 ? String(data.partner1.B31).trim() : '';
      var p2name = data.partner2 && data.partner2.D31 ? String(data.partner2.D31).trim() : '';

      // Clear owner validation temporarily
      try { inp.getRange('C105:C124').clearDataValidations(); } catch(e){}

      var acctData = [];
      for (var i=0; i<12; i++) {
        var a = data.accounts[i] || {};
        var owner = String(a.owner||'Joint').trim();
        var ownerLower = owner.toLowerCase();
        var validOwners = [p1name.toLowerCase(), p2name.toLowerCase(), 'joint', ''];
        if (validOwners.indexOf(ownerLower) < 0) {
          // Unrecognized owner name — use index position to guess:
          // accounts 0,1,2 (first 3) → P1; account 3 → P2; rest → Joint
          owner = (i < 3) ? (p1name||'Joint') : (i===3 ? (p2name||'Joint') : 'Joint');
        }
        if (!owner || owner === '') owner = 'Joint';
        // ret comes in as whole number (7.0), store as decimal (0.07)
        var ret = Number(a.ret||a.expectedReturn||0);
        if (ret > 1) ret = ret/100;
        acctData.push([
          a.inc||a.showInCalc||'Yes',        // A = Include
          a.name||'',                        // B = Name
          owner,                             // C = Owner
          a.type||'',                        // D = Type
          Number(a.bal||a.balance)||0,       // E = Balance
          ret,                               // F = Return (decimal)
          a.status||'Use for Withdrawals',   // G = Status
          a.dash||a.showOnDashboard||'No',   // H = Dashboard
          '',                                // I = Summary (read-only, skip)
          (function(){ var v=parseFloat(a.contrib); return (isFinite(v) && v>=0 && v<=200000) ? Math.round(v) : 0; })(),  // J = Annual Contribution
          (function(){ var v=parseFloat(a.match);  return (isFinite(v) && v>=0 && v<=200000) ? Math.round(v) : 0; })(),  // K = Employer Match
          '',                                // L = Contrib Start (written separately below)
          '',                                // M = Contrib End (written separately below)
          ''                                 // N = Withdrawal Start (written separately below)
        ]);
      }
      // Force J and K to number format BEFORE clearing/writing
      // This prevents Sheets from auto-formatting as dates
      inp.getRange('J105:K116').setNumberFormat('0');
      SpreadsheetApp.flush();
      // Clear J-N completely before writing to remove any corrupt values
      inp.getRange('J105:N116').clearContent();
      SpreadsheetApp.flush(); // Ensure clear completes before writing
      // Log what we're about to write to J and K
      Logger.log('Writing accounts J/K: ' + acctData.map(function(r){return r[9]+'/'+r[10];}).join(', '));
      Logger.log('Writing owners: ' + acctData.map(function(r){return r[2];}).join(', '));
      inp.getRange('A105:N116').setValues(acctData.slice(0,12));
      // Write contribution dates separately using setD (handles date format correctly)
      var acctRows = [105,106,107,108,109,110,111,112,113,114,115,116];
      for (var ai=0; ai<12; ai++) {
        var aa = data.accounts[ai] || {};
        var r2 = acctRows[ai];
        if (aa.contribStart) setD('L'+r2, aa.contribStart);
        if (aa.contribEnd)   setD('M'+r2, aa.contribEnd);
        if (aa.withdrawStart) setD('N'+r2, aa.withdrawStart);
      }

      // Restore owner validation with actual partner names
      try {
        var p1n = data.partner1 && data.partner1.B31 ? String(data.partner1.B31).trim() : 'Partner 1';
        var p2n = data.partner2 && data.partner2.D31 ? String(data.partner2.D31).trim() : 'Partner 2';
        var ownerRule = SpreadsheetApp.newDataValidation()
          .requireValueInList([p1n, p2n, 'Joint'], true)
          .setAllowInvalid(true).build();
        inp.getRange('C105:C116').setDataValidation(ownerRule);
      } catch(e){ Logger.log('Owner validation error: ' + e); }
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
      var allArr = [];
      for (var i=0; i<50; i++) {
        var yr = 2026 + i;
        var p1 = 0, p2 = 0;
        data.rothPlan.forEach(function(row) {
          if (Number(row.year) === yr) { p1=Number(row.p1)||0; p2=Number(row.p2)||0; }
        });
        allArr.push([yr, p1, p2]);
      }
      // Force year column to plain number format first
      inp.getRange('A150:A199').setNumberFormat('0');
      inp.getRange('B150:C199').setNumberFormat('0');
      inp.getRange('A150:C199').setValues(allArr);
    }

    SpreadsheetApp.flush();

    // Apply theme colors to contribution columns in Master sheet
    try { applyMasterThemeFormatting(ss); } catch(e) { Logger.log('Theme formatting skipped: '+e); }

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

// ── MASTER SHEET THEME FORMATTING ────────────────────────────────────
// Applies conditional formatting to new contribution columns P-T
// so they automatically match the partner theme colors chosen in Inputs
function applyMasterThemeFormatting(ss) {
  try {
    var master = ss.getSheetByName('Master');
    if (!master) return;

    var p1Themes = {
      'Steel Blue': '#60828C',  // Craig option 1
      'Slate Teal': '#456A73',  // Craig option 2
    };
    var p2Themes = {
      'Muted Mauve': '#B07BBF', // Gena option 1
      'Frost Plum':  '#8D6B92', // Gena option 2
    };
    var lightThemes = []; // none of these are light enough to need dark text

    // Clear existing conditional formatting on P:T columns only
    var existingRules = master.getConditionalFormatRules();
    var keptRules = existingRules.filter(function(rule) {
      var ranges = rule.getRanges();
      // Keep rules not touching columns P-T (cols 16-20)
      return !ranges.some(function(r) {
        var col = r.getColumn();
        return col >= 16 && col <= 20;
      });
    });

    var newRules = keptRules.slice();

    // P6:Q (Craig's contribution columns) — Partner 1 theme
    var p1Range = master.getRange('P6:Q');
    Object.keys(p1Themes).forEach(function(themeName) {
      var hex = p1Themes[themeName];
      var fontColor = (lightThemes.indexOf(themeName) >= 0) ? '#041E2F' : '#FFFFFF';
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=Inputs!$B$34="' + themeName + '"')
        .setBackground(hex)
        .setFontColor(fontColor)
        .setBold(true)
        .setRanges([p1Range])
        .build();
      newRules.push(rule);
    });

    // R6:S (Gena's contribution columns) — Partner 2 theme
    var p2Range = master.getRange('R6:S');
    Object.keys(p2Themes).forEach(function(themeName) {
      var hex = p2Themes[themeName];
      var fontColor = (lightThemes.indexOf(themeName) >= 0) ? '#041E2F' : '#FFFFFF';
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=Inputs!$D$34="' + themeName + '"')
        .setBackground(hex)
        .setFontColor(fontColor)
        .setBold(true)
        .setRanges([p2Range])
        .build();
      newRules.push(rule);
    });

    master.setConditionalFormatRules(newRules);

    // T column (Total) — static dark navy, no conditional formatting needed
    // Just set it directly
    master.getRange('T6:T').setBackground('#1A3A52').setFontColor('#FFFFFF').setFontWeight('bold');

    Logger.log('Master theme formatting applied');
  } catch(e) {
    Logger.log('applyMasterThemeFormatting error: ' + e);
  }
}


// ════════════════════════════════════════════════════════════════════
// STRESS TEST — write bear-market params, recalc, read stressed results,
// then RESTORE the customer's original B12-B19 so their plan is untouched.
// Drag values arrive as WHOLE NUMBERS (e.g. 10) and are written as decimals
// (0.10) using the same /100 convention as the normal save (setPct).
// ════════════════════════════════════════════════════════════════════
function runStressTest(ss, inp, p) {
  try {
    // --- 1. Save the customer's current stress cells (B12-B19) ---
    var orig = inp.getRange('B12:B20').getValues(); // 9 rows: B12-B19 bear params + B20 inflation

    // --- 2. Parse incoming test params (whole-number drags) ---
    function pnum(v, d){ var n = Number(v); return isNaN(n) ? d : n; }
    var eType  = (p.eType  !== undefined && p.eType  !== '') ? String(p.eType)  : 'bear';
    var eStart = pnum(p.eStart, 0);
    var eDur   = pnum(p.eDur,   0);
    var eDrag  = pnum(p.eDrag,  0);   // whole number, e.g. 10
    var lType  = (p.lType  !== undefined && p.lType  !== '') ? String(p.lType)  : '';
    var lStart = pnum(p.lStart, 0);
    var lDur   = pnum(p.lDur,   0);
    var lDrag  = pnum(p.lDrag,  0);
    // Inflation override (whole-number %, e.g. 5). null = leave the customer's saved inflation untouched.
    var inflPct = (p.infl !== undefined && p.infl !== '') ? pnum(p.infl, null) : null;

    // --- 3. Write the test values ---
    // Only write an EARLY bear if a duration & drag were given; same for LATE.
    var newVals = orig.map(function(r){ return [r[0]]; }); // clone
    if (eDur > 0 && eDrag !== 0) {
      newVals[0] = [eType];          // B12 type
      newVals[1] = [eStart];         // B13 start
      newVals[2] = [eDur];           // B14 duration
      newVals[3] = [eDrag/100];      // B15 drag -> decimal
    } else {
      // no early stress this test — clear the trigger so it doesn't apply
      newVals[0] = [''];
    }
    if (lDur > 0 && lDrag !== 0) {
      newVals[4] = [lType || 'bear']; // B16 type
      newVals[5] = [lStart];          // B17 start
      newVals[6] = [lDur];            // B18 duration
      newVals[7] = [lDrag/100];       // B19 drag -> decimal
    } else {
      newVals[4] = ['']; // clear late trigger
    }
    // Inflation override -> B20 (index 8). Stored as a DECIMAL (0.05), same /100 convention as setPct.
    // Only override when a value was passed; otherwise the customer's saved inflation is preserved.
    if (inflPct !== null && !isNaN(inflPct)) {
      newVals[8] = [inflPct / 100];
    }
    inp.getRange('B12:B20').setValues(newVals);

    // --- 4. Force recalc and WAIT for the safe-spend SOLVER to actually settle. ---
    // The Master array recalculates fast, but the safe-spend (B119) sits DOWNSTREAM of
    // it and lags. A fixed 1.2s wait was reading B119 before it finished — handing back
    // the un-stressed (base) safe-spend. Instead, poll B119 until it stops moving.
    SpreadsheetApp.flush();
    var _prevSafe = null, _stable = 0;
    for (var _w = 0; _w < 14; _w++) {          // up to ~14 x 1500ms ≈ 21s max
      Utilities.sleep(1500);
      SpreadsheetApp.flush();
      var _curSafe = Number(inp.getRange('B119').getValue()) || 0;
      if (_w >= 2 && _prevSafe !== null && Math.abs(_curSafe - _prevSafe) < 1) {
        _stable++;
        if (_stable >= 2) break;               // settled: two stable reads after a few passes
      } else {
        _stable = 0;
      }
      _prevSafe = _curSafe;
    }

    // --- 5. Read stressed results ---
    function cell(a1){ var v = inp.getRange(a1).getValue(); return Number(v)||0; }
    var safeExtra = cell('B119');
    var ending    = cell('B133');
    var surplus   = cell('B120');
    var legacyGoal= cell('B8');

    // Year balance first drops below legacy goal + full stressed year-by-year series (for a real chart)
    var belowYear = null;
    var series = []; // [{year, bal}] stressed ending-liquid per year, straight from the recalculated Master
    var master = ss.getSheetByName('Master');
    if (master) {
      var md = master.getRange('A8:CC200').getValues();
      for (var i=0;i<md.length;i++){
        var row = md[i];
        var rawYr = row[0], yr;
        if (rawYr instanceof Date) yr = rawYr.getFullYear();
        else { yr = Number(rawYr); if (yr>40000 && yr<100000){ var d=new Date((yr-25569)*86400*1000); yr=d.getFullYear(); } }
        if (!yr || yr<2020 || yr>2200) continue;
        var endLiq = Number(row[80])||0;
        series.push({ year:yr, bal:Math.round(endLiq) });
        if (belowYear===null && legacyGoal>0 && endLiq < legacyGoal){ belowYear = yr; }
      }
    }

    // --- 6. RESTORE the customer's original cells (B12-B19 bear params + B20 inflation) ---
    inp.getRange('B12:B20').setValues(orig);
    SpreadsheetApp.flush();

    return {
      ok:true,
      safeExtra:Math.round(safeExtra),
      ending:Math.round(ending),
      surplus:Math.round(surplus),
      legacyGoal:Math.round(legacyGoal),
      belowYear:belowYear,
      series:series
    };
  } catch(err) {
    // Best-effort restore even on error
    try { if (orig) inp.getRange('B12:B20').setValues(orig); SpreadsheetApp.flush(); } catch(e2){}
    return { ok:false, error: err.toString() };
  }
}

// ── Hidden tab persistence for the Test 1-5 slots ──────────────────
function getStressTab(ss) {
  var sh = ss.getSheetByName('_StressTests');
  if (!sh) {
    sh = ss.insertSheet('_StressTests');
    sh.getRange('A1').setValue('RetireBlueprint Pro — saved stress tests (do not edit)');
  }
  try { sh.hideSheet(); } catch(e){} // keep it tucked away
  return sh;
}

function loadStressTests(ss) {
  try {
    var sh = ss.getSheetByName('_StressTests');
    if (!sh) return { ok:true, tests:[] };
    var json = sh.getRange('A2').getValue();
    if (!json) return { ok:true, tests:[] };
    return { ok:true, tests: JSON.parse(String(json)) };
  } catch(err) {
    return { ok:true, tests:[] }; // never block the page on a load failure
  }
}

function saveStressTests(ss, tests) {
  try {
    var sh = getStressTab(ss);
    sh.getRange('A2').setValue(JSON.stringify(tests || []));
    SpreadsheetApp.flush();
    return { ok:true };
  } catch(err) {
    return { ok:false, error: err.toString() };
  }
}
