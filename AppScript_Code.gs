// RetireBlueprint Pro — Apps Script
// Built from ground-truth sheet audit. Every cell mapped correctly.
// NEVER write to formula cells: B33,B36,B42,B43,B45,B48,B51(formula?),
//   B86,B87,B119,B120,B128-B130,B133,B134,B140,B142,B143,B144,B146,G122-G124

var SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxCPNzU1A9jT8eiOrc-fg2kWSo9aZmSHYBFfPILZzdZD-IEhDCr0uYhGeOqUSysaCpP/exec';

function doGet(e) {
  var action = e.parameter.action;
  var sheetId = e.parameter.sheetId;

  if (!sheetId) return json({error:'No sheetId'});

  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var inp = ss.getSheetByName('Inputs');
    if (!inp) return json({error:'No Inputs sheet'});

    if (action === 'read') return json(readAll(ss, inp));
    if (action === 'save') {
      var enc = e.parameter.enc;
      var raw = enc === 'b64'
        ? Utilities.newBlob(Utilities.base64Decode(e.parameter.data)).getDataAsString()
        : e.parameter.data;
      var data = JSON.parse(raw);
      return json(writeInputs(ss, inp, data));
    }
    return json({error:'Unknown action: '+action});
  } catch(err) {
    return json({error: err.toString()});
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ───────────────────────────────────────────────────
function readAll(ss, inp) {
  function v(cell)  { return inp.getRange(cell).getValue(); }
  function pct(val) { return Number(val)||0; }           // stored as decimal e.g. 0.07
  function dt(val)  {
    if (!val) return '';
    try {
      var d = new Date(val);
      return isNaN(d) ? '' : (d.getMonth()+1)+'/'+d.getDate()+'/'+d.getFullYear();
    } catch(e) { return ''; }
  }

  // Partner 1
  var p1 = {
    name:v('B31'), dob:dt(v('B32')), age:Number(v('B33')),
    deathAge:Number(v('B35')), theme:String(v('B34')),
    salary:Number(v('B37')), salaryStart:dt(v('B38')), salaryEnd:dt(v('B39')),
    ssBase:Number(v('B40')), ssFra:Number(v('B41')),
    ssMonthly:Number(v('B43')), ssStartDate:dt(v('B44')),
    pension:Number(v('B46')), pensionStart:dt(v('B47')),
    otherIncome:Number(v('B49')), otherStart:dt(v('B50')), otherEnd:dt(v('B51')),
    healthPreMedicare:Number(v('B52')), healthMedicare:Number(v('B53')), medicareAge:Number(v('B54')),
  };

  // Partner 2
  var p2 = {
    name:v('D31'), dob:dt(v('D32')), age:Number(v('D33')),
    deathAge:Number(v('D35')), theme:String(v('D34')),
    salary:Number(v('D37')), salaryStart:dt(v('D38')), salaryEnd:dt(v('D39')),
    ssBase:Number(v('D40')), ssFra:Number(v('D41')),
    ssMonthly:Number(v('D43')), ssStartDate:dt(v('D44')),
    pension:Number(v('D46')), pensionStart:dt(v('D47')),
    otherIncome:Number(v('D49')), otherStart:dt(v('D50')), otherEnd:dt(v('D51')),
    healthPreMedicare:Number(v('D52')), healthMedicare:Number(v('D53')), medicareAge:Number(v('D54')),
  };

  // Global settings
  var gl = {
    planStartYear:dt(v('B6')), projectionYears:Number(v('B7')),
    legacyGoal:Number(v('B8')), safetyFloor:Number(v('B9')),
    survivorReduction:pct(v('B10')),
    assumeExtraSpend:String(v('B11')),
    stressEarly:String(v('B12')), stressEarlyStart:Number(v('B13')),
    stressEarlyDur:Number(v('B14')), stressEarlyDrag:pct(v('B15')),
    stressLate:String(v('B16')), stressLateStart:Number(v('B17')),
    stressLateDur:Number(v('B18')), stressLateDrag:pct(v('B19')),
    filingStatus:String(v('B24')), standardDeduction:Number(v('B25')),
  };

  // Tax
  var tax = {
    inflation:pct(v('B20')), ssCola:pct(v('B21')),
    baseTaxAdj:pct(v('B22')), marginalExtra:pct(v('B23')),
    stateRate:pct(v('B26')), stdDedInflation:pct(v('B27')),
    healthcareInflation:pct(v('B28')),
    rothConversionSavings:Number(v('B144')),
  };

  // Accounts
  var acctRows = inp.getRange('A105:H116').getValues();
  var accounts = [], portfolioTotal = 0;
  acctRows.forEach(function(row) {
    if (!row[1]) return;
    var bal = Number(row[4])||0;
    var ret = Number(row[5])||0; // stored as decimal in sheet
    if (String(row[0]).toLowerCase()==='yes') portfolioTotal += bal;
    accounts.push({
      showInCalc:String(row[0]), name:String(row[1]), owner:String(row[2]),
      type:String(row[3]), balance:bal,
      expectedReturn:ret*100, // multiply by 100 for display (0.07 → 7)
      status:String(row[6]), showOnDashboard:String(row[7])
    });
  });

  // Expenses — A=label, B=monthly, C=yearly(formula), D=notes
  var expRowNums = [58,59,60,61,62,63,64,65,66,67,68,69,71,72,73,74,75,76,77,78,79,80,81,83,84,85];
  var aVals = inp.getRange('A58:A85').getValues();
  var bVals = inp.getRange('B58:B85').getValues();
  var dVals = inp.getRange('D58:D85').getValues();
  var expenses = expRowNums.map(function(r) {
    var i = r - 58;
    return {
      row:r,
      name:String(aVals[i][0]||''),
      monthly:Number(bVals[i][0])||0,
      annual:(Number(bVals[i][0])||0)*12,
      note:String(dVals[i][0]||'')
    };
  });

  // Debts — rows 88-89 + 92-101
  // A=include, B=name, C=monthly, D=annual, E=start, F=end, H=balance
  var debt88 = inp.getRange('A88:H89').getValues();
  var debt92 = inp.getRange('A92:H101').getValues();
  var debts = [];
  function parseDebt(row) {
    return {inc:String(row[0]||'No'), name:String(row[1]||''),
      mo:Number(row[2])||0, ann:Number(row[3])||0,
      start:dt(row[4]), end:dt(row[5]), bal:Number(row[7])||0};
  }
  debt88.forEach(function(r){ debts.push(parseDebt(r)); });
  debt92.forEach(function(r){ debts.push(parseDebt(r)); });

  // Spending/solver
  var baseSpend = Number(v('B87'))||0;
  var safeExtra = Number(v('B119'))||0;
  function phaseYear(cell) {
    var val = v(cell);
    if (!val) return null;
    var d = new Date(val);
    return isNaN(d) ? null : d.getFullYear();
  }
  var spending = {
    baseAnnual:baseSpend, safeExtra:safeExtra,
    surplus:Number(v('B120'))||0,
    phase1Start:phaseYear('C122'), phase1End:phaseYear('D122'),
    phase2Start:phaseYear('C123'), phase2End:phaseYear('D123'),
    phase3Start:phaseYear('C124'), phase3End:phaseYear('D124'),
    phase1Weight:Number(v('E122'))||1.15,
    phase2Weight:Number(v('E123'))||1.0,
    phase3Weight:Number(v('E124'))||0.9,
    phase1Override:Number(v('F122'))||0,
    phase2Override:Number(v('F123'))||0,
    phase3Override:Number(v('F124'))||0,
    phase1Extra:Number(v('G122'))||0,
    phase2Extra:Number(v('G123'))||0,
    phase3Extra:Number(v('G124'))||0,
  };

  // Legacy
  var endLiquid = Number(v('B133'))||0;
  var legacyGoal = Number(v('B8'))||0;
  var legacy = {
    goal:legacyGoal, safetyFloor:Number(v('B9')),
    projectedEnding:endLiquid,
    variance:endLiquid - legacyGoal,
    achievedPct:legacyGoal>0 ? Math.round(endLiquid/legacyGoal*100) : 100,
    status:String(v('B134')||''),
  };

  // Roth
  var roth = {
    year:Number(v('B139'))||0,
    bracket:String(v('B141')||'12%'),
    taxableIncome:Number(v('B140'))||0,
    bracketThreshold:Number(v('B142'))||0,
    optimalAmount:Number(v('B143'))||0,
    marginalTax:Number(v('B144'))||0,
    assumedRate:pct(v('B145')),
    taxSavings:Number(v('B146'))||0,
  };

  // Roth plan amounts
  var rothB = inp.getRange('B150:B169').getValues();
  var rothC = inp.getRange('C150:C169').getValues();
  var rothPlan = [];
  for (var i=0; i<20; i++) {
    var p1amt = Number(rothB[i][0])||0;
    var p2amt = Number(rothC[i][0])||0;
    rothPlan.push({year:2026+i, p1:p1amt, p2:p2amt});
  }

  // Check-in
  var checkIn = {
    actualBalance:Number(v('B149'))||0,
  };

  return {
    meta:{planYear:new Date().getFullYear(), sheetId:''},
    people:{craig:p1, gena:p2},
    global:gl, tax:tax,
    portfolio:{accounts:accounts, total:portfolioTotal},
    expenses:expenses, debts:debts,
    spending:spending, legacy:legacy,
    roth:roth, rothPlan:rothPlan,
    checkIn:checkIn,
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
      if (g.B6)  setD('B6', g.B6);
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
      var skip = {62:true, 70:true, 82:true, 86:true, 87:true};
      data.expenses.forEach(function(exp) {
        var r = Number(exp.row);
        if (!r || r<58 || r>85 || skip[r]) return;
        if (exp.name !== undefined) inp.getRange('A'+r).setValue(String(exp.name));
        inp.getRange('B'+r).setValue(Number(exp.monthly)||0);
        // NEVER WRITE C (formula =B*12)
        if (exp.note !== undefined) inp.getRange('D'+r).setValue(String(exp.note||''));
      });
    }

    // ── DEBTS ─────────────────────────────────────────────────
    // A=include, B=name, C=monthly, D=annual(C*12), E=start, F=end, H=balance
    if (data.debts && data.debts.length) {
      var debtRows = [88,89,92,93,94,95,96,97,98,99,100,101];
      data.debts.forEach(function(d, i) {
        if (i >= debtRows.length) return;
        var r = debtRows[i];
        inp.getRange('A'+r).setValue(d.inc||'No');
        if (d.name) inp.getRange('B'+r).setValue(String(d.name));
        var mo = Number(d.mo)||0;
        inp.getRange('C'+r).setValue(mo);
        inp.getRange('D'+r).setValue(mo*12);
        if (d.start) setD('E'+r, d.start);
        if (d.end)   setD('F'+r, d.end);
        var bal = Number(d.bal)||0;
        if (bal) inp.getRange('H'+r).setValue(bal);
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
          a.dash||a.showOnDashboard||'No'  // H
        ]);
      }
      inp.getRange('A105:H116').setValues(acctData);

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
      var p1arr = [], p2arr = [];
      for (var i=0; i<20; i++) { p1arr.push([0]); p2arr.push([0]); }
      data.rothPlan.forEach(function(row) {
        var idx = Number(row.year) - 2026;
        if (idx>=0 && idx<20) {
          p1arr[idx] = [Number(row.p1)||0];
          p2arr[idx] = [Number(row.p2)||0];
        }
      });
      inp.getRange('B150:B169').setValues(p1arr);
      inp.getRange('C150:C169').setValues(p2arr);
    }

    SpreadsheetApp.flush();
    return {success:true, timestamp:new Date().toISOString()};
  } catch(err) {
    return {success:false, error:err.toString()};
  }
}
