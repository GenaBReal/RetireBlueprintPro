// RetireBlueprint Pro — Complete Apps Script (READ + WRITE)
// Deploy as Web App: Execute as Me, Who has access: Anyone
// This ONE deployment serves ALL customers — each passes their own Sheet ID

var FALLBACK_SHEET_ID = '1bb7or0XWuU0tAtdwkIJm1HUetY2BiEbxVBVoH6jsYAY'; // Blank customer template

function doGet(e) {
  var callback = e.parameter.callback || 'callback';
  var action   = e.parameter.action   || 'read';
  var sheetId  = e.parameter.sheetId  || FALLBACK_SHEET_ID;
  try {
    if (action === 'save') {
      var rawData = e.parameter.data || '{}';
      var enc = e.parameter.enc || '';
      if (enc === 'b64') {
        rawData = Utilities.newBlob(Utilities.base64Decode(rawData)).getDataAsString();
      }
      var data = JSON.parse(rawData);
      return jsonp(callback, writeInputs(data, sheetId));
    }
    return jsonp(callback, readAll(sheetId));
  } catch(err) {
    return jsonp(callback, {error: err.toString()});
  }
}

function jsonp(cb, data) {
  return ContentService.createTextOutput(cb+'('+JSON.stringify(data)+')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function readAll(sheetId) {
  var ss  = SpreadsheetApp.openById(sheetId);
  var inp = ss.getSheetByName('Inputs');
  var mst = ss.getSheetByName('Master');
  var chk = ss.getSheetByName('Annual Check-In');
  var tax = ss.getSheetByName('Tax Engine');

  var planStart = inp.getRange('B6').getValue();
  var planYear  = planStart ? new Date(planStart).getFullYear() : 2026;

  function v(cell)  { return inp.getRange(cell).getValue(); }
  function pct(val) { var s=String(val).replace('%',''); var n=parseFloat(s); return isNaN(n)?0:(n>1?n/100:n); }
  function dt(val)  { if(!val) return ''; try{var d=new Date(val);return (d.getMonth()+1)+'/'+d.getDate()+'/'+d.getFullYear();}catch(e){return String(val);} }

  var craig = {
    name:v('B31'), dob:dt(v('B32')), age:Number(v('B33')), projectedDeathAge:Number(v('B35')),
    projectedDeathYear:Number(v('B36')), salary:Number(v('B37')),
    salaryStart:dt(v('B38')), retirementDate:dt(v('B39')), ssBase:Number(v('B40')), ssFra:Number(v('B41')),
    ssMonthly:Number(v('B43')), ssAnnual:Number(v('B43'))*12,
    ssStartDate:dt(v('B44')), pension:Number(v('B46')), pensionAnnual:Number(v('B46'))*12,
    pensionStart:dt(v('B47')), otherIncome:Number(v('B49')),
    otherStart:dt(v('B50')), otherEnd:dt(v('B51')),
    healthPreMedicare:Number(v('B52')), healthMedicare:Number(v('B53')), medicareAge:Number(v('B54')),
    theme:String(v('B34')),
  };
  var gena = {
    name:v('D31'), dob:dt(v('D32')), age:Number(v('D33')), projectedDeathAge:Number(v('D35')),
    projectedDeathYear:Number(v('D36')), salary:Number(v('D37')),
    salaryStart:dt(v('D38')), retirementDate:dt(v('D39')), ssBase:Number(v('D40')), ssFra:Number(v('D41')),
    ssMonthly:Number(v('D43')), ssAnnual:Number(v('D43'))*12,
    ssStartDate:dt(v('D44')), pension:Number(v('D46')), pensionAnnual:Number(v('D46'))*12,
    pensionStart:dt(v('D47')), otherIncome:Number(v('D49')),
    otherStart:dt(v('D50')), otherEnd:dt(v('D51')),
    healthPreMedicare:Number(v('D52')), healthMedicare:Number(v('D53')), medicareAge:Number(v('D54')),
    theme:String(v('D34')),
  };

  var acctRows = inp.getRange('A105:H116').getValues();
  var accounts = [], total = 0;
  acctRows.forEach(function(row) {
    if (!row[1]) return;
    var bal = Number(row[4])||0;
    if (String(row[0]).toLowerCase()==='yes') total += bal;
    accounts.push({
      showInCalc:String(row[0]), name:String(row[1]), owner:String(row[2]),
      type:String(row[3]), balance:bal, expectedReturn:Number(row[5])||0,
      status:String(row[6]), showOnDashboard:String(row[7])
    });
  });

  var mdata = mst.getRange('A8:BX83').getValues();
  var years=[],endLiq=[],preTax=[],roth=[],taxable=[],hsa=[],income=[],wd=[];
  mdata.forEach(function(row){
    var yr=Number(row[0]); if(!yr||yr<2026) return;
    years.push(yr);
    endLiq.push(  Math.round((Number(row[75])||0)/1000));
    preTax.push(  Math.round((Number(row[67])||0)/1000));
    roth.push(    Math.round((Number(row[68])||0)/1000));
    taxable.push( Math.round((Number(row[69])||0)/1000));
    hsa.push(     Math.round((Number(row[70])||0)/1000));
    income.push(  Math.round((Number(row[14])||0)/1000));
    wd.push(      Math.round((Number(row[39])||0)/1000));
  });
  var peakVal=0,peakYear=planYear;
  endLiq.forEach(function(v,i){if(v>peakVal){peakVal=v;peakYear=years[i];}});

  var checkIn=[];
  try {
    chk.getRange('A8:F57').getValues().forEach(function(row){
      var yr=Number(row[0]); if(!yr) return;
      checkIn.push({year:yr,projected:Number(row[1])||0,actual:row[2]?Number(row[2]):null,
        variance:Number(row[3])||null,status:String(row[4]||''),notes:String(row[5]||'')});
    });
  } catch(e){}

  var brackets=[];
  try { tax.getRange('C8:D14').getValues().forEach(function(r){brackets.push({max:Number(r[0]),rate:Number(r[1])});}); } catch(e){}

  var legacyGoal=Number(v('B8')), safeFloor=Number(v('B9')), endLiquid=Number(v('B133'));
  var baseSpend=Number(v('B87')), safeExtra=Number(v('B119'));

  return {
    meta:{ generated:dt(new Date()), planYear:planYear },
    people:{ craig:craig, gena:gena },
    portfolio:{ total:total, peakValue:peakVal*1000, peakYear:peakYear, accounts:accounts },
    debts: (function(){
      // Debt rows: 88,89 then 92-101 (rows 90-91 are headers)
      // Sheet: A=include, B=name, C=monthly, D=annual, H=balance
      var rows88 = inp.getRange('A88:H89').getValues();
      var rows92 = inp.getRange('A92:H101').getValues();
      var result = [];
      rows88.forEach(function(r){
        result.push({inc:String(r[0]||'No'),name:String(r[1]||''),
          mo:Number(r[2])||0,ann:Number(r[3])||0,
          bal:Number(r[7])||0}); // H is index 7
      });
      rows92.forEach(function(r){
        result.push({inc:String(r[0]||'No'),name:String(r[1]||''),
          mo:Number(r[2])||0,ann:Number(r[3])||0,
          bal:Number(r[7])||0}); // H is index 7
      });
      return result;
    })(),
    spending:{
      baseAnnual:baseSpend, safeExtra:safeExtra,
      phase1Start:String(v('C122')), phase1End:String(v('D122')),
      phase2Start:String(v('C123')), phase2End:String(v('D123')),
      phase3Start:String(v('C124')), phase3End:String(v('D124')),
      phase1Weight:Number(v('E122')), phase2Weight:Number(v('E123')), phase3Weight:Number(v('E124')),
      phase1Extra:Number(v('G122')), phase2Extra:Number(v('G123')), phase3Extra:Number(v('G124')),
      phase1Total:baseSpend+Number(v('G122')), phase2Total:baseSpend+Number(v('G123')), phase3Total:baseSpend+Number(v('G124')),
    },
    expenses: (function(){
      // Batch read all expense rows at once
      // Sheet: A=name(editable), B=monthly, C=yearly(formula), D=notes
      var rows=[58,59,60,61,62,63,64,65,66,67,68,69,71,72,73,74,75,76,77,78,79,80,81,83,84,85];
      var aVals = inp.getRange('A58:A85').getValues(); // expense name
      var bVals = inp.getRange('B58:B85').getValues(); // monthly amount
      var dVals = inp.getRange('D58:D85').getValues(); // notes
      return rows.map(function(r){
        var i = r - 58; // index into the arrays
        return {
          row: r,
          name: String(aVals[i] ? aVals[i][0] : ''),
          monthly: Number(bVals[i] ? bVals[i][0] : 0)||0,
          annual: (Number(bVals[i] ? bVals[i][0] : 0)||0)*12,
          note: String(dVals[i] ? dVals[i][0] : '')
        };
      });
    })(),
    legacy:{ goal:legacyGoal, safetyFloor:safeFloor, projectedEnding:endLiquid,
      variance:endLiquid-legacyGoal, achievedPct:endLiquid>=legacyGoal?100:Math.round(endLiquid/legacyGoal*100),
      status:String(v('B134')) },
    tax:{
      effectiveRate:Number(v('B22')), baseTaxAdj:pct(v('B22')), marginalExtra:pct(v('B23')),
      stateRate:pct(v('B26')), standardDeduction:Number(v('B25')),
      rothConversionSavings:Number(v('B146')), inflation:Number(v('B20')), ssCola:Number(v('B21')),
      healthcareInflation:pct(v('B28')), stdDedInflation:pct(v('B27')), brackets:brackets,
    },
    global:{
      planYear:planYear, legacyGoal:Number(v('B8')), safetyFloor:Number(v('B9')),
      survivorReduction:Number(v('B10')), assumeExtraSpend:String(v('B11')),
      stressEarly:String(v('B12')), stressEarlyStart:Number(v('B13')),
      stressEarlyDur:Number(v('B14')), stressEarlyDrag:pct(v('B15')),
      stressLate:String(v('B16')), stressLateStart:Number(v('B17')),
      stressLateDur:Number(v('B18')), stressLateDrag:pct(v('B19')),
      filingStatus:String(v('B24')), standardDeduction:Number(v('B25')),
      projectionMaxYears:Number(v('B7')),
    },
    roth:{ year:Number(v('B139')), bracket:String(v('B141')),
      assumedRate:Number(v('B145')), optimalAmount:Number(v('B143')), taxSavings:Number(v('B146')) },
    rothPlan: (function(){
      // Roth plan: rows 150-169 = years 2026-2045
      // A=year formula, B=partner1 conversion, C=partner2 conversion
      var bVals = inp.getRange('B150:B169').getValues();
      var cVals = inp.getRange('C150:C169').getValues();
      var result = [];
      for (var i=0; i<20; i++) {
        var yr = 2026 + i;
        var p1 = Number(bVals[i][0])||0;
        var p2 = Number(cVals[i][0])||0;
        if (p1 || p2) result.push({year:yr, p1:p1, p2:p2});
      }
      return result;
    })(),
    projections:{ years:years, endLiquid:endLiq, preTax:preTax, roth:roth,
      taxable:taxable, hsa:hsa, income:income, withdrawals:wd },
    checkIn:checkIn,
  };
}

function writeInputs(data, sheetId) {
  try {
    var ss  = SpreadsheetApp.openById(sheetId);
    var inp = ss.getSheetByName('Inputs');

    // FORMULA CELLS — never overwrite these, they are sheet-calculated:
    // B33,D33 (age), B36,D36 (death year), B42,D42 (claiming age),
    // B43,D43 (SS monthly), B45,D45 (SS end), B48,D48 (pension end),
    // B86,B87,B119,B120,B128,B129,B130,B134,B140,B143,B144,B146,B149

    function set(cell, val) { if(val!==undefined && val!==null && val!=='') inp.getRange(cell).setValue(val); }
    function setN(cell,val) { if(val!==undefined) inp.getRange(cell).setValue(Number(val)||0); }
    function setD(cell,val) {
      if(!val || val==='' || val==='0' || val===0) return;
      try {
        var s = String(val).trim();
        var yr, mo, dy;
        // YYYY-MM-DD (HTML date input format) — most reliable
        if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
          var parts = s.split('-');
          yr = parseInt(parts[0]); mo = parseInt(parts[1])-1; dy = parseInt(parts[2]);
        }
        // MM/DD/YYYY or M/D/YYYY
        else if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)) {
          var parts = s.split('/');
          yr = parseInt(parts[2]); mo = parseInt(parts[0])-1; dy = parseInt(parts[1]);
        }
        // MM/DD/YY
        else if (/^\d{1,2}\/\d{1,2}\/\d{2}$/.test(s)) {
          var parts = s.split('/');
          yr = parseInt(parts[2])+2000; mo = parseInt(parts[0])-1; dy = parseInt(parts[1]);
        }
        else {
          var d = new Date(val);
          if(!isNaN(d)) inp.getRange(cell).setValue(d);
          return;
        }
        if(yr && yr > 1900 && yr < 2200) {
          // Use explicit year/month/day to avoid timezone issues with pre-1970 dates
          var d = new Date(yr, mo, dy, 12, 0, 0); // noon to avoid DST shifts
          inp.getRange(cell).setValue(d);
        }
      } catch(e) { Logger.log('setD error for ' + cell + ': ' + e); }
    }
    function setPct(cell,val) { if(val!==undefined) inp.getRange(cell).setValue(Number(val)||0); }

    if (data.global) {
      var g=data.global;
      if(g.B6) {
        var yr = parseInt(g.B6);
        if(!isNaN(yr) && yr > 2000) {
          // Store as Jan 15 to avoid timezone rollback to Dec 31 previous year
          inp.getRange('B6').setValue(new Date(yr, 0, 15));
        }
      }
      setN('B7',g.B7); setN('B8',g.B8); setN('B9',g.B9);
      setN('B10',g.B10); // Survivor reduction stored as decimal e.g. 0.25 set('B11',g.B11); set('B12',g.B12);
      setN('B13',g.B13); setN('B14',g.B14); setPct('B15',g.B15);
      set('B16',g.B16); setN('B17',g.B17); setN('B18',g.B18);
      setPct('B19',g.B19); setPct('B20',g.B20); setPct('B21',g.B21);
      setPct('B22',g.B22); setPct('B23',g.B23);
      set('B24',g.B24); setN('B25',g.B25); setPct('B26',g.B26); setPct('B27',g.B27); setPct('B28',g.B28);
    }
    if (data.partner1) {
      var p=data.partner1;
      set('B31',p.B31); setD('B32',p.B32); set('B34',p.B34);
      setN('B35',p.B35); setN('B37',p.B37); setD('B38',p.B38); setD('B39',p.B39);
      setN('B40',p.B40); setN('B41',p.B41);
      setD('B44',p.B44); setN('B46',p.B46); setD('B47',p.B47);
      setN('B49',p.B49); setD('B50',p.B50); setD('B51',p.B51);
      setN('B52',p.B52); setN('B53',p.B53); setN('B54',p.B54||65);
    }
    if (data.partner2) {
      var p=data.partner2;
      set('D31',p.D31); setD('D32',p.D32); set('D34',p.D34);
      setN('D35',p.D35); setN('D37',p.D37); setD('D38',p.D38); setD('D39',p.D39);
      setN('D40',p.D40); setN('D41',p.D41);
      setD('D44',p.D44); setN('D46',p.D46); setD('D47',p.D47);
      setN('D49',p.D49); setD('D50',p.D50); setD('D51',p.D51);
      setN('D52',p.D52); setN('D53',p.D53); setN('D54',p.D54||65);
    }
    // Flush partner names first
    SpreadsheetApp.flush();

    if (data.accounts && data.accounts.length) {
      var p1name = data.partner1 && data.partner1.B31 ? String(data.partner1.B31).trim() : '';
      var p2name = data.partner2 && data.partner2.D31 ? String(data.partner2.D31).trim() : '';

      // Temporarily remove validation on owner column so any value can be written
      inp.getRange('C105:C116').clearDataValidations();

      // Build batch array for accounts — 1 write instead of 96
      var acctData = [];
      for (var ai=0; ai<12; ai++) {
        var a = data.accounts[ai] || {};
        var owner = String(a.owner||'Joint').trim();
        if (owner==='Partner 1'||owner==='Craig') owner = p1name||'Joint';
        if (owner==='Partner 2'||owner==='Gena')  owner = p2name||'Joint';
        if (!owner) owner = 'Joint';
        acctData.push([
          a.inc||a.showInCalc||'No',
          a.name||'',
          owner,
          a.type||'',
          Number(a.bal||a.balance)||0,
          Number(a.ret||a.expectedReturn)||0,
          a.status||'Use for Withdrawals',
          a.dash||a.showOnDashboard||'No'
        ]);
      }
      inp.getRange('A105:H116').setValues(acctData);

      // Restore validation using the dynamic range from Technical Style Reference
      var ownerValidation = SpreadsheetApp.newDataValidation()
        .requireValueInRange(
          ss.getSheetByName('Technical Style Reference').getRange('F6:F8'), true)
        .setAllowInvalid(false)
        .build();
      inp.getRange('C105:C116').setDataValidation(ownerValidation);
    }
    if (data.roth) {
      // B139 = Roth year (user input), B141 = bracket (user input), B145 = rate (user input)
      // B140,B143,B144,B146 are formula cells — do NOT write
      setN('B139',data.roth.B139); set('B141',data.roth.B141); setPct('B145',data.roth.B145);
    }
    // Write Roth conversion plan amounts (rows 150-169 = years 2026-2045)
    if (data.rothPlan && data.rothPlan.length) {
      var p1Data = [], p2Data = [];
      for (var ri=0; ri<20; ri++) { p1Data.push([0]); p2Data.push([0]); }
      data.rothPlan.forEach(function(row) {
        var idx = Number(row.year) - 2026;
        if (idx >= 0 && idx < 20) {
          p1Data[idx] = [Number(row.p1)||0];
          p2Data[idx] = [Number(row.p2)||0];
        }
      });
      inp.getRange('B150:B169').setValues(p1Data);
      inp.getRange('C150:C169').setValues(p2Data);
    }
    // Write debt rows
    // Sheet structure: A=include, B=name, C=monthly, D=annual(=C*12), H=balance
    // Rows 88-89 are main debt rows, 90-91 are headers, 92-101 are extended debts
    if (data.debts && data.debts.length) {
      var debtSheetRows = [88, 89, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101];
      data.debts.forEach(function(d, i) {
        if (i >= debtSheetRows.length) return;
        var r = debtSheetRows[i];
        var mo = Number(d.mo) || 0;
        var bal = Number(d.bal) || 0;
        inp.getRange('A'+r).setValue(d.inc || 'No');
        if (d.name) inp.getRange('B'+r).setValue(d.name);
        inp.getRange('C'+r).setValue(mo);
        inp.getRange('D'+r).setValue(mo * 12);
        if (d.start) { try { var ds=d.start.split('-'); if(ds.length===3) inp.getRange('E'+r).setValue(new Date(parseInt(ds[0]),parseInt(ds[1])-1,parseInt(ds[2]),12,0,0)); } catch(e){} }
        if (d.end)   { try { var de=d.end.split('-');   if(de.length===3) inp.getRange('F'+r).setValue(new Date(parseInt(de[0]),parseInt(de[1])-1,parseInt(de[2]),12,0,0)); } catch(e){} }
        if (bal) inp.getRange('H'+r).setValue(bal);
      });
    }
    // Write expense rows — only rows with data
    if (data.expenses && data.expenses.length) {
      var hasExpenses = data.expenses.some(function(e){ return e.name || e.monthly; });
      if (hasExpenses) {
        var bData = [], cData = [], tData = []; // b=name, c=monthly, t=notes
        for (var i=0; i<28; i++) {
          bData.push(['']); cData.push([0]); tData.push(['']);
        }
        data.expenses.forEach(function(exp) {
          var r = Number(exp.row);
          if (!r || r < 58 || r > 85) return;
          var idx = r - 58;
          bData[idx] = [exp.name||''];
          cData[idx] = [Number(exp.monthly)||0];
          tData[idx] = [exp.note||''];
        });
        // B column written row by row below
        // Write expenses — skip formula rows 62 and 70
      // Sheet: A=name, B=monthly, C=yearly(formula-never write), D=notes
      var expRows = [58,59,60,61,63,64,65,66,67,68,69,71,72,73,74,75,76,77,78,79,80,81,83,84,85];
      expRows.forEach(function(r) {
        var idx = r - 58;
        if(bData[idx][0] !== undefined) inp.getRange('A'+r).setValue(bData[idx][0]); // name → A
        var mo = Number(cData[idx] ? cData[idx][0] : 0) || 0;
        inp.getRange('B'+r).setValue(mo); // monthly → B
        inp.getRange('C'+r).setValue(mo * 12); // yearly → C (in case formula was lost)
        if(tData[idx][0] !== undefined) inp.getRange('D'+r).setValue(tData[idx][0]); // notes → D
      });
        // T column written row by row below
      }
    }
    SpreadsheetApp.flush();
    return {success:true, timestamp:new Date().toISOString()};
  } catch(err) {
    return {success:false, error:err.toString()};
  }
}
