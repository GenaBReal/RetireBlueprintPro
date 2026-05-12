// RetireBlueprint Pro — Complete Apps Script (READ + WRITE)
// Deploy as Web App: Execute as Me, Who has access: Anyone
// This ONE deployment serves ALL customers — each passes their own Sheet ID

var FALLBACK_SHEET_ID = '1wWzDjO2VtKb5EEQ0xolKdd1Je961TCqGDI8i-ab3Dfc'; // Your master template

function doGet(e) {
  var callback = e.parameter.callback || 'callback';
  var action   = e.parameter.action   || 'read';
  // Accept sheet ID from the request — falls back to your master if not provided
  var sheetId  = e.parameter.sheetId  || FALLBACK_SHEET_ID;
  try {
    if (action === 'save') {
      var data = JSON.parse(decodeURIComponent(e.parameter.data || '{}'));
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
    name:v('B31'), age:Number(v('B33')), projectedDeathAge:Number(v('B35')),
    projectedDeathYear:Number(v('B36')), salary:Number(v('B37')),
    ssMonthly:Number(v('B43')), ssAnnual:Number(v('B43'))*12,
    ssStartDate:dt(v('B44')), pension:Number(v('B46')), pensionAnnual:Number(v('B46'))*12,
    otherIncome:Number(v('B49')), theme:String(v('B34')),
  };
  var gena = {
    name:v('D31'), age:Number(v('D33')), projectedDeathAge:Number(v('D35')),
    projectedDeathYear:Number(v('D36')), salary:Number(v('D37')),
    ssMonthly:Number(v('D43')), ssAnnual:Number(v('D43'))*12,
    ssStartDate:dt(v('D44')), pension:Number(v('D46')), pensionAnnual:0,
    otherIncome:Number(v('D49')), theme:String(v('D34')),
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
    spending:{
      baseAnnual:baseSpend, safeExtra:safeExtra,
      phase1Extra:Number(v('G122')), phase2Extra:Number(v('G123')), phase3Extra:Number(v('G124')),
      phase1Total:baseSpend+Number(v('G122')), phase2Total:baseSpend+Number(v('G123')), phase3Total:baseSpend+Number(v('G124')),
    },
    legacy:{ goal:legacyGoal, safetyFloor:safeFloor, projectedEnding:endLiquid,
      variance:endLiquid-legacyGoal, achievedPct:endLiquid>=legacyGoal?100:Math.round(endLiquid/legacyGoal*100),
      status:String(v('B134')) },
    tax:{
      effectiveRate:Number(v('B22')), stateRate:pct(v('B26')), standardDeduction:Number(v('B25')),
      rothConversionSavings:Number(v('B146')), inflation:Number(v('B20')), ssCola:Number(v('B21')),
      healthcareInflation:pct(v('B28')), brackets:brackets,
    },
    roth:{ year:Number(v('B139')), bracket:String(v('B141')),
      optimalAmount:Number(v('B143')), taxSavings:Number(v('B146')) },
    projections:{ years:years, endLiquid:endLiq, preTax:preTax, roth:roth,
      taxable:taxable, hsa:hsa, income:income, withdrawals:wd },
    checkIn:checkIn,
  };
}

function writeInputs(data, sheetId) {
  try {
    var ss  = SpreadsheetApp.openById(sheetId);
    var inp = ss.getSheetByName('Inputs');
    function set(cell, val) { if(val!==undefined && val!==null && val!=='') inp.getRange(cell).setValue(val); }
    function setN(cell,val) { if(val!==undefined) inp.getRange(cell).setValue(Number(val)||0); }
    function setD(cell,val) { if(val) { var d=new Date(val); if(!isNaN(d)) inp.getRange(cell).setValue(d); } }
    function setPct(cell,val) { if(val!==undefined) inp.getRange(cell).setValue(Number(val)||0); }

    if (data.global) {
      var g=data.global;
      if(g.B6) inp.getRange('B6').setValue(new Date(g.B6+'-01-01'));
      setN('B7',g.B7); setN('B8',g.B8); setN('B9',g.B9);
      setPct('B10',g.B10); set('B11',g.B11); set('B12',g.B12);
      setN('B13',g.B13); setN('B14',g.B14); setPct('B15',g.B15);
      set('B16',g.B16); setN('B17',g.B17); setN('B18',g.B18);
      setPct('B19',g.B19); setPct('B20',g.B20); setPct('B21',g.B21);
      set('B24',g.B24); setN('B25',g.B25); setPct('B26',g.B26); setPct('B28',g.B28);
    }
    if (data.partner1) {
      var p=data.partner1;
      set('B31',p.B31); setD('B32',p.B32); set('B34',p.B34);
      setN('B35',p.B35); setN('B37',p.B37); setD('B39',p.B39);
      setN('B40',p.B40); setN('B41',p.B41); setN('B43',p.B43);
      setD('B44',p.B44); setN('B46',p.B46); setD('B47',p.B47);
      setN('B49',p.B49); setD('B50',p.B50); setD('B51',p.B51);
      setN('B52',p.B52); setN('B53',p.B53);
    }
    if (data.partner2) {
      var p=data.partner2;
      set('D31',p.D31); setD('D32',p.D32); set('D34',p.D34);
      setN('D35',p.D35); setN('D37',p.D37); setD('D39',p.D39);
      setN('D40',p.D40); setN('D41',p.D41); setN('D43',p.D43);
      setD('D44',p.D44); setN('D46',p.D46); setD('D47',p.D47);
      setN('D49',p.D49); setD('D50',p.D50); setD('D51',p.D51);
      setN('D52',p.D52); setN('D53',p.D53);
    }
    if (data.accounts && data.accounts.length) {
      data.accounts.forEach(function(a,i){
        var r=105+i; if(r>116) return;
        inp.getRange('A'+r).setValue(a.inc||a.showInCalc||'No');
        inp.getRange('B'+r).setValue(a.name||'');
        inp.getRange('C'+r).setValue(a.owner||'Joint');
        inp.getRange('D'+r).setValue(a.type||'');
        inp.getRange('E'+r).setValue(Number(a.bal||a.balance)||0);
        inp.getRange('F'+r).setValue(Number(a.ret||a.expectedReturn)||0);
        inp.getRange('G'+r).setValue(a.status||'Use for Withdrawals');
        inp.getRange('H'+r).setValue(a.dash||a.showOnDashboard||'No');
      });
    }
    if (data.roth) {
      setN('B139',data.roth.B139); set('B141',data.roth.B141); setPct('B145',data.roth.B145);
    }
    SpreadsheetApp.flush();
    return {success:true, timestamp:new Date().toISOString()};
  } catch(err) {
    return {success:false, error:err.toString()};
  }
}
