// RetireBlueprint Pro — Master Formatter v5
// SPLIT INTO 3 PARTS to avoid Google's 6-minute timeout
// 
// HOW TO RUN:
// 1. Paste this into FormatMaster.gs and save
// 2. Run formatPart1 first — wait for the alert saying Part 1 done
// 3. Then run formatPart2 — wait for alert
// 4. Then run formatPart3 — wait for alert
// Done! All 3 parts together = fully formatted Master tab
// ─────────────────────────────────────────────────────────

// ── SHARED SETUP — reads colors and plan end from Inputs ─────────
function getConfig() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var inp = ss.getSheetByName('Inputs');
  var mst = ss.getSheetByName('Master');

  var p1Theme = String(inp.getRange('B34').getValue() || 'Steel Blue');
  var p2Theme = String(inp.getRange('D34').getValue() || 'Frost Plum');
  var themes = {
    'Steel Blue':   {hdr:'2E4053',stripe:'F4F7FA',dot:'E8EEF3'},
    'Cool Slate':   {hdr:'2E4053',stripe:'F4F7FA',dot:'E8EEF3'},
    'Frost Plum':   {hdr:'6A4A6D',stripe:'FAF5FB',dot:'F0E6F2'},
    'Forest Green': {hdr:'2E7032',stripe:'F3FAF4',dot:'E2F0E3'},
    'Warm Amber':   {hdr:'C9A040',stripe:'FDFAF2',dot:'F5EDD4'},
  };

  // Read projected death years from Inputs (B36 = P1, D36 = P2)
  var p1DeathYear = Number(inp.getRange('B36').getValue()) || 2055;
  var p2DeathYear = Number(inp.getRange('D36').getValue()) || 2075;
  var planEndYear = Math.max(p1DeathYear, p2DeathYear);

  // Find the plan start year from Inputs B6
  var planStartYear = new Date(inp.getRange('B6').getValue()).getFullYear() || 2026;

  // Calculate last data row (row 8 = plan start year, each row = 1 year)
  var lastPlanRow = 8 + (planEndYear - planStartYear);
  lastPlanRow = Math.min(lastPlanRow, 160); // safety cap

  // Also find where the first partner dies (for partial coloring if needed)
  var p1DeathRow = 8 + (p1DeathYear - planStartYear);
  var p2DeathRow = 8 + (p2DeathYear - planStartYear);

  return {
    p1Name:  String(inp.getRange('B31').getValue() || 'Partner 1'),
    p2Name:  String(inp.getRange('D31').getValue() || 'Partner 2'),
    p1Theme: p1Theme,
    p2Theme: p2Theme,
    p1C: themes[p1Theme] || themes['Steel Blue'],
    p2C: themes[p2Theme] || themes['Frost Plum'],
    NAVY:'041E2F', NAVY2:'0A3145', TEAL:'1CC7D0', ORANGE:'F4B86A',
    WHITE:'FFFFFF', OFFWHITE:'F8F9FA',
    JOINT_H:'1E4D3A', JOINT_S:'F3FAF7', JOINT_D:'E4F2EC',
    TOTAL_S:'F4F7FA', TOTAL_D:'E8EEF4',
    P1: [3,5,6,7,8,9,29,32,42,43,44,55,56,57],
    P2: [4,10,11,12,13,14,30,33,45,46,47,58,59,60],
    JT: [48,49,50,51,52,53,61,62,63,64,65,66],
    TOT:[15,24,28,31,34,39,40,54,67,73,74,75,76],
    planEndYear:  planEndYear,
    lastPlanRow:  lastPlanRow,
    p1DeathRow:   p1DeathRow,
    p2DeathRow:   p2DeathRow,
    planStartYear: planStartYear,
  };
}

function L(letter) {
  var col=0;
  for(var i=0;i<letter.length;i++) col=col*26+letter.charCodeAt(i)-64;
  return col;
}

// ── PART 1: Headers (rows 5 and 6) ───────────────────────────────
function formatPart1() {
  var cfg = getConfig();
  var mst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master');
  var lastCol = mst.getLastColumn();

  SpreadsheetApp.getActiveSpreadsheet().toast('Part 1 of 3 — Formatting headers...','RetireBlueprint Pro',5);

  // Clear formats rows 5-7
  mst.getRange(5,1,3,lastCol).clearFormat();
  mst.getRange(5,1,1,lastCol).clearContent();

  // Row 5 — section group headers
  var sections = [
    {col:1, end:4,  label:'LIFETIME PLAN',                            bg:cfg.NAVY,  fg:cfg.ORANGE},
    {col:5, end:16, label:'GUARANTEED INCOME — THE HAVE',             bg:cfg.NAVY2, fg:cfg.ORANGE},
    {col:17,end:24, label:'ANNUAL BASE EXPENSES — THE NEED',          bg:cfg.NAVY2, fg:cfg.ORANGE},
    {col:25,end:41, label:'REALITY CHECK — GAP, TAXES & WITHDRAWALS', bg:cfg.NAVY,  fg:cfg.ORANGE},
    {col:42,end:54, label:'TAX-OPTIMIZED WITHDRAWALS',                bg:cfg.NAVY2, fg:cfg.ORANGE},
    {col:55,end:66, label:'ACCOUNT BALANCES (POST GROWTH)',           bg:cfg.NAVY2, fg:cfg.ORANGE},
    {col:67,end:76, label:'YEARLY PORTFOLIO TOTALS',                  bg:cfg.NAVY,  fg:cfg.ORANGE},
  ];
  sections.forEach(function(s) {
    var end = Math.min(s.end, lastCol);
    var r   = mst.getRange(5,s.col,1,end-s.col+1);
    try { r.mergeAcross(); } catch(e) {}
    r.setBackground('#'+s.bg).setFontColor('#'+s.fg)
     .setFontSize(8).setFontWeight('bold')
     .setHorizontalAlignment('center').setVerticalAlignment('middle');
    mst.getRange(5,s.col).setValue(s.label);
  });
  mst.setRowHeight(5,24);

  // Row 6 — column headers: base style
  mst.setRowHeight(6,56);
  mst.getRange(6,1,1,lastCol)
     .setFontSize(8).setFontWeight('bold').setWrap(true)
     .setHorizontalAlignment('center').setVerticalAlignment('middle')
     .setBackground('#'+cfg.NAVY2).setFontColor('#'+cfg.TEAL);

  // Row 6 — colored thick bottom border + background per group
  var thick = SpreadsheetApp.BorderStyle.SOLID_THICK;

  function hdrStyle(colNum, bgHex, fgHex, borderHex) {
    if (colNum > lastCol) return;
    mst.getRange(6,colNum)
       .setBackground('#'+bgHex).setFontColor('#'+fgHex)
       .setBorder(null,null,true,null,null,null,'#'+borderHex,thick);
  }

  // Neutral teal border for all first
  for (var c=1;c<=lastCol;c++) {
    mst.getRange(6,c).setBorder(null,null,true,null,null,null,'#'+cfg.TEAL,thick);
  }
  // Partner 1
  cfg.P1.forEach(function(c){ hdrStyle(c,cfg.p1C.hdr,cfg.WHITE,cfg.p1C.hdr); });
  // Partner 2
  cfg.P2.forEach(function(c){ hdrStyle(c,cfg.p2C.hdr,cfg.WHITE,cfg.p2C.hdr); });
  // Joint
  cfg.JT.forEach(function(c){ hdrStyle(c,cfg.JOINT_H,cfg.WHITE,cfg.JOINT_H); });
  // Totals
  cfg.TOT.forEach(function(c){ hdrStyle(c,cfg.NAVY,cfg.ORANGE,cfg.ORANGE); });
  // Year & Date
  hdrStyle(1,cfg.NAVY,cfg.ORANGE,cfg.ORANGE);
  hdrStyle(2,cfg.NAVY,cfg.TEAL,cfg.TEAL);

  // Freeze rows
  mst.setFrozenRows(6);

  // Column widths
  mst.setColumnWidth(1,52); mst.setColumnWidth(2,72);
  mst.setColumnWidth(3,62); mst.setColumnWidth(4,62);
  for (var col=5;col<=lastCol;col++) mst.setColumnWidth(col,105);

  SpreadsheetApp.getUi().alert('✅ Part 1 done — headers formatted!\n\nPlan runs to ' + cfg.planEndYear + ' (row ' + cfg.lastPlanRow + ')\n\nNow select formatPart2 from the dropdown and click Run.');
}

// ── PART 2: Data rows 8–85 (first half) ──────────────────────────
function formatPart2() {
  var cfg = getConfig();
  var mst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master');
  var lastCol = mst.getLastColumn();
  var startRow = 8;
  var endRow   = Math.min(85, cfg.lastPlanRow);
  var numRows  = endRow - startRow + 1;
  if (numRows < 1) { SpreadsheetApp.getUi().alert('Part 2 skipped — plan ends before row 85.'); return; }

  SpreadsheetApp.getActiveSpreadsheet().toast('Part 2 of 3 — Data rows 8–85...','RetireBlueprint Pro',5);

  // Build background grid for this chunk
  var bgGrid = [];
  for (var row=startRow; row<=endRow; row++) {
    var isAlt = (row%2===0);
    var rowBgs = [];
    for (var col=1; col<=lastCol; col++) {
      var base = isAlt ? '#'+cfg.OFFWHITE : '#'+cfg.WHITE;
      if (cfg.P1.indexOf(col)>-1)  base = isAlt ? '#'+cfg.p1C.dot    : '#'+cfg.p1C.stripe;
      if (cfg.P2.indexOf(col)>-1)  base = isAlt ? '#'+cfg.p2C.dot    : '#'+cfg.p2C.stripe;
      if (cfg.JT.indexOf(col)>-1)  base = isAlt ? '#'+cfg.JOINT_D    : '#'+cfg.JOINT_S;
      if (cfg.TOT.indexOf(col)>-1) base = isAlt ? '#'+cfg.TOTAL_D    : '#'+cfg.TOTAL_S;
      rowBgs.push(base);
    }
    bgGrid.push(rowBgs);
  }

  var rng = mst.getRange(startRow,1,numRows,lastCol);
  rng.setBackgrounds(bgGrid);
  rng.setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
  rng.setBorder(true,true,true,true,true,true,'#DDE3E9',SpreadsheetApp.BorderStyle.SOLID);
  rng.setNumberFormat('$#,##0');
  mst.getRange(startRow,1,numRows,1).setNumberFormat('0');
  mst.getRange(startRow,2,numRows,1).setNumberFormat('m/d/yyyy');
  mst.getRange(startRow,3,numRows,2).setNumberFormat('0');
  mst.getRange(startRow,18,numRows,1).setNumberFormat('0.0%');
  mst.getRange(startRow,21,numRows,1).setNumberFormat('0.0%');

  SpreadsheetApp.getUi().alert('✅ Part 2 done — rows 8–85 formatted!\n\nNow select formatPart3 from the dropdown and click Run.');
}

// ── PART 3: Data rows 86–160 + section borders ───────────────────
function formatPart3() {
  var cfg = getConfig();
  var mst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master');
  var lastCol = mst.getLastColumn();
  var lastRow  = cfg.lastPlanRow;
  var startRow = 86;
  var numRows  = lastRow - startRow + 1;

  SpreadsheetApp.getActiveSpreadsheet().toast('Part 3 of 3 — Data rows 86–end + borders...','RetireBlueprint Pro',5);

  if (numRows > 0) {
    var bgGrid = [];
    for (var row=startRow; row<=lastRow; row++) {
      var isAlt = (row%2===0);
      var rowBgs = [];
      for (var col=1; col<=lastCol; col++) {
        var base = isAlt ? '#'+cfg.OFFWHITE : '#'+cfg.WHITE;
        if (cfg.P1.indexOf(col)>-1)  base = isAlt ? '#'+cfg.p1C.dot    : '#'+cfg.p1C.stripe;
        if (cfg.P2.indexOf(col)>-1)  base = isAlt ? '#'+cfg.p2C.dot    : '#'+cfg.p2C.stripe;
        if (cfg.JT.indexOf(col)>-1)  base = isAlt ? '#'+cfg.JOINT_D    : '#'+cfg.JOINT_S;
        if (cfg.TOT.indexOf(col)>-1) base = isAlt ? '#'+cfg.TOTAL_D    : '#'+cfg.TOTAL_S;
        rowBgs.push(base);
      }
      bgGrid.push(rowBgs);
    }
    var rng = mst.getRange(startRow,1,numRows,lastCol);
    rng.setBackgrounds(bgGrid);
    rng.setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
    rng.setBorder(true,true,true,true,true,true,'#DDE3E9',SpreadsheetApp.BorderStyle.SOLID);
    rng.setNumberFormat('$#,##0');
    mst.getRange(startRow,1,numRows,1).setNumberFormat('0');
    mst.getRange(startRow,2,numRows,1).setNumberFormat('m/d/yyyy');
    mst.getRange(startRow,3,numRows,2).setNumberFormat('0');
    mst.getRange(startRow,18,numRows,1).setNumberFormat('0.0%');
    mst.getRange(startRow,21,numRows,1).setNumberFormat('0.0%');
  }

  // Section boundary borders (orange medium lines)
  var med = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  [5,17,25,42,55,67].forEach(function(col) {
    if (col<=lastCol) {
      mst.getRange(5,col,cfg.lastPlanRow-4,1)
         .setBorder(null,true,null,null,null,null,'#'+cfg.ORANGE,med);
    }
  });

  // Clear ALL formatting on rows AFTER plan end
  var totalMasterRows = 200; // clear well beyond any possible plan end
  if (cfg.lastPlanRow < totalMasterRows) {
    var clearStart = cfg.lastPlanRow + 1;
    var clearCount = totalMasterRows - clearStart + 1;
    if (clearCount > 0) {
      var clearRng = mst.getRange(clearStart, 1, clearCount, lastCol);
      clearRng.clearFormat();
      clearRng.setBackground('#FFFFFF');
      // Explicitly remove all borders
      clearRng.setBorder(false, false, false, false, false, false,
                         null, SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  SpreadsheetApp.getUi().alert(
    '✅ ALL DONE! Master tab fully formatted!\n\n' +
    cfg.p1Name + ' columns → ' + cfg.p1Theme + '\n' +
    cfg.p2Name + ' columns → ' + cfg.p2Theme + '\n\n' +
    'Plan ends: ' + cfg.planEndYear + ' (row ' + cfg.lastPlanRow + ')\n' +
    'Rows after plan end are cleared to plain white.\n\n' +
    'Close and reopen your sheet to see the RetireBlueprint Pro menu.'
  );
}

// ── MENU ─────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('RetireBlueprint Pro')
    .addItem('Step 1 — Format Headers',    'formatPart1')
    .addItem('Step 2 — Format Rows 8–85',  'formatPart2')
    .addItem('Step 3 — Format Rows 86+',   'formatPart3')
    .addToUi();
}
