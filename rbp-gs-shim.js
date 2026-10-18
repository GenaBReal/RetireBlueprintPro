/* ─────────────────────────────────────────────────────────────────────────────
   RetireBlueprint Pro — Google Sheets adapter for the browser (Google sign-in path)

   The connector logic (rbp-gs-core.js) was written for Google Apps Script, where code
   reads and writes a sheet with calls like inp.getRange('B8').getValue(). This file
   provides those same calls in the customer's browser:

     1. RBPShim.load(sheetId, opts)  — fetches the needed tabs from Google in one go
        (values exactly as Apps Script would return them: numbers, text, true/false,
        and real Date objects for date-formatted cells).
     2. The connector code then runs unchanged against that in-memory copy.
        Every write is recorded instead of sent.
     3. book.flush()                  — sends all recorded changes to Google in order:
        structure/format changes first, then cell values.

   Nothing here talks to RetireBlueprint Pro — only to Google, with the customer's own sign-in.
   ───────────────────────────────────────────────────────────────────────────── */
(function (root) {
  var G = function () { return root.RBPGoogle; };

  // ── A1 helpers ────────────────────────────────────────────────────────────
  function colNum(letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64); return n; }
  function colLetters(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function parseA1(a1) {
    var m = String(a1).toUpperCase().match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!m) throw new Error('Bad range ' + a1);
    var r1 = +m[2], c1 = colNum(m[1]), r2 = m[3] ? +m[4] : r1, c2 = m[3] ? colNum(m[3]) : c1;
    return { r: Math.min(r1, r2), c: Math.min(c1, c2), nr: Math.abs(r2 - r1) + 1, nc: Math.abs(c2 - c1) + 1 };
  }
  function q(title) { return "'" + String(title).replace(/'/g, "''") + "'"; }

  // ── value conversion (Google API ⇄ Apps Script) ────────────────────────────
  function serialToDate(n) {
    var days = Math.floor(n), ms = Math.round((n - days) * 86400000);
    return new Date(1899, 11, 30 + days, 0, 0, 0, ms);      // local components, like Apps Script shows them
  }
  function cellToValue(cell) {
    if (!cell) return '';
    var ev = cell.effectiveValue;
    if (!ev) return (cell.formattedValue && /^#/.test(cell.formattedValue)) ? cell.formattedValue : '';
    if (ev.errorValue) return cell.formattedValue || '#ERROR!';
    if (ev.boolValue !== undefined) return ev.boolValue;
    if (ev.stringValue !== undefined) return ev.stringValue;
    if (ev.numberValue !== undefined) {
      var t = cell.effectiveFormat && cell.effectiveFormat.numberFormat && cell.effectiveFormat.numberFormat.type;
      if (t === 'DATE' || t === 'DATE_TIME' || t === 'TIME') return serialToDate(ev.numberValue);
      return ev.numberValue;
    }
    return '';
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function valueToApi(v) {
    if (v instanceof Date) {
      if (isNaN(v)) return '';
      return v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate()) + ' ' +
             pad2(v.getHours()) + ':' + pad2(v.getMinutes()) + ':' + pad2(v.getSeconds());
    }
    if (v === null || v === undefined) return '';
    if (typeof v === 'number') return isFinite(v) ? v : '';
    return v;                       // strings & booleans as typed (USER_ENTERED, same as Apps Script setValue)
  }

  // ── Workbook ──────────────────────────────────────────────────────────────
  function Book(id, meta) {
    this.id = id;
    this.tz = (meta.properties && meta.properties.timeZone) || 'America/Los_Angeles';
    var ic = meta.properties && meta.properties.iterativeCalculationSettings;
    this.iterative = !!ic;
    this.sheets = {};
    this.ops = [];          // ordered structural requests
    this.writes = {};       // 'Sheet!A1' -> value (last write wins)
    this.writeOrder = [];
    this.meta = (meta.developerMetadata || []).map(function (d) { return { id: d.metadataId, key: d.metadataKey, value: d.metadataValue }; });
    var self = this;
    (meta.sheets || []).forEach(function (s) {
      var p = s.properties || {};
      self.sheets[p.title] = new Sheet(self, p);
    });
  }
  Book.prototype.getId = function () { return this.id; };
  Book.prototype.getSheetByName = function (n) { return this.sheets[n] || null; };
  Book.prototype.getSpreadsheetTimeZone = function () { return this.tz; };
  Book.prototype.isIterativeCalculationEnabled = function () { return this.iterative; };
  Book.prototype.setIterativeCalculationEnabled = function (on) {
    if (!on) return;
    this.iterative = true;
    this.ops.push({ updateSpreadsheetProperties: {
      properties: { iterativeCalculationSettings: { maxIterations: 50, convergenceThreshold: 0.05 } },
      fields: 'iterativeCalculationSettings' } });
  };
  Book.prototype.insertSheet = function (name) {
    var sid = Math.floor(100000 + Math.random() * 1e9);
    this.ops.push({ addSheet: { properties: { sheetId: sid, title: name } } });
    var sh = new Sheet(this, { sheetId: sid, title: name, gridProperties: { rowCount: 1000, columnCount: 26 } });
    sh.loaded = true;
    this.sheets[name] = sh;
    return sh;
  };
  Book.prototype.getDeveloperMetadata = function () {
    var self = this;
    return this.meta.filter(function (m) { return !m.removed; }).map(function (m) {
      return { getKey: function () { return m.key; }, getValue: function () { return m.value; },
               remove: function () { m.removed = true; if (m.id) self.ops.push({ deleteDeveloperMetadata: { dataFilter: { developerMetadataLookup: { metadataId: m.id } } } }); } };
    });
  };
  Book.prototype.addDeveloperMetadata = function (key, value) {
    this.meta.push({ key: key, value: value });
    this.ops.push({ createDeveloperMetadata: { developerMetadata: { metadataKey: key, metadataValue: value, location: { spreadsheet: true }, visibility: 'DOCUMENT' } } });
  };
  Book.prototype.hasChanges = function () { return this.ops.length > 0 || this.writeOrder.length > 0; };
  // Send everything recorded so far to Google: structure first, then values.
  Book.prototype.flush = function () {
    var self = this, ops = this.ops.slice(), order = this.writeOrder.slice(), writes = this.writes;
    this.ops = []; this.writeOrder = []; this.writes = {};
    var data = order.map(function (k) { return { range: k, values: [[valueToApi(writes[k])]] }; });
    return G().batchUpdate(self.id, ops).then(function () { return G().valuesBatchUpdate(self.id, data); });
  };

  // ── Sheet ─────────────────────────────────────────────────────────────────
  function Sheet(book, p) {
    this.book = book;
    this.title = p.title;
    this.sid = p.sheetId;
    this.hidden = !!p.hidden;
    this.maxRows = (p.gridProperties && p.gridProperties.rowCount) || 1000;
    this.maxCols = (p.gridProperties && p.gridProperties.columnCount) || 26;
    this.cells = {};          // 'r,c' -> {v, f}
    this.loaded = false;
    this.window = null;       // {r1,c1,r2,c2} loaded area (null = whole sheet)
  }
  Sheet.prototype.getName = function () { return this.title; };
  Sheet.prototype._get = function (r, c) { return this.cells[r + ',' + c]; };
  Sheet.prototype._check = function (r, c) {
    if (!this.loaded) throw new Error('Tab ' + this.title + ' was not loaded');
    var w = this.window;
    if (w && (r < w.r1 || r > w.r2 || c < w.c1 || c > w.c2)) throw new Error('Read outside loaded area: ' + this.title + ' ' + colLetters(c) + r);
  };
  Sheet.prototype.getRange = function (a, b, c, d) {
    if (typeof a === 'string') { var p = parseA1(a); return new Range(this, p.r, p.c, p.nr, p.nc); }
    return new Range(this, a, b, c || 1, d || 1);
  };
  Sheet.prototype.getLastRow = function () {
    var last = 0;
    for (var k in this.cells) { var x = this.cells[k]; if (x.v !== '' || x.f) { var r = +k.split(',')[0]; if (r > last) last = r; } }
    return last;
  };
  Sheet.prototype.getLastColumn = function () {
    var last = 0;
    for (var k in this.cells) { var x = this.cells[k]; if (x.v !== '' || x.f) { var c = +k.split(',')[1]; if (c > last) last = c; } }
    return last;
  };
  Sheet.prototype.getDataRange = function () { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); };
  Sheet.prototype.getMaxRows = function () { return this.maxRows; };
  Sheet.prototype.deleteRows = function (start, n) {
    this.book.ops.push({ deleteDimension: { range: { sheetId: this.sid, dimension: 'ROWS', startIndex: start - 1, endIndex: start - 1 + n } } });
    this.maxRows -= n;
  };
  Sheet.prototype.hideSheet = function () {
    if (this.hidden) return;
    this.hidden = true;
    this.book.ops.push({ updateSheetProperties: { properties: { sheetId: this.sid, hidden: true }, fields: 'hidden' } });
  };

  // ── Range ─────────────────────────────────────────────────────────────────
  function Range(sheet, r, c, nr, nc) { this.sh = sheet; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  Range.prototype._a1 = function () {
    var a = colLetters(this.c) + this.r;
    if (this.nr > 1 || this.nc > 1) a += ':' + colLetters(this.c + this.nc - 1) + (this.r + this.nr - 1);
    return a;
  };
  Range.prototype._grid = function (pick) {
    var out = [];
    for (var i = 0; i < this.nr; i++) {
      var row = [];
      for (var j = 0; j < this.nc; j++) {
        this.sh._check(this.r + i, this.c + j);
        row.push(pick(this.sh._get(this.r + i, this.c + j)));
      }
      out.push(row);
    }
    return out;
  };
  // Fresh read straight from Google (after sending any pending writes) — used by the stress
  // tests, which must see the sheet's recalculated numbers, not the copy loaded earlier.
  Range.prototype.getValuesAsync = function () {
    var self = this, book = this.sh.book, a1 = q(this.sh.title) + '!' + this._a1();
    var fields = 'sheets(data(startRow,startColumn,rowData(values(effectiveValue,formattedValue,userEnteredValue/formulaValue,effectiveFormat/numberFormat/type))))';
    return book.flush().then(function () {
      return G().getSpreadsheet(book.id, 'ranges=' + encodeURIComponent(a1) + '&fields=' + encodeURIComponent(fields));
    }).then(function (j) {
      var out = [], rows = [];
      try { rows = j.sheets[0].data[0].rowData || []; } catch (e) {}
      for (var i = 0; i < self.nr; i++) {
        var row = [], vals = (rows[i] && rows[i].values) || [];
        for (var k = 0; k < self.nc; k++) {
          var cell = vals[k], v = cellToValue(cell), f = (cell && cell.userEnteredValue && cell.userEnteredValue.formulaValue) || '';
          row.push(v);
          self.sh.cells[(self.r + i) + ',' + (self.c + k)] = { v: v, f: f };   // keep the local copy current too
        }
        out.push(row);
      }
      return out;
    });
  };
  Range.prototype.getValueAsync = function () { return this.getValuesAsync().then(function (g) { return g[0][0]; }); };
  Range.prototype.getValues = function () { return this._grid(function (x) { return x ? x.v : ''; }); };
  Range.prototype.getValue = function () { return this.getValues()[0][0]; };
  Range.prototype.getFormulas = function () { return this._grid(function (x) { return (x && x.f) || ''; }); };
  Range.prototype.getFormula = function () { return this.getFormulas()[0][0]; };
  Range.prototype._write = function (i, j, v) {
    var r = this.r + i, c = this.c + j, key = q(this.sh.title) + '!' + colLetters(c) + r;
    var isF = typeof v === 'string' && v.charAt(0) === '=';
    this.sh.cells[r + ',' + c] = isF ? { v: '', f: v } : { v: (v === null || v === undefined) ? '' : v, f: '' };
    if (!(key in this.sh.book.writes)) this.sh.book.writeOrder.push(key);
    this.sh.book.writes[key] = v;
  };
  Range.prototype.setValues = function (vals) {
    for (var i = 0; i < this.nr; i++) for (var j = 0; j < this.nc; j++) this._write(i, j, vals[i] ? vals[i][j] : '');
    return this;
  };
  Range.prototype.setValue = function (v) { for (var i = 0; i < this.nr; i++) for (var j = 0; j < this.nc; j++) this._write(i, j, v); return this; };
  Range.prototype.setFormula = function (f) { return this.setValue(f); };
  Range.prototype.clearContent = function () { return this.setValue(''); };
  Range.prototype._gridRange = function () {
    return { sheetId: this.sh.sid, startRowIndex: this.r - 1, endRowIndex: this.r - 1 + this.nr, startColumnIndex: this.c - 1, endColumnIndex: this.c - 1 + this.nc };
  };
  Range.prototype.setNumberFormat = function (pattern) {
    this.sh.book.ops.push({ repeatCell: { range: this._gridRange(),
      cell: { userEnteredFormat: { numberFormat: { type: 'NUMBER', pattern: pattern } } }, fields: 'userEnteredFormat.numberFormat' } });
    return this;
  };
  Range.prototype.clearDataValidations = function () {
    this.sh.book.ops.push({ setDataValidation: { range: this._gridRange() } });
    return this;
  };
  Range.prototype.setDataValidation = function (rule) {
    this.sh.book.ops.push({ setDataValidation: { range: this._gridRange(), rule: rule } });
    return this;
  };
  // Cosmetic calls the connector may make — not needed on this path.
  ['setFontColor', 'setFontWeight', 'setFontSize', 'setBackground', 'setNote'].forEach(function (m) { Range.prototype[m] = function () { return this; }; });

  // ── Apps Script globals the connector code expects ────────────────────────
  function validationBuilder() {
    var list = [], strict = true;
    var b = {
      requireValueInList: function (vals) { list = vals.slice(); return b; },
      setAllowInvalid: function (allow) { strict = !allow; return b; },
      build: function () {
        return { condition: { type: 'ONE_OF_LIST', values: list.map(function (v) { return { userEnteredValue: String(v) }; }) },
                 showCustomUi: true, strict: strict };
      }
    };
    return b;
  }
  function formatDate(d, tz, pattern) {
    if (!(d instanceof Date)) return String(d || '');
    if (pattern === 'M/d/yyyy') return (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear();
    return d.toISOString();
  }
  // others: { sheetId: Book } — extra sheets already loaded for this request (e.g. the
  // customer's Tiller sheet, read-only). openById hands those back; anything else is refused.
  function env(book, others) {
    return {
      SpreadsheetApp: {
        flush: function () {},          // writes are sent together by book.flush() afterwards
        flushAsync: function () { return book ? book.flush() : Promise.resolve(); },   // stress tests: send writes now
        newDataValidation: validationBuilder,
        DeveloperMetadataVisibility: { DOCUMENT: 'DOCUMENT', PROJECT: 'PROJECT' },
        openById: function (id) {
          if (others && others[id]) return others[id];
          var e = new Error('tiller_access'); e.code = 'tiller_access'; throw e;
        }
      },
      Utilities: { formatDate: formatDate, sleep: function () {},
                   sleepAsync: function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); } },
      Logger: { log: function () { try { if (root.RBP_DEBUG) console.log.apply(console, arguments); } catch (e) {} } },
      Session: { getScriptTimeZone: function () { return book ? book.tz : 'America/Los_Angeles'; } },
      PropertiesService: { getDocumentProperties: function () { return { getProperty: function () { return 'true'; }, setProperty: function () {} }; } }
    };
  }

  // ── Loading ───────────────────────────────────────────────────────────────
  // opts.tabs: { 'Inputs': null (whole tab) | 'A8:CD200' (just that area) }
  function load(id, opts) {
    var metaFields = 'properties(timeZone,iterativeCalculationSettings),developerMetadata(metadataId,metadataKey,metadataValue),' +
                     'sheets(properties(sheetId,title,hidden,gridProperties(rowCount,columnCount)))';
    return G().getSpreadsheet(id, 'fields=' + encodeURIComponent(metaFields)).then(function (meta) {
      var book = new Book(id, meta);
      var want = opts.tabs || {}, ranges = [], order = [];
      Object.keys(want).forEach(function (title) {
        var sh = book.sheets[title];
        if (!sh) return;                                   // tab not in this sheet: connector code handles a missing tab
        var area = want[title];
        if (area) { var p = parseA1(area); sh.window = { r1: p.r, c1: p.c, r2: p.r + p.nr - 1, c2: p.c + p.nc - 1 }; }
        ranges.push(area ? q(title) + '!' + area : q(title));
        order.push(sh);
      });
      if (!ranges.length) return book;
      var fields = 'sheets(properties(title),data(startRow,startColumn,rowData(values(effectiveValue,formattedValue,userEnteredValue/formulaValue,effectiveFormat/numberFormat/type))))';
      var qs = ranges.map(function (r) { return 'ranges=' + encodeURIComponent(r); }).join('&') + '&fields=' + encodeURIComponent(fields);
      return G().getSpreadsheet(id, qs).then(function (full) {
        (full.sheets || []).forEach(function (s) {
          var sh = book.sheets[s.properties.title];
          if (!sh) return;
          sh.loaded = true;
          (s.data || []).forEach(function (d) {
            var r0 = (d.startRow || 0) + 1, c0 = (d.startColumn || 0) + 1;
            (d.rowData || []).forEach(function (row, i) {
              (row.values || []).forEach(function (cell, j) {
                var v = cellToValue(cell), f = (cell && cell.userEnteredValue && cell.userEnteredValue.formulaValue) || '';
                if (v !== '' || f) sh.cells[(r0 + i) + ',' + (c0 + j)] = { v: v, f: f };
              });
            });
          });
        });
        order.forEach(function (sh) { sh.loaded = true; });
        return book;
      });
    });
  }

  root.RBPShim = { load: load, env: env, parseA1: parseA1, colLetters: colLetters, serialToDate: serialToDate };
})(typeof window !== 'undefined' ? window : this);
