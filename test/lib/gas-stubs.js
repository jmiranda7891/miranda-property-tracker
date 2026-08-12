// A minimal in-memory fake of the Apps Script services Code.js touches (SpreadsheetApp,
// PropertiesService, Session, Utilities, Logger), enough to run the REAL Code.js in a bare
// vm context and exercise its actual RBAC/CRUD/audit logic. Nothing here reaches a real
// Google service. Pattern ported from the CL Social Media App repo's test/shortlist.js,
// extended with a fake spreadsheet since this app's server does real sheet reads/writes
// (the sibling suite's server functions didn't need that).
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { REPO } = require('./harness');

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    this.sheet = sheet; this.row = row; this.col = col;
    this.numRows = numRows || 1; this.numCols = numCols || 1;
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.row + r;
      const arr = [];
      for (let c = 0; c < this.numCols; c++) {
        const colIdx = this.col - 1 + c;
        if (rowIdx === 1) arr.push(this.sheet.headers[colIdx] != null ? this.sheet.headers[colIdx] : '');
        else {
          const dataRow = this.sheet.rows[rowIdx - 2];
          arr.push(dataRow && dataRow[colIdx] != null ? dataRow[colIdx] : '');
        }
      }
      out.push(arr);
    }
    return out;
  }
  setValues(vals) {
    for (let r = 0; r < vals.length; r++) {
      const rowIdx = this.row + r;
      if (rowIdx === 1) {
        for (let c = 0; c < vals[r].length; c++) this.sheet.headers[this.col - 1 + c] = vals[r][c];
      } else {
        const dataIdx = rowIdx - 2;
        while (this.sheet.rows.length <= dataIdx) this.sheet.rows.push([]);
        for (let c = 0; c < vals[r].length; c++) this.sheet.rows[dataIdx][this.col - 1 + c] = vals[r][c];
      }
    }
  }
  setValue(v) { this.setValues([[v]]); }
}

class FakeSheet {
  constructor(name, headers) { this.name = name; this.headers = (headers || []).slice(); this.rows = []; }
  getName() { return this.name; }
  setName(n) { this.name = n; }
  getLastColumn() { return this.headers.length; }
  getRange(r, c, numRows, numCols) { return new FakeRange(this, r, c, numRows, numCols); }
  getDataRange() { return new FakeRange(this, 1, 1, this.rows.length + 1, Math.max(this.headers.length, 1)); }
  appendRow(arr) { this.rows.push(arr.slice()); }
  deleteRow(r) { this.rows.splice(r - 2, 1); }
  autoResizeColumns() {}
}

class FakeSpreadsheet {
  constructor(name, registry) {
    this.name = name; this.registry = registry;
    this.id = 'ss_' + Math.random().toString(36).slice(2);
    this._sheets = [new FakeSheet('Sheet1', [])];
  }
  getId() { return this.id; }
  getUrl() { return 'https://fake.sheet.test/' + this.id; }
  getSheetByName(n) { return this._sheets.filter((s) => s.name === n)[0] || null; }
  getSheets() { return this._sheets; }
  insertSheet(name) { const s = new FakeSheet(name, []); this._sheets.push(s); return s; }
  deleteSheet(sheet) { this._sheets = this._sheets.filter((s) => s !== sheet); }
}

function makeSpreadsheetApp() {
  const registry = {};
  return {
    __registry: registry,
    create(name) { const ss = new FakeSpreadsheet(name, registry); registry[ss.getId()] = ss; return ss; },
    openById(id) { const ss = registry[id]; if (!ss) throw new Error('Spreadsheet not found: ' + id); return ss; }
  };
}

// Loads the REAL Code.js into a bare vm context with fake Apps Script services.
// `props` seeds Script Properties (e.g. {REGISTRY_ID: '...'}). Returns the sandbox, plus
// `__setUser(email)` to change which user Session.getActiveUser() reports mid-test, and
// `__spreadsheetApp` to inspect created sheets directly.
function loadServer(props) {
  const store = Object.assign({}, props || {});
  const currentEmail = { v: 'admin@example.com' };
  const spreadsheetApp = makeSpreadsheetApp();
  const sandbox = {
    Logger: { log() {} },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (store[k] != null ? store[k] : null),
        setProperty: (k, v) => { store[k] = v; }
      })
    },
    Session: { getActiveUser: () => ({ getEmail: () => currentEmail.v }) },
    Utilities: {
      getUuid: () => 'uuid-' + Math.random().toString(36).slice(2, 10),
      sleep: () => {}
    },
    SpreadsheetApp: spreadsheetApp,
    HtmlService: {
      createTemplateFromFile: () => ({ evaluate: () => ({ setTitle: () => ({ addMetaTag: () => ({}) }) }) }),
      createHtmlOutputFromFile: () => ({ getContent: () => '' })
    }
  };
  vm.createContext(sandbox);
  new vm.Script(fs.readFileSync(path.join(REPO, 'Code.js'), 'utf8'), { filename: 'Code.js' }).runInContext(sandbox);
  sandbox.__setUser = (email) => { currentEmail.v = email; };
  sandbox.__spreadsheetApp = spreadsheetApp;
  return sandbox;
}

module.exports = { loadServer, FakeSheet, FakeSpreadsheet };
