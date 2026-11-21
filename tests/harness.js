// Runs Apps Script files in Node for unit tests: loads them into one shared global scope (as Apps
// Script does) with fake Google services. Only what the tests need is faked; a call to anything
// else fails loudly, so a test never quietly depends on a real service.
//
//   const { load } = require('../tests/harness');
//   const gas = load(['lib/lib.js'], { properties: { ELMO_KEY: 'k' } });
//   gas.formatTimecode_(1000);            // functions are on the returned object
//   gas.get('MAX_SHOWS');                 // top-level consts are read with get()

'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');

const ROOT = path.resolve(__dirname, '..');

// Values made inside the scripts belong to another JavaScript realm, so assert.deepStrictEqual
// would see their arrays and objects as different types. Compare plain(x) instead.
function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

// --- Time zones (for Utilities.parseDate / formatDate) ----------------------

function wallClock(tz, ms) {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  }).formatToParts(new Date(ms));
  const p = {};
  parts.forEach((x) => { p[x.type] = x.value; });
  return { y: +p.year, M: +p.month, d: +p.day, H: +p.hour, m: +p.minute, s: +p.second };
}

function offsetMs(tz, ms) {
  const w = wallClock(tz, ms);
  return Date.UTC(w.y, w.M - 1, w.d, w.H, w.m, w.s) - Math.floor(ms / 1000) * 1000;
}

const pad = (n, w = 2) => String(n).padStart(w, '0');

const fakeUtilities = {
  DigestAlgorithm: { SHA_256: 'sha256' },
  computeDigest(algorithm, bytes) {
    const digest = crypto.createHash(algorithm).update(Buffer.from(bytes.map((b) => b & 0xff))).digest();
    return Array.from(digest, (b) => (b > 127 ? b - 256 : b)); // Apps Script bytes are signed
  },
  parseDate(text, tz, format) {
    if (format !== 'yyyy-MM-dd HH:mm') throw new Error('fake Utilities.parseDate: unsupported format ' + format);
    const m = /^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2})$/.exec(text);
    if (!m) throw new Error('Unparseable date: ' + text);
    const wall = Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
    let t = wall - offsetMs(tz, wall);
    t = wall - offsetMs(tz, t); // second pass settles dates next to a DST change
    return new Date(t);
  },
  formatDate(date, tz, format) {
    const w = wallClock(tz, date.getTime());
    return format.replace('yyyy', w.y).replace('MM', pad(w.M)).replace('dd', pad(w.d))
      .replace('HH', pad(w.H)).replace('mm', pad(w.m)).replace('ss', pad(w.s));
  },
  getUuid: () => crypto.randomUUID(),
};

// --- Properties ---------------------------------------------------------------

function fakeProperties(initial) {
  const store = Object.assign({}, initial);
  return {
    store,
    getProperty: (k) => (k in store ? store[k] : null),
    setProperty(k, v) { store[k] = String(v); return this; },
    deleteProperty(k) { delete store[k]; return this; },
    getProperties: () => Object.assign({}, store),
  };
}

// --- UrlFetchApp --------------------------------------------------------------
// respond(url, options) returns { code, body } (body: string or object, sent as JSON).
// Every request is kept in .requests as { url, options }.

function fakeResponse(r) {
  const body = typeof r.body === 'string' ? r.body : JSON.stringify(r.body === undefined ? '' : r.body);
  return {
    getResponseCode: () => r.code,
    getContentText: () => body,
    getBlob: () => fakeBlob(body, 'application/octet-stream'),
  };
}

function fakeUrlFetchApp(respond) {
  const requests = [];
  const one = (url, options) => {
    requests.push({ url, options: options || {} });
    return fakeResponse(respond ? respond(url, options || {}) : { code: 200, body: '' });
  };
  return {
    requests,
    fetch: one,
    fetchAll: (list) => list.map((req) => one(req.url, req)),
  };
}

function fakeBlob(content, contentType, name) {
  const bytes = Array.from(Buffer.from(content));
  return {
    getBytes: () => bytes,
    getContentType: () => contentType,
    getName: () => name,
    setName(n) { name = n; return this; },
  };
}

// --- Sheets -------------------------------------------------------------------
// A sheet holds rows of values (row 1 is the header). Ranges read and write that grid.

class FakeRange {
  constructor(sheet, row, col, rows, cols) {
    Object.assign(this, { sheet, row, col, rows, cols });
  }
  cells(fn) {
    const out = [];
    for (let r = 0; r < this.rows; r++) {
      const line = [];
      for (let c = 0; c < this.cols; c++) line.push(fn(this.row + r - 1, this.col + c - 1));
      out.push(line);
    }
    return out;
  }
  getValues() { return this.cells((r, c) => (this.sheet.data[r] || [])[c] ?? ''); }
  // As Sheets shows them: numbers under a duration format ([h]:mm:ss) as h:mm:ss, the rest as text.
  getDisplayValues() {
    return this.cells((r, c) => {
      const v = (this.sheet.data[r] || [])[c] ?? '';
      if (typeof v === 'number' && this.sheet.formats.get(r + ',' + c) === '[h]:mm:ss') {
        const t = Math.floor(v * 86400 + 1e-6);
        const pad = (n) => String(n).padStart(2, '0');
        return `${Math.floor(t / 3600)}:${pad(Math.floor(t / 60) % 60)}:${pad(t % 60)}`;
      }
      return String(v);
    });
  }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  setValues(values) {
    values.forEach((line, r) => line.forEach((v, c) => {
      const rr = this.row + r - 1;
      while (this.sheet.data.length <= rr) this.sheet.data.push([]);
      this.sheet.data[rr][this.col + c - 1] = v;
    }));
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  // Rich text: the text goes in as the value, a link into sheet.links by "row,col".
  setRichTextValues(values) {
    values.forEach((line, r) => line.forEach((v, c) => {
      const key = (this.row + r - 1) + ',' + (this.col + c - 1);
      if (v.getLinkUrl()) this.sheet.links.set(key, v.getLinkUrl()); else this.sheet.links.delete(key);
    }));
    return this.setValues(values.map((line) => line.map((v) => v.getText())));
  }
  clearContent() {
    this.cells((r, c) => { if (this.sheet.data[r]) this.sheet.data[r][c] = ''; this.sheet.links.delete(r + ',' + c); });
    return this;
  }
  setNumberFormat(format) { this.cells((r, c) => this.sheet.formats.set(r + ',' + c, format)); return this; }
  // Checkboxes are kept as the set of their "row,col" in sheet.checkboxes; inserting sets false.
  insertCheckboxes() {
    this.cells((r, c) => this.sheet.checkboxes.add(r + ',' + c));
    return this.setValues(this.cells(() => false));
  }
  clearDataValidations() { this.cells((r, c) => this.sheet.checkboxes.delete(r + ',' + c)); return this; }
  setFontWeight() { return this; }
  setBackgrounds(b) { this.sheet.backgrounds = b; return this.perCell('cellBackgrounds', b); }
  setFontColors(c) { return this.perCell('fontColors', c); }
  setNotes(n) { return this.perCell('notes', n); }
  getNotes() { return this.cells((r, c) => this.sheet.notes.get(r + ',' + c) || ''); }
  // Stores a grid of per-cell settings in sheet[map] by "row,col" (0-based); null or '' removes.
  perCell(map, grid) {
    grid.forEach((line, r) => line.forEach((v, c) => {
      const key = (this.row + r - 1) + ',' + (this.col + c - 1);
      if (v === null || v === '') this.sheet[map].delete(key); else this.sheet[map].set(key, v);
    }));
    return this;
  }
  getA1Notation() { return 'R' + this.row + 'C' + this.col; }
  getSheet() { return this.sheet; }
}

class FakeSheet {
  constructor(name, data, id) {
    this.name = name;
    this.data = (data || []).map((r) => r.slice());
    this.id = id === undefined ? Math.floor(Math.random() * 1e9) : id;
    this.checkboxes = new Set();
    this.links = new Map(); // "row,col" (0-based) -> URL, from setRichTextValues
    this.formats = new Map(); // "row,col" (0-based) -> number format
    this.cellBackgrounds = new Map(); // "row,col" (0-based) -> colour
    this.fontColors = new Map();
    this.notes = new Map();
  }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  getLastRow() {
    let n = this.data.length;
    while (n > 0 && this.data[n - 1].every((v) => v === '' || v == null)) n--;
    return n;
  }
  getLastColumn() { return Math.max(0, ...this.data.map((r) => r.length)); }
  getRange(row, col, rows, cols) {
    if (typeof row === 'string') { // A1 notation for one cell, e.g. "B3"
      const m = /^([A-Z]+)(\d+)$/.exec(row);
      if (!m) throw new Error('fake getRange: unsupported A1 notation ' + row);
      col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
      row = +m[2];
    }
    return new FakeRange(this, row, col, rows || 1, cols || 1);
  }
  getDataRange() { return new FakeRange(this, 1, 1, this.getLastRow(), this.getLastColumn()); }
  setFrozenRows() { return this; }
}

function fakeSpreadsheet(sheets, id, name) {
  const ss = {
    sheets,
    getId: () => id || 'spreadsheet-id',
    getName: () => name || 'Master schedule',
    getUrl: () => `https://docs.google.com/spreadsheets/d/${id || 'spreadsheet-id'}/edit`,
    getSheets: () => sheets,
    getSheetByName: (n) => sheets.find((s) => s.getName() === n) || null,
    insertSheet(n) { const s = new FakeSheet(n, []); sheets.push(s); return s; },
    toast() {},
  };
  sheets.forEach((s) => { s.getParent = () => ss; });
  return ss;
}

// --- Drive ----------------------------------------------------------------------

function iterator(items) {
  let i = 0;
  return { hasNext: () => i < items.length, next: () => items[i++] };
}

function fakeFile(o) {
  return {
    getId: () => o.id || o.name,
    getName: () => o.name,
    getMimeType: () => o.mimeType || 'application/octet-stream',
    getLastUpdated: () => new Date(o.updated || o.created || 0),
    getDateCreated: () => new Date(o.created || 0),
    getSize: () => o.size || 0,
    getUrl: () => 'https://drive/' + (o.id || o.name),
  };
}

// fakeFolder({ name, files: [{ name, mimeType, created }], folders: [ ...same shape ] })
function fakeFolder(o, parent) {
  const folder = {
    getId: () => o.id || o.name,
    getName: () => o.name,
    getUrl: () => 'https://drive/folders/' + (o.id || o.name),
    getParents: () => iterator(parent ? [parent] : []),
  };
  const files = (o.files || []).map(fakeFile);
  const folders = (o.folders || []).map((f) => fakeFolder(f, folder));
  folder.getFiles = () => iterator(files);
  folder.getFolders = () => iterator(folders);
  folder.getFoldersByName = (n) => iterator(folders.filter((f) => f.getName() === n));
  return folder;
}

// --- Loading --------------------------------------------------------------------

function unexpected(name) {
  return new Proxy({}, {
    get(_, prop) {
      if (prop === Symbol.toPrimitive || prop === 'then') return undefined;
      throw new Error(`Test called ${name}.${String(prop)}, which is not faked`);
    },
  });
}

// SpreadsheetApp.newRichTextValue(): a builder for text with a link.
function newRichTextValue() {
  let text = '';
  let url = null;
  const b = {
    setText(t) { text = t; return b; },
    setLinkUrl(u) { url = u; return b; },
    build: () => ({ getText: () => text, getLinkUrl: () => url }),
  };
  return b;
}

/**
 * Loads files (paths relative to the repo root) into one fresh global scope and returns it.
 * opts: properties (Script Properties), fetch (UrlFetchApp respond function), spreadsheet,
 * folders (id -> fake folder for DriveApp.getFolderById), globals (anything else, overriding).
 * The fakes are on the result's .fakes.
 */
function load(files, opts = {}) {
  const logs = [];
  const mail = [];
  const triggers = [];
  const fakes = {
    logs, mail, triggers,
    properties: fakeProperties(opts.properties),
    urlFetch: fakeUrlFetchApp(opts.fetch),
  };
  const root = fakeFolder({ name: 'My Drive', id: 'root' });
  const globals = {
    console: { log: (...a) => logs.push(a.join(' ')), error: (...a) => logs.push(a.join(' ')) },
    Logger: { log: (m) => logs.push(String(m)) },
    Session: {
      getScriptTimeZone: () => 'Europe/Stockholm',
      getActiveUser: () => ({ getEmail: () => 'user@example.com' }),
      getEffectiveUser: () => ({ getEmail: () => 'user@example.com' }),
    },
    Utilities: fakeUtilities,
    PropertiesService: { getScriptProperties: () => fakes.properties },
    UrlFetchApp: fakes.urlFetch,
    GmailApp: { sendEmail: (to, subject, body, options) => mail.push({ to, subject, body, options }) },
    SpreadsheetApp: opts.spreadsheet
      ? { getActive: () => opts.spreadsheet, getActiveSpreadsheet: () => opts.spreadsheet, flush() {}, newRichTextValue }
      : unexpected('SpreadsheetApp'),
    DriveApp: {
      getFolderById(id) {
        const f = (opts.folders || {})[id];
        if (!f) throw new Error('No such folder ' + id);
        return f;
      },
      getRootFolder: () => root,
    },
    ScriptApp: {
      getOAuthToken: () => 'token',
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: (t) => triggers.splice(triggers.indexOf(t), 1),
      newTrigger(handler) {
        let source = null;
        const b = {
          forSpreadsheet(ss) { source = ss.getId(); return b; },
          onEdit() { return b; },
          timeBased() { return b; },
          everyMinutes(n) { b.minutes = n; return b; },
          create() {
            const t = { getHandlerFunction: () => handler, getTriggerSourceId: () => source };
            triggers.push(t);
            return t;
          },
        };
        return b;
      },
    },
    HtmlService: unexpected('HtmlService'),
    LockService: unexpected('LockService'),
    Sheets: unexpected('Sheets'),
    MimeType: { GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet' },
  };
  Object.assign(globals, opts.globals);
  const context = vm.createContext(globals);
  for (const file of files) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, file), 'utf8'), context, { filename: file });
  }
  context.get = (expression) => vm.runInContext(expression, context);
  context.fakes = fakes;
  return context;
}

/** All .js files in a folder, relative to the repo root, sorted. */
function scriptsIn(dir) {
  return fs.readdirSync(path.join(ROOT, dir)).filter((f) => f.endsWith('.js')).sort()
    .map((f) => path.join(dir, f));
}

module.exports = {
  load, scriptsIn, plain, fakeBlob, fakeFolder, fakeFile, FakeSheet, fakeSpreadsheet,
};
