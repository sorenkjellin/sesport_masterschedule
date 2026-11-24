// Unit tests for Program's Längd (program.js). Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, scriptsIn, FakeSheet, fakeSpreadsheet } = require('../../../tests/harness');

const FILES = ['lib/lib.js', ...scriptsIn('apps/master-schedule/src')];
const H = 1 / 24;

// rows: [length, file ID or '']. files: { id: [E2, F2] or an Error }.
function program(rows, files, { header = 'Längd', locale = 'sv_SE' } = {}) {
  const sheet = new FakeSheet('Program', [['', '', '', '', '', header, 'Körschema'],
    ...rows.map(([length, id]) => ['', '', '', '', '', length, id])]);
  const ss = fakeSpreadsheet([sheet]);
  const state = { reads: [], lists: [], changed: [], cache: {}, toasts: [], busy: false, beforeRead: null };
  ss.getSpreadsheetLocale = () => locale;
  ss.toast = (m) => state.toasts.push(m);
  const gas = load(FILES, {
    spreadsheet: ss,
    fetch: (url) => {
      state.lists.push(decodeURIComponent(url));
      return { code: 200, body: { files: state.changed.map((id) => ({ id })) } };
    },
    globals: {
      CacheService: { getScriptCache: () => ({
        get: (k) => state.cache[k] || null,
        put: (k, v) => { state.cache[k] = v; },
      }) },
      LockService: { getScriptLock: () => ({ tryLock: () => !state.busy, releaseLock() {} }) },
      Sheets: { Spreadsheets: {
        get: () => ({ sheets: [{ data: [{ rowData: sheet.data.map((r, i) => ({ values: [
          {
            formattedValue: String(r[5]),
            effectiveValue: typeof r[5] === 'number' ? { numberValue: r[5] } : { stringValue: r[5] },
            note: sheet.notes.get(i + ',5') || undefined,
            effectiveFormat: sheet.formats.get(i + ',5')
              ? { numberFormat: { type: 'TIME', pattern: sheet.formats.get(i + ',5').replace(/"/g, '') } } : {},
          },
          r[6] ? { chipRuns: [{}, { chip: { richLinkProperties: {
            uri: 'https://docs.google.com/spreadsheets/d/' + r[6] + '/edit' } } }] } : {},
        ] })) }] }] }),
        Values: { get: (id, range) => {
          assert.equal(range, 'E2:F2');
          state.reads.push(id);
          if (state.beforeRead) state.beforeRead();
          if (files[id] instanceof Error) throw files[id];
          return { values: [files[id]] };
        } },
      } },
    },
  });
  const cell = (row) => ({
    value: sheet.data[row - 1][5],
    note: sheet.notes.get((row - 1) + ',5') || '',
    color: sheet.fontColors.get((row - 1) + ',5'),
    format: sheet.formats.get((row - 1) + ',5'),
  });
  const edit = (row, col, value, rows = 1) => {
    if (value !== undefined) sheet.data[row - 1][col - 1] = value;
    gas.handleEdit({ range: sheet.getRange(row, col, rows, 1) });
  };
  return { gas, sheet, state, files, cell, edit };
}
const near = (a, b) => assert.ok(Math.abs(a - b) < 1e-6, a + ' should be ' + b);

test('length: F2 - E2, also past midnight and with an empty start', () => {
  const { gas } = program([], {});
  near(gas.korSchemaLength_(18 * H, 20.5 * H), 2.5 * H);
  near(gas.korSchemaLength_(23 * H, 1.5 * H), 2.5 * H);
  near(gas.korSchemaLength_(23 * H, 25.5 * H), 2.5 * H);
  near(gas.korSchemaLength_('', 2 * H), 2 * H);
  assert.equal(gas.korSchemaLength_(0.1 + 0.2, 0.3), 0, 'rounding is not 24 hours');
  assert.throws(() => gas.korSchemaLength_(0, ''), /F2 är tom/);
  assert.throws(() => gas.korSchemaLength_(0, undefined), /F2 är tom/);
  assert.throws(() => gas.korSchemaLength_(0, '#REF!'), /F2 är inte en tid/);
  assert.throws(() => gas.korSchemaLength_('kl 18', 1), /E2 är inte en tid/);
});

test('a linked row gets the length, grey and a note; an unlinked row is left alone', () => {
  const p = program([[0.5, 'AAA'], ['90 min', '']], { AAA: [18 * H, 20 * H] });
  p.gas.checkProgramLengths();
  near(p.cell(2).value, 2 * H);
  assert.equal(p.cell(2).note, 'Längden styrs av grundkörschemat');
  assert.equal(p.cell(2).color, '#666666');
  assert.equal(p.cell(2).format, '[h]"."mm');
  assert.deepEqual(p.cell(3), { value: '90 min', note: '', color: undefined, format: '[h]"."mm' });
});

test('only changed files are read again; all of them once an hour', () => {
  const p = program([['', 'AAA'], ['', 'BBB']], { AAA: [0, H], BBB: [0, 2 * H] });
  p.gas.checkProgramLengths();
  assert.deepEqual(p.state.reads, ['AAA', 'BBB']);
  p.gas.checkProgramLengths();
  assert.equal(p.state.reads.length, 2);
  assert.match(p.state.lists[0], /modifiedTime > '\d{4}-/);
  assert.match(p.state.lists[0], /corpora=allDrives/);
  p.files.BBB = [0, 3 * H];
  p.state.changed = ['BBB', 'other'];
  p.gas.checkProgramLengths();
  assert.deepEqual(p.state.reads.slice(2), ['BBB']);
  near(p.cell(3).value, 3 * H);
  p.state.changed = [];
  const cache = JSON.parse(p.state.cache.PROGRAM_LENGTHS);
  cache.fullAt -= 3600001;
  p.state.cache.PROGRAM_LENGTHS = JSON.stringify(cache);
  p.gas.checkProgramLengths();
  assert.deepEqual(p.state.reads.slice(3), ['AAA', 'BBB']);
});

test('typing over a linked length puts it back with a message; an unlinked one stays', () => {
  const p = program([['', 'AAA'], ['', '']], { AAA: [0, H] });
  p.gas.checkProgramLengths();
  p.edit(2, 6, 0.9);
  near(p.cell(2).value, H);
  assert.equal(p.state.toasts.length, 1);
  p.edit(3, 6, 0.9);
  assert.equal(p.cell(3).value, 0.9);
  assert.equal(p.state.toasts.length, 1);
  assert.equal(p.state.reads.length, 1, 'no file is read again');
  assert.equal(p.state.lists.length, 0);
  // Several rows at once.
  p.sheet.data[1][5] = ''; p.sheet.data[2][5] = '';
  p.edit(2, 6, undefined, 2);
  near(p.cell(2).value, H);
  assert.equal(p.cell(3).value, '');
});

test('an edit reads only its own rows, and edits in other columns do nothing', () => {
  const p = program([['', 'AAA'], ['', '']], { AAA: [0, H], BBB: [0, 2 * H] });
  p.edit(3, 7, 'BBB');
  near(p.cell(3).value, 2 * H);
  assert.deepEqual(p.state.reads, ['BBB']);
  assert.equal(p.state.toasts.length, 0);
  p.edit(2, 5, 'x');
  p.edit(2, 8, 'x');
  assert.deepEqual(p.state.reads, ['BBB']);
  // The first check still reads everything, not just what is unread.
  p.gas.checkProgramLengths();
  assert.deepEqual(p.state.reads, ['BBB', 'AAA', 'BBB']);
  // An edit keeps the other rows' lengths in the cache.
  p.edit(3, 6, 0.9);
  p.edit(2, 6, 0.9);
  assert.equal(p.state.reads.length, 3);
});

test('removing the link keeps the length, clears grey and note, and frees the cell', () => {
  const p = program([['', 'AAA']], { AAA: [0, H] });
  p.gas.checkProgramLengths();
  p.edit(2, 7, '');
  near(p.cell(2).value, H);
  assert.equal(p.cell(2).note, '');
  assert.ok(!p.cell(2).color);
  p.edit(2, 6, 0.9);
  assert.equal(p.cell(2).value, 0.9);
});

test("a file that can't be read keeps the last length, says why, and recovers", () => {
  const p = program([['', 'AAA'], [0.7, 'BBB']], { AAA: [0, H], BBB: new Error('No access') });
  p.gas.checkProgramLengths();
  assert.equal(p.cell(3).value, 0.7);
  assert.match(p.cell(3).note, /No access.*står kvar/);
  p.files.AAA = [0, ''];
  p.state.changed = ['AAA'];
  p.files.BBB = [0, 2 * H];
  p.gas.checkProgramLengths();
  near(p.cell(2).value, H);
  assert.match(p.cell(2).note, /F2 är tom/);
  p.edit(2, 6, 0.9);
  near(p.cell(2).value, H);
  near(p.cell(3).value, 2 * H);
  assert.equal(p.cell(3).note, 'Längden styrs av grundkörschemat');
});

test("someone's own note is kept", () => {
  const p = program([['', 'AAA']], { AAA: [0, H] });
  p.sheet.notes.set('1,5', 'Ring Kim');
  p.gas.checkProgramLengths();
  assert.equal(p.cell(2).note, 'Ring Kim\nLängden styrs av grundkörschemat');
  p.edit(2, 7, '');
  assert.equal(p.cell(2).note, 'Ring Kim');
});

test('rows moved while files are read get the right lengths', () => {
  const p = program([['', 'AAA'], ['', 'BBB']], { AAA: [0, H], BBB: [0, 2 * H] });
  let moved = false;
  p.state.beforeRead = () => {
    if (!moved) [p.sheet.data[1], p.sheet.data[2]] = [p.sheet.data[2], p.sheet.data[1]];
    moved = true;
  };
  p.gas.checkProgramLengths();
  near(p.cell(2).value, 2 * H);
  near(p.cell(3).value, H);
});

test('nothing is written when up to date, when busy, or when column F is not Längd', () => {
  const p = program([['', 'AAA']], { AAA: [0, H] });
  p.state.busy = true;
  p.gas.checkProgramLengths();
  assert.equal(p.cell(2).value, '');
  p.state.busy = false;
  p.gas.checkProgramLengths();
  assert.equal(p.gas.updateProgramLengths_({}), 0);
  const q = program([[0.5, 'AAA']], { AAA: [0, H] }, { header: 'Annat' });
  assert.throws(() => q.gas.checkProgramLengths(), /should be "Längd", is "Annat"/);
  assert.equal(q.cell(2).value, 0.5);
});

test('a plain link works like a chip; a folder or a person is not a körschema', () => {
  const { gas } = program([], {});
  assert.equal(gas.linkedFileId_({ hyperlink: 'https://docs.google.com/spreadsheets/d/AAA/edit' }), 'AAA');
  assert.equal(gas.linkedFileId_({ hyperlink: 'https://drive.google.com/drive/folders/AAA' }), null);
  assert.equal(gas.linkedFileId_({ chipRuns: [{ chip: { personProperties: {} } }] }), null);
  assert.equal(gas.linkedFileId_({}), null);
});

test('the whole column is formatted h.mm, by the locale, once, but not plain numbers typed by hand', () => {
  const p = program([['', ''], [0.5, ''], [90, ''], ['', 'AAA']], { AAA: [0, H] });
  p.sheet.formats.set('2,5', 'hh:mm');
  let set = 0;
  const getRange = p.sheet.getRange.bind(p.sheet);
  p.sheet.getRange = (...a) => {
    const range = getRange(...a);
    const original = range.setNumberFormat.bind(range);
    range.setNumberFormat = (f) => { set++; return original(f); };
    return range;
  };
  p.gas.checkProgramLengths();
  assert.equal(p.cell(2).format, '[h]"."mm');
  assert.equal(p.cell(3).format, '[h]"."mm');
  assert.equal(p.cell(4).format, undefined);
  assert.equal(p.cell(5).format, '[h]"."mm');
  assert.equal(set, 3);
  p.gas.checkProgramLengths();
  assert.equal(set, 3, 'not again');
  const q = program([['', 'AAA']], { AAA: [0, H] }, { locale: 'en_GB' });
  q.gas.checkProgramLengths();
  assert.equal(q.cell(2).format, '[h]":"mm');
});

test("the note's first wording is replaced", () => {
  const p = program([['', 'AAA']], { AAA: [0, H] });
  p.sheet.notes.set('1,5', 'Ring Kim\nKörschema: längden hämtas från filen i kolumn G.');
  p.gas.checkProgramLengths();
  assert.equal(p.cell(2).note, 'Ring Kim\nLängden styrs av grundkörschemat');
});
