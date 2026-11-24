// Unit tests for Kalender's Slutar (kalenderEndTimes.js, linkedFiles.js). Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, scriptsIn, FakeSheet, fakeSpreadsheet } = require('../../../tests/harness');

const FILES = ['lib/lib.js', ...scriptsIn('apps/master-schedule/src')];
const H = 1 / 24;

// rows: [end, file ID or '', ID, { start, genre, league }]. files: { id: [[tab title, F2, J2], ...] or an Error }.
function kalender(rows, files, { header = 'Slutar', program = [], lengthHeader = 'Längd' } = {}) {
  const line = (end, file, id, o = {}) => {
    const r = Array(19).fill('');
    [r[2], r[11], r[18]] = [end, file, id];
    [r[1], r[3], r[4]] = [o.start ?? '', o.genre || '', o.league || ''];
    return r;
  };
  const sheet = new FakeSheet('Kalender', [line(header, 'Länk', 'Körschema-ID'), ...rows.map((r) => line(...r))]);
  rows.forEach((r, i) => sheet.formats.set((i + 1) + ',2', 'HH:mm'));
  const ss = fakeSpreadsheet([sheet]);
  const state = { reads: [], lists: 0, changed: [], cache: {}, toasts: [], beforeRead: null, beforeWrite: null };
  ss.getSpreadsheetLocale = () => 'sv_SE';
  ss.toast = (m) => state.toasts.push(m);
  const grid = () => ({ sheets: [{ data: [{ rowData: sheet.data.map((r, i) => ({ values: r.slice(1).map((v, c) => {
    if (c === 10) return v ? { hyperlink: 'https://docs.google.com/spreadsheets/d/' + v + '/edit' } : {};
    return {
      formattedValue: String(v),
      effectiveValue: typeof v === 'number' ? { numberValue: v } : { stringValue: v },
      note: sheet.notes.get(i + ',' + (c + 1)) || undefined,
      effectiveFormat: sheet.formats.get(i + ',' + (c + 1)) ? { numberFormat: { type: 'TIME' } } : {},
    };
  }) })) }] }] });
  const gas = load(FILES, {
    spreadsheet: ss,
    fetch: () => { state.lists++; return { code: 200, body: { files: state.changed.map((id) => ({ id })) } }; },
    globals: {
      CacheService: { getScriptCache: () => ({ get: (k) => state.cache[k] || null, put: (k, v) => { state.cache[k] = v; } }) },
      Sheets: { Spreadsheets: {
        get: (id) => {
          if (!files[id]) return grid();
          state.reads.push(id);
          if (state.beforeRead) state.beforeRead();
          if (files[id] instanceof Error) throw files[id];
          return { sheets: files[id].map(([title]) => ({ properties: { title } })) };
        },
        Values: { get: (id, range) => {
          assert.equal(range, 'Program!A1:F');
          return { values: [['Genre', 'Liga', '', '', '', lengthHeader], ...program] };
        }, batchGet: (id, o) => ({ valueRanges: Array.from(o.ranges).map((range, i) => {
          const [title, f2, j2] = files[id][i];
          assert.equal(range, "'" + title.replace(/'/g, "''") + "'!F2:J2");
          return { values: [[f2, '', '', '', j2]] };
        }) }) },
      } },
    },
  });
  gas.updateProgramLengths_ = () => 0;
  gas.removeUnusedKorSchemaTriggersHourly_ = () => {};
  const cell = (row) => ({
    value: sheet.data[row - 1][2],
    note: sheet.notes.get((row - 1) + ',2') || '',
    color: sheet.fontColors.get((row - 1) + ',2'),
    format: sheet.formats.get((row - 1) + ',2'),
  });
  const edit = (row, col, value) => {
    sheet.data[row - 1][col - 1] = value;
    gas.handleEdit({ range: sheet.getRange(row, col, 1, 1) });
  };
  return { gas, sheet, state, files, cell, edit };
}
const near = (a, b) => assert.ok(Math.abs(a - b) < 1e-6, a + ' should be ' + b);
const NOTE = 'Sluttid styrs av körschemats längd';

test("each linked row gets its own tab's F2, grey and a note; an unlinked row is left alone", () => {
  const p = kalender([['', 'AAA', 'id2'], ['', 'AAA', 'id1'], [0.5, '', '']],
    { AAA: [['Mall', 1, ''], ["Lag's match", 20 * H, 'id1'], ['B', 25.5 * H, 'id2']] });
  p.gas.checkProgramLengths();
  near(p.cell(2).value, 25.5 * H);
  near(p.cell(3).value, 20 * H);
  assert.deepEqual(p.cell(2), { value: p.cell(2).value, note: NOTE, color: '#666666', format: 'HH:mm' });
  assert.deepEqual(p.cell(4), { value: 0.5, note: '', color: undefined, format: 'HH:mm' });
  assert.deepEqual(p.state.reads, ['AAA'], 'the file is read once');
});

test('a cell without a time format gets one', () => {
  const p = kalender([['', 'AAA', 'id1']], { AAA: [['A', 20 * H, 'id1']] });
  p.sheet.formats.clear();
  p.gas.checkProgramLengths();
  assert.equal(p.cell(2).format, 'hh"."mm');
});

test('typing over a governed end time puts it back; without a link it stays', () => {
  const p = kalender([['', 'AAA', 'id1'], ['', '', '']], { AAA: [['A', 20 * H, 'id1']] });
  p.gas.checkProgramLengths();
  p.edit(2, 3, 0.1);
  near(p.cell(2).value, 20 * H);
  assert.deepEqual(p.state.toasts, ['Sluttiden styrs av körschemats längd']);
  p.edit(3, 3, 0.1);
  assert.equal(p.cell(3).value, 0.1);
  assert.equal(p.state.toasts.length, 1);
  assert.equal(p.state.reads.length, 1);
});

test('clearing L frees the cell and keeps the end time', () => {
  const p = kalender([['', 'AAA', 'id1']], { AAA: [['A', 20 * H, 'id1']] });
  p.gas.checkProgramLengths();
  p.edit(2, 12, '');
  near(p.cell(2).value, 20 * H);
  assert.equal(p.cell(2).note, '');
  assert.ok(!p.cell(2).color);
  p.edit(2, 3, 0.1);
  assert.equal(p.cell(2).value, 0.1);
});

test('a change in the körschema is followed', () => {
  const p = kalender([['', 'AAA', 'id1']], { AAA: [['A', 20 * H, 'id1']] });
  p.gas.checkProgramLengths();
  p.files.AAA = [['A', 21 * H, 'id1']];
  p.gas.checkProgramLengths();
  near(p.cell(2).value, 20 * H);
  p.state.changed = ['AAA'];
  p.gas.checkProgramLengths();
  near(p.cell(2).value, 21 * H);
});

test('an end time just synced by the trigger is not replaced by an older cached one', () => {
  const p = kalender([['', 'AAA', 'id1']], { AAA: [['A', 20 * H, 'id1']] });
  p.gas.checkProgramLengths();
  // The körschema's edit trigger: F2 is now 21, and Drive does not report the change yet.
  p.files.AAA = [['A', 21 * H, 'id1']];
  const cells = { F2: 21 * H, J1: 20 * H, J2: 'id1' };
  const tab = { getParent: () => ({ getId: () => 'AAA' }),
    getRange: (a) => ({ getValue: () => cells[a], setValue: (v) => { cells[a] = v; } }) };
  p.gas.SpreadsheetApp.openById = () => p.sheet.getParent();
  p.gas.syncKorSchemaValueToKalender_({ range: { getSheet: () => tab } });
  near(p.cell(2).value, 21 * H);
  p.gas.checkProgramLengths();
  near(p.cell(2).value, 21 * H);
  assert.equal(p.state.reads.length, 2);
  // F2 emptied: Kalender keeps its end time.
  cells.F2 = '';
  p.files.AAA = [['A', '', 'id1']];
  p.gas.syncKorSchemaValueToKalender_({ range: { getSheet: () => tab } });
  near(p.cell(2).value, 21 * H);
  p.gas.checkProgramLengths();
  near(p.cell(2).value, 21 * H);
  assert.match(p.cell(2).note, /sluttiden kunde inte hämtas \(F2 är tom\)/);
  p.edit(2, 3, 0.1);
  near(p.cell(2).value, 21 * H);
});

test('problems keep the value and say why', () => {
  const p = kalender([[0.1, 'AAA', 'other'], [0.2, 'AAA', ''], [0.3, 'BBB', 'id1']],
    { AAA: [['A', 20 * H, 'id1']], BBB: new Error('No access') });
  p.gas.checkProgramLengths();
  assert.deepEqual([2, 3, 4].map((r) => p.cell(r).value), [0.1, 0.2, 0.3]);
  assert.match(p.cell(2).note, /ingen flik har radens ID/);
  assert.match(p.cell(3).note, /saknar ID/);
  assert.match(p.cell(4).note, /No access/);
  p.files.BBB = [['B', 20 * H, 'id1']];
  p.gas.checkProgramLengths();
  near(p.cell(4).value, 20 * H);
  assert.equal(p.cell(4).note, NOTE);
});

test('a quota error changes nothing for a file already read', () => {
  const p = kalender([['', 'AAA', 'id1']], { AAA: [['A', 20 * H, 'id1']] });
  p.gas.checkProgramLengths();
  p.files.AAA = new Error('Quota exceeded for quota metric');
  p.state.changed = ['AAA'];
  p.gas.checkProgramLengths();
  assert.equal(p.cell(2).note, NOTE);
});

test('at most 15 files are read per run, and the rest by the next ones', () => {
  const files = {};
  const rows = Array.from({ length: 20 }, (_, i) => {
    files['FILE' + i] = [['A', 20 * H, 'id' + i]];
    return ['', 'FILE' + i, 'id' + i];
  });
  const p = kalender(rows, files);
  p.gas.checkProgramLengths();
  assert.equal(p.state.reads.length, 15);
  p.gas.checkProgramLengths();
  assert.equal(p.state.reads.length, 20);
  near(p.cell(21).value, 20 * H);
  p.gas.checkProgramLengths();
  assert.equal(p.state.reads.length, 20);
});

test('rows moved while files are read, or just before writing, get no wrong end time', () => {
  const p = kalender([['', 'AAA', 'id1'], ['', '', '']], { AAA: [['A', 20 * H, 'id1']] });
  const swap = () => { [p.sheet.data[1], p.sheet.data[2]] = [p.sheet.data[2], p.sheet.data[1]]; };
  p.state.beforeRead = swap;
  p.gas.checkProgramLengths();
  assert.equal(p.cell(2).value, '');
  near(p.cell(3).value, 20 * H);
  // Moved after the rows were read: checked before writing.
  p.sheet.data[2][2] = 0.1;
  const read = p.gas.readKalenderRows_;
  p.gas.readKalenderRows_ = (ss) => { const rows = read(ss); swap(); return rows; };
  p.gas.updateKalenderEndTimes_({ changes: {} });
  assert.equal(p.cell(3).value, '');
});

test('Drive is asked for changes once per run', () => {
  const p = kalender([['', 'AAA', 'id1']], { AAA: [['A', 20 * H, 'id1']] });
  const changes = {};
  p.gas.spreadsheetsChangedSince_(1000000, changes);
  p.gas.spreadsheetsChangedSince_(999000, changes);
  assert.equal(p.state.lists, 1);
  p.gas.spreadsheetsChangedSince_(1, changes);
  assert.equal(p.state.lists, 2, 'not when more is needed');
});

test('one sheet failing does not stop the other, and the run is marked failed', () => {
  const p = kalender([['', 'AAA', 'id1']], { AAA: [['A', 20 * H, 'id1']] });
  p.gas.updateProgramLengths_ = () => { throw new Error('Program column F'); };
  assert.throws(() => p.gas.checkProgramLengths(), /Program column F/);
  near(p.cell(2).value, 20 * H);
  const q = kalender([[0.1, 'AAA', 'id1']], { AAA: [['A', 20 * H, 'id1']] }, { header: 'Annat' });
  assert.throws(() => q.gas.checkProgramLengths(), /should be "Slutar", is "Annat"/);
  assert.equal(q.cell(2).value, 0.1);
});

// --- The copy of Program's Längd ----------------------------------------------------------------

const PROGRAM = [
  ['Handboll', 'SHL', '', '', '', 9 * H],
  ['Ishockey', 'SHL', '', '', '', 90], // a plain number is not a length
  ['Ishockey', 'SHL ', '', '', '', 2.5 * H],
  ['Ishockey', 'SHL', '', '', '', 5 * H],
  ['Fotboll', 'Allsvenskan', '', '', '', 0],
  ['Golf', 'Touren'],
];
const at = (start, genre, league) => ({ start, genre, league });
const copy = (o, extra) => kalender([[0.1, '', '', o]], {}, { program: PROGRAM, ...extra });

test('editing Börjar, Genre or Ligans namn sets Slutar to Börjar + Längd', () => {
  for (const [col, value] of [[2, 18 * H], [4, 'ishockey'], [5, 'SHL']]) {
    const p = copy(at(18 * H, 'Ishockey', 'SHL'));
    p.edit(2, col, value);
    near(p.cell(2).value, 20.5 * H);
    assert.equal(p.cell(2).note, '');
    assert.ok(!p.cell(2).color);
    assert.equal(p.cell(2).format, 'HH:mm');
    // It is a copy: typed over, also after the timer, until one of them changes again.
    p.edit(2, 3, 0.2);
    p.gas.checkProgramLengths();
    assert.equal(p.cell(2).value, 0.2);
    p.edit(2, 2, 19 * H);
    near(p.cell(2).value, 21.5 * H);
  }
});

test('past midnight, and a start typed as text', () => {
  const p = copy(at('23.00', 'Ishockey', 'SHL'));
  p.edit(2, 5, 'SHL');
  near(p.cell(2).value, 1.5 * H);
});

test('Slutar is emptied when the combination is not in Program', () => {
  // At once when Genre no longer fits the league, and filled in again when it does.
  const p = copy(at(18 * H, 'Ishockey', 'SHL'));
  p.edit(2, 4, 'Golf');
  assert.equal(p.cell(2).value, '');
  p.edit(2, 4, 'Ishockey');
  near(p.cell(2).value, 20.5 * H);
});

test('Slutar is left alone without all three values, when Program has no length, and by other columns', () => {
  for (const o of [at('', 'Ishockey', 'SHL'), at(18 * H, 'Ishockey', ''), at(18 * H, '', 'SHL'),
    at('kväll', 'Ishockey', 'SHL'), at(18 * H, 'Fotboll', 'Allsvenskan'), at(18 * H, 'Golf', 'Touren')]) {
    const p = copy(o);
    p.edit(2, 5, o.league);
    assert.equal(p.cell(2).value, 0.1, JSON.stringify(o));
  }
  const p = copy(at(18 * H, 'Ishockey', 'SHL'));
  p.edit(2, 8, 'x');
  assert.equal(p.cell(2).value, 0.1);
});

test('a row with a körschema keeps its end time', () => {
  const p = kalender([['', 'AAA', 'id1', at(18 * H, 'Ishockey', 'SHL')]],
    { AAA: [['A', 22 * H, 'id1']] }, { program: PROGRAM });
  p.gas.checkProgramLengths();
  p.edit(2, 5, 'SHL');
  near(p.cell(2).value, 22 * H);
  p.edit(2, 5, 'KHL');
  near(p.cell(2).value, 22 * H);
});

test('several rows pasted at once; an end time pasted along is kept, a missing one filled in', () => {
  const p = kalender([[0.1, '', '', at(18 * H, 'Ishockey', 'SHL')], ['', '', '', at(10 * H, 'Handboll', 'SHL')]],
    {}, { program: PROGRAM });
  p.gas.handleKalenderEndEdit_ = () => {};
  const paste = (col, width) => p.gas.copyProgramLengths_({ range: p.sheet.getRange(2, col, 2, width) });
  paste(1, 5);
  assert.equal(p.cell(2).value, 0.1);
  near(p.cell(3).value, 19 * H);
  paste(4, 2);
  near(p.cell(2).value, 20.5 * H);
});

test('nothing is copied when Program column F is not Längd', () => {
  const p = copy(at(18 * H, 'Ishockey', 'SHL'), { lengthHeader: 'Annat' });
  assert.throws(() => p.edit(2, 5, 'SHL'), /should be "Längd"/);
  assert.equal(p.cell(2).value, 0.1);
});
