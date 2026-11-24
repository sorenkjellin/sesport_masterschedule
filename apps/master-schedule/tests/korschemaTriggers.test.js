// Unit tests for deleting unused körschema triggers (korschema.js). Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, scriptsIn, FakeSheet, fakeSpreadsheet } = require('../../../tests/harness');

const FILES = ['lib/lib.js', ...scriptsIn('apps/master-schedule/src')];

// Kalender rows: [date, the file ID linked in L or '']. The archive date is 2026-11-20.
function project(rows, triggerFiles, { form = 'hyperlink' } = {}) {
  const kalender = new FakeSheet('Kalender', [Array(19).fill('h'),
    ...rows.map(([date, id]) => [date, ...Array(10).fill(''), id, ...Array(7).fill('')])]);
  kalender.deleteRow = (r) => kalender.data.splice(r - 1, 1);
  const archive = new FakeSheet('KalenderArkiv', [['Datum'], ['']]);
  const range = archive.getRange.bind(archive);
  archive.getRange = (...a) => (a[0] === 'A:A' ? { getValues: () => [['Datum'], ['']] } : range(...a));
  const settings = new FakeSheet('Inställningar', []);
  settings.getRange = () => ({ getValue: () => new Date('2026-11-20T00:00:00Z') });
  const ss = fakeSpreadsheet([kalender, archive, settings]);
  const state = { cache: {}, failSheets: false, alerts: [], trash: [], unreadable: '' };
  const url = (id) => 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
  const gas = load(FILES, {
    spreadsheet: ss,
    globals: {
      CacheService: { getScriptCache: () => ({
        get: (k) => state.cache[k] || null, put: (k, v) => { state.cache[k] = v; },
      }) },
      Sheets: { Spreadsheets: { get: (id, options) => {
        if (state.failSheets) throw new Error('quota');
        assert.deepEqual(Array.from(options.ranges), ['Kalender!L2:L']);
        return { sheets: [{ data: [{ rowData: kalender.data.slice(1).map((r) => (!r[11] ? {} : { values: [
          form === 'hyperlink' ? { hyperlink: url(r[11]) }
            : { userEnteredValue: { formulaValue: '=HYPERLINK("' + url(r[11]) + '"; "Körschema")' } },
        ] })) }] }] };
      } } },
    },
  });
  gas.SpreadsheetApp.openById = (id) => ({ getId: () => id });
  gas.DriveApp.getFileById = (id) => {
    if (state.unreadable === id) throw new Error('No access');
    return { isTrashed: () => state.trash.includes(id) };
  };
  gas.SpreadsheetApp.getUi = () => ({ alert: (m) => state.alerts.push(m) });
  triggerFiles.forEach((id) => gas.ensureEditTrigger_('syncKorSchemaValueToKalender_', { getId: () => id }));
  gas.ensureEditTrigger_('handleEdit', { getId: () => 'master' });
  gas.installProgramLengthTrigger();
  const left = () => gas.fakes.triggers
    .map((t) => t.getTriggerSourceId() || t.getHandlerFunction()).sort();
  return { gas, state, kalender, left };
}
const OLD = new Date('2026-11-10T12:00:00Z');
const NEW = new Date('2026-11-25T12:00:00Z');

test('triggers of files no Kalender row links to are deleted; the others, and other triggers, stay', () => {
  for (const form of ['hyperlink', 'formula']) {
    const p = project([[NEW, 'AAA'], [NEW, ''], [NEW, 'BBB']], ['AAA', 'BBB', 'OLD1', 'OLD2'], { form });
    assert.equal(p.gas.removeUnusedKorSchemaTriggers(), 2);
    assert.deepEqual(p.left(), ['AAA', 'BBB', 'checkProgramLengths', 'master']);
  }
});

test('with no links in L, everything goes only when column S is in place and has no IDs', () => {
  const run = (change) => {
    const p = project([[NEW, '']], ['AAA', 'BBB']);
    p.kalender.data[0][18] = 'Körschema-ID';
    change(p.kalender.data);
    return p.gas.removeUnusedKorSchemaTriggers();
  };
  assert.equal(run(() => {}), 2);
  assert.equal(run((d) => { d[0][18] = 'ID'; }), 2);
  assert.equal(run((d) => { d[0][18] = ''; }), 2, 'no header yet');
  assert.equal(run((d) => { d.length = 1; }), 2, 'no rows');
  assert.equal(run((d) => { d[1][18] = 'some-id'; }), 0, 'IDs but no links');
  assert.equal(run((d) => { d[0][18] = 'Annat'; }), 0, 'another column');
  assert.equal(run((d) => { d[0][18] = ''; d[0][19] = 'Körschema-ID'; }), 0, 'moved right');
  assert.equal(run((d) => { d[0][18] = ''; d[0][17] = 'ID'; }), 0, 'moved left');
});

test('archiving deletes the triggers of the archived rows, but not of a file another row still uses', () => {
  const p = project([[OLD, 'AAA'], [OLD, 'SHARED'], [NEW, 'SHARED']], ['AAA', 'SHARED']);
  p.gas.archiveOldCalendarRows();
  assert.equal(p.kalender.data.length, 2);
  assert.deepEqual(p.left(), ['SHARED', 'checkProgramLengths', 'master']);
});

test('archiving the last körschema deletes the last trigger', () => {
  const p = project([[OLD, 'AAA'], [NEW, '']], ['AAA']);
  p.kalender.data[0][18] = 'Körschema-ID';
  p.gas.archiveOldCalendarRows();
  assert.deepEqual(p.left(), ['checkProgramLengths', 'master']);
});

test('archiving completes when the triggers cannot be checked', () => {
  const p = project([[OLD, 'AAA'], [NEW, 'BBB']], ['AAA']);
  p.state.failSheets = true;
  p.gas.archiveOldCalendarRows();
  assert.match(p.state.alerts[0], /^Klar!/);
  assert.equal(p.left().length, 3);
});

test("the owner's minute check deletes them too, once an hour", () => {
  const p = project([[NEW, 'AAA']], ['AAA', 'OLD1']);
  p.gas.updateProgramLengths_ = () => 0;
  p.gas.checkProgramLengths();
  assert.deepEqual(p.left(), ['AAA', 'checkProgramLengths', 'master']);
  p.gas.ensureEditTrigger_('syncKorSchemaValueToKalender_', { getId: () => 'OLD2' });
  p.gas.checkProgramLengths();
  assert.equal(p.left().length, 4);
  p.state.cache = {};
  p.gas.checkProgramLengths();
  assert.equal(p.left().length, 3);
});

test('a trashed file loses its trigger though Kalender links to it, and gets it back when restored', () => {
  const p = project([[NEW, 'AAA'], [NEW, 'BBB'], [NEW, 'CCC']], ['AAA', 'BBB', 'CCC']);
  p.gas.updateProgramLengths_ = () => 0;
  const hour = () => { p.state.cache = {}; p.gas.checkProgramLengths(); };
  p.state.trash = ['AAA', 'BBB'];
  p.state.unreadable = 'CCC';
  hour();
  assert.deepEqual(p.left(), ['CCC', 'checkProgramLengths', 'master'], 'unreadable: kept');
  hour();
  p.state.trash = ['BBB'];
  p.state.unreadable = 'AAA';
  hour();
  assert.equal(p.left().includes('AAA'), false);
  p.state.unreadable = '';
  hour();
  assert.deepEqual(p.left(), ['AAA', 'CCC', 'checkProgramLengths', 'master']);
  // Restored after Kalender dropped it: forgotten.
  p.kalender.data.splice(2, 1);
  p.state.trash = [];
  hour();
  assert.deepEqual(p.left(), ['AAA', 'CCC', 'checkProgramLengths', 'master']);
  assert.equal(p.gas.trashedKorscheman_().length, 0);
  // Triggers are only added back, never added to a file that had none.
  p.kalender.data[1][11] = 'NEWFILE';
  hour();
  assert.deepEqual(p.left(), ['CCC', 'checkProgramLengths', 'master']);
});
