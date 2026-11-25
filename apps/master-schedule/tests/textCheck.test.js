// Unit tests for Data -> Kontrollera texter (textCheck.js). Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, scriptsIn, plain, FakeSheet, fakeSpreadsheet } = require('../../../tests/harness');

const FILES = ['lib/lib.js', ...scriptsIn('apps/master-schedule/src')];
const YELLOW = '#fff8c5';

// rows: [text, date]. answer(list sent) -> [{ id, text }], or { code, body } for the raw response.
function schedule(rows, answer = () => []) {
  const line = (text, date, i) => {
    const r = Array(19).fill('');
    [r[0], r[1], r[3], r[4], r[12]] = [date, '1' + i + ':00', 'Handboll', 'Ligan', text];
    return r;
  };
  const header = Array(19).fill('');
  [header[0], header[1], header[3], header[4], header[12]] =
    ['Datum', 'Börjar', 'Genre', 'Ligans namn', 'Avsnittsbeskrivning'];
  const sheet = new FakeSheet('Kalender', [header, ...rows.map(([t, d], i) => line(t, d || '2099-01-01', i))]);
  const sent = [];
  const gas = load(FILES, {
    spreadsheet: fakeSpreadsheet([sheet]),
    properties: { GEMINI_API_KEY: 'secret-key' },
    fetch: (url, options) => {
      const list = JSON.parse(JSON.parse(options.payload).contents[0].parts[0].text);
      sent.push(list);
      const a = answer(list, sent.length);
      if (!Array.isArray(a)) return a;
      return { code: 200, body: { candidates: [{ finishReason: 'STOP', content: { parts: [
        { thought: true, text: 'hm' }, { text: JSON.stringify(a) }] } }] } };
    },
    globals: { Sheets: { Spreadsheets: { get: (id, o) => {
      assert.match(o.ranges[0], /^'Kalender'!M\d+$/);
      const color = sheet.cellBackgrounds.get((+/\d+$/.exec(o.ranges[0])[0] - 1) + ',12');
      return { sheets: [{ data: [{ rowData: [{ values: [color ? { userEnteredFormat: { backgroundColor: {} } } : {}] }] }] }] };
    } } } },
  });
  return { gas, sheet, sent, text: (row) => sheet.data[row - 1][12], color: (row) => sheet.cellBackgrounds.get((row - 1) + ',12'),
    state: () => JSON.parse(gas.fakes.documentProperties.store.TEXT_CHECK_STATE || 'null') };
}
const fix = (list) => list.filter((t) => t.text.includes('matchn')).map((t) => ({ id: t.id, text: t.text.replace('matchn', 'matchen') }));

test('only texts typed on rows from today on are sent, identical ones once, with the key in a header', () => {
  const today = new Intl.DateTimeFormat('sv-SE', { timeZone: 'Europe/Stockholm' }).format(new Date());
  const p = schedule([['Andra matchn'], ['Andra matchn'], ['Bra text', today], ['Gammal matchn', '2020-01-01'], [''],
    ['Formel matchn'], ['Utan datum', 'snart']], fix);
  p.sheet.formulas.set('6,12', '=A1');
  const r = plain(p.gas.textCheckStart('w1'));
  assert.deepEqual(p.sent, [[
    { id: 1, context: 'Handboll · Ligan', text: 'Andra matchn' },
    { id: 2, context: 'Handboll · Ligan', text: 'Bra text' }]]);
  assert.deepEqual(r.items.map((i) => [i.row, i.suggestion]), [[2, 'Andra matchen'], [3, 'Andra matchen']]);
  assert.deepEqual([r.checked, r.failed], [3, 0]);
  const request = p.gas.fakes.urlFetch.requests[0];
  assert.ok(!request.url.includes('secret-key'));
  assert.equal(request.options.headers['x-goog-api-key'], 'secret-key');
  assert.match(request.url, /gemini-3\.5-flash-lite:generateContent$/);
});

test('the shown cell is yellow and selected, and gets its own colour back', () => {
  const p = schedule([['Ena matchn'], ['Andra matchn']], fix);
  p.sheet.cellBackgrounds.set('1,12', '#ff0000');
  const { items } = p.gas.textCheckStart('w');
  assert.equal(p.gas.textCheckStep('w', { item: null, action: null, next: items[0] }).shown, 2);
  assert.equal(p.color(2), YELLOW);
  assert.equal(p.sheet.active.getRow(), 2);
  assert.equal(p.gas.textCheckStep('w', { item: items[0], action: 'update', text: 'Ena matchen!', next: items[1] }).shown, 3);
  assert.deepEqual([p.text(2), p.color(2), p.color(3)], ['Ena matchen!', '#ff0000', YELLOW]);
  p.gas.textCheckStep('w', { item: items[1], action: 'keep', next: null });
  assert.deepEqual([p.text(3), p.color(3)], ['Andra matchn', undefined]);
  p.gas.textCheckEnd('w');
  assert.equal(p.state(), null);
});

test('kept, updated and correct texts are not sent again until they change', () => {
  const p = schedule([['Ena matchn'], ['Andra matchn'], ['Bra text']], fix);
  const { items } = p.gas.textCheckStart('w');
  p.gas.textCheckStep('w', { item: items[0], action: 'update', text: 'Ena matchen', next: null });
  p.gas.textCheckStep('w', { item: items[1], action: 'keep', next: null });
  assert.equal(p.gas.textCheckStart('w').items.length, 0);
  assert.equal(p.sent.length, 1);
  p.sheet.data[3][12] = 'Ny text';
  p.gas.textCheckStart('w');
  assert.deepEqual(p.sent[1].map((t) => t.text), ['Ny text']);
  // Another model: what the first found correct is asked again, the user's decisions stand.
  p.gas.fakes.properties.setProperty('GEMINI_MODEL', 'gemini-x');
  p.gas.textCheckStart('w');
  assert.deepEqual(p.sent[2].map((t) => t.text), ['Ny text']);
});

test('a text changed meanwhile is left alone, and a moved row is found', () => {
  const p = schedule([['Ena matchn'], ['Andra matchn']], fix);
  const { items } = p.gas.textCheckStart('w');
  p.sheet.data[1][12] = 'Nyskriven';
  p.sheet.data.splice(1, 0, Array(19).fill(''));
  assert.equal(p.gas.textCheckStep('w', { item: items[0], action: 'update', text: 'x', next: items[1] }).changed, false);
  assert.deepEqual([p.text(3), p.color(4)], ['Nyskriven', YELLOW]);
  p.gas.textCheckStep('w', { item: items[1], action: 'update', text: '=Andra', next: null });
  assert.equal(p.text(4), "'=Andra");
  assert.throws(() => p.gas.textCheckStep('w', { item: items[1], action: 'update', text: ' ', next: null }), /tom/);
});

test('the timer puts the colour back when the window is gone, also after a sort', () => {
  const p = schedule([['Ena matchn']], fix);
  const { items } = p.gas.textCheckStart('w');
  p.gas.textCheckStep('w', { item: null, action: null, next: items[0] });
  p.gas.removeStaleTextCheck_();
  assert.equal(p.color(2), YELLOW, 'still open');
  p.sheet.cellBackgrounds.delete('1,12');
  p.sheet.cellBackgrounds.set('4,12', YELLOW);
  p.sheet.data.push([], [], Array(19).fill('x'));
  const store = p.gas.fakes.documentProperties.store;
  store.TEXT_CHECK_STATE = JSON.stringify({ ...p.state(), seen: Date.now() - 200000 });
  p.gas.removeStaleTextCheck_();
  assert.deepEqual([p.color(5), p.state()], [undefined, null]);
  // The window wakes up again: it goes on.
  assert.equal(p.gas.textCheckAlive('w'), true);
  assert.equal(p.gas.textCheckStep('w', { item: null, action: null, next: items[0] }).shown, 2);
});

test('a second window is told, and can take over', () => {
  const p = schedule([['Ena matchn']], fix);
  const { items } = p.gas.textCheckStart('a');
  p.gas.textCheckStep('a', { item: null, action: null, next: items[0] });
  assert.deepEqual(plain(p.gas.textCheckStart('b')), { busy: true });
  p.gas.textCheckEnd('b');
  assert.equal(p.color(2), YELLOW);
  assert.equal(p.gas.textCheckStart('b', true).items.length, 1);
  assert.equal(p.color(2), undefined);
  assert.equal(p.gas.textCheckAlive('a'), false);
  assert.throws(() => p.gas.textCheckStep('a', { item: items[0], action: 'update', text: 'x', next: null }), /tagits över/);
  assert.equal(p.text(2), 'Ena matchn');
});

test('rewrites, unchanged texts, unknown ids and empty answers are no suggestions', () => {
  const texts = ['Matchen spelas i Malmö klockan sju', 'Bra text', 'Tom text'];
  const p = schedule(texts.map((t) => [t]), () => [
    { id: 1, text: 'Häng med när drabbningen avgörs i en fullsatt arena ikväll!' },
    { id: 2, text: 'Bra text ' }, { id: 3, text: '' }, { id: 9, text: 'Okänd' }]);
  assert.equal(p.gas.textCheckStart('w').items.length, 0);
});

test('a batch that fails is tried once more, and its texts are reported and not remembered', () => {
  const rows = Array.from({ length: 30 }, (x, i) => ['Text nummer ' + i]);
  let p = schedule(rows, (list, n) => (n === 2 ? { code: 503, body: 'busy' } : []));
  assert.equal(p.gas.textCheckStart('w').failed, 0);
  assert.deepEqual(p.sent.map((l) => l.length), [25, 5, 5]);

  p = schedule(rows, (list) => (list.length === 5 ? { code: 429, body: 'quota' } : []));
  const r = p.gas.textCheckStart('w');
  assert.deepEqual([r.checked, r.failed, r.error], [25, 5, 'Gemini failed (429): quota']);
  p.sent.length = 0;
  p.gas.textCheckStart('w');
  assert.deepEqual(p.sent.map((l) => l.length), [5, 5]);

  p = schedule([['Text']], () => ({ code: 200, body: { candidates: [{ finishReason: 'MAX_TOKENS' }] } }));
  assert.match(p.gas.textCheckStart('w').error, /MAX_TOKENS/);
});

test('the differences are shown word by word', () => {
  const { gas } = schedule([]);
  assert.deepEqual(plain(gas.wordDiff_('Andra matchn i serien', 'Andra matchen i serien')), [
    { text: 'Andra ', type: 'same' }, { text: 'matchn', type: 'removed' }, { text: 'matchen', type: 'added' },
    { text: ' i serien', type: 'same' }]);
  assert.equal(gas.editDistance_('matchn', 'matchen'), 1);
});

test('the menu item shows the Kalender tab and opens the window', () => {
  const p = schedule([]);
  const opened = [];
  const html = { setWidth: () => html, setHeight: () => html };
  p.gas.HtmlService = { createTemplateFromFile: () => ({ evaluate: () => html }) };
  p.gas.SpreadsheetApp.getUi = () => ({ showModelessDialog: (h, title) => opened.push(title) });
  p.gas.kontrolleraTexter();
  assert.equal(p.sheet.activated, true);
  assert.deepEqual(opened, ['Kontrollera texter']);
});
