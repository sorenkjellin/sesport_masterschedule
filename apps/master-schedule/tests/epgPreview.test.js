// Unit tests for Data -> Förhandsgranska EPG (epgPreview.js). Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, scriptsIn, plain, FakeSheet, fakeSpreadsheet } = require('../../../tests/harness');

const FILES = ['lib/lib.js', ...scriptsIn('apps/master-schedule/src')];
const HEADERS = ['Date', 'Start time', 'End time', 'Title', 'League', 'Episode title', 'Live Rerun', 'Desc.',
  'Episode description', 'Series Image', 'Episode image', 'Genre', 'Series Img 2:3', 'Episode Img 2:3',
  'Cast/Credits', 'Cast/credits role'];
const show = (o) => HEADERS.map((h) => o[h] ?? '');
const base = { 'Title': 'SM i boxning 2026', 'League': 'SM', 'Episode title': 'Dag 1', 'Genre': 'Boxning',
  'Desc.': 'Sport', 'Episode description': 'Direkt.', 'Series Image': 'https://cdn/sm', 'Live Rerun': 'Live' };
const gas = load(FILES, { spreadsheet: fakeSpreadsheet([]) });
const preview = (rows, headers = HEADERS) => plain(gas.buildEpgPreview_([headers, ...rows.map(show)]));

test('shows are grouped by day in time order; a missing end time is warned about, and runs until the next show', () => {
  const p = preview([
    { ...base, 'Date': '2026-11-28', 'Start time': '12:45', 'Episode title': 'Dag 2' },
    { ...base, 'Date': '27.11.2026', 'Start time': '17:00', 'Episode image': 'https://cdn/dag1',
      'Episode Img 2:3': 'https://cdn/p', 'Cast/credits role': 'Kommentator: A B' },
    { ...base, 'Date': '2026-11-28', 'Start time': '23:30', 'End time': '00:15', 'Episode title': '' },
    show({}).map(() => ''),
  ]);
  assert.deepEqual(p.days.map((d) => [d.date, d.shows.map((s) => s.start + '-' + s.end)]),
    [['2026-11-27', ['17:00-12:45']], ['2026-11-28', ['12:45-23:30', '23:30-00:15']]]);
  assert.deepEqual(p.days[0].shows[0], { date: '2026-11-27', start: '17:00', end: '12:45', endFromNext: true,
    length: '19 h 45 min', genre: 'Boxning', category: 'Sport', title: 'Dag 1', subtitle: 'SM', series: 'SM i boxning 2026',
    description: 'Direkt.', image: 'https://cdn/dag1', portrait: 'https://cdn/p', live: 'Live', cast: '',
    role: 'Kommentator: A B',
    warnings: ['Sluttid saknas: Simply visar den till nästa program (19 h 45 min)'] });
  const late = p.days[1].shows[1];
  assert.deepEqual([late.title, late.series, late.endFromNext, late.length, late.warnings],
    ['SM i boxning 2026', '', false, '45 min', []]);
  assert.deepEqual(p.days[1].shows[0].warnings, ['Sluttid saknas: Simply visar den till nästa program (10 h 45 min)']);
  assert.deepEqual([p.count, p.warnings, p.missing], [3, 2, []]);
});

test('long and overlapping shows, unreadable times and missing images and descriptions are warned about', () => {
  const p = preview([
    { ...base, 'Date': '2026-11-27', 'Start time': '17:00', 'End time': '12:45' },
    { ...base, 'Date': '2026-11-28', 'Start time': '12:45', 'End time': '14:55', 'Series Image': '',
      'Episode description': '', 'Episode title': 'Dag 2' },
    { ...base, 'Date': '2026-11-28', 'Start time': '14:30', 'End time': 'sen', 'Series Image': 'images/x.png',
      'Episode title': 'Talang' },
    { ...base, 'Date': 'snart', 'Start time': '', 'Episode title': 'Okänd' },
  ]);
  const warnings = p.days.flatMap((d) => d.shows.map((s) => [s.title, s.end, s.length, s.warnings]));
  assert.deepEqual(warnings, [
    ['Dag 1', '12:45', '19 h 45 min', ['Längre än 8 timmar: stämmer sluttiden?']],
    ['Dag 2', '14:55', '2 h 10 min', ['Bild saknas', 'Beskrivning saknas']],
    ['Talang', 'sen', '', ['Sluttiden kan inte läsas', 'Bilden är ingen webbadress: images/x.png',
      'Börjar innan Dag 2 har slutat (14:55)']],
    ['Okänd', '', '', ['Datum eller starttid kan inte läsas']],
  ]);
  // The last show without an end time.
  const last = preview([{ ...base, 'Date': '2026-11-27', 'Start time': '17:00' }]).days[0].shows[0];
  assert.deepEqual([last.end, last.length, last.warnings], ['', '', ['Sluttid saknas']]);
  assert.equal(p.days[2].date, 'snart');
  assert.equal(p.warnings, 4);
});

test('columns are found by header wherever they are, and missing ones are named', () => {
  const headers = HEADERS.slice().reverse().map((h) => h.toUpperCase());
  const p = plain(gas.buildEpgPreview_([headers, show({ ...base, 'Date': '2026-11-27', 'Start time': '17:00',
    'End time': '18:00' }).reverse()]));
  assert.equal(p.days[0].shows[0].title, 'Dag 1');
  assert.deepEqual(preview([], ['Date', 'Start time']).missing, ['End time', 'Title', 'Episode image', 'Series Image']);
  // The example row from the live EPG tab.
  const live = '27.11.2026\t17:00\t\tSM i boxning 2026\tSM\tDag 1\tLive\tSport\tDirektsändning.\thttps://cdn.x/e61\t' +
    'https://cdn.x/bd2\tBoxning\thttps://cdn.x/941\thttps://cdn.x/cef\t\tKommentator: Philip Dersen';
  const one = plain(gas.buildEpgPreview_([HEADERS, live.split('\t')])).days[0].shows[0];
  assert.deepEqual([one.title, one.subtitle, one.series, one.image, one.portrait, one.end, one.warnings],
    ['Dag 1', 'SM', 'SM i boxning 2026', 'https://cdn.x/bd2', 'https://cdn.x/cef', '', ['Sluttid saknas']]);
  // A proper row: no episode image, so the series image.
  const proper = '29.11.2026\t16:30\t18:12\tHandbollsligan Herr\tHandbollsligan Herr\tLugi HF - IK Sävehof\tLive\tSport\t' +
    'Svensk handboll.\thttps://cdn.x/6e3\t\tHandboll\thttps://cdn.x/849';
  const two = plain(gas.buildEpgPreview_([HEADERS, proper.split('\t')])).days[0].shows[0];
  assert.deepEqual([two.title, two.subtitle, two.series, two.image, two.portrait, two.length, two.warnings],
    ['Lugi HF - IK Sävehof', 'Handbollsligan Herr', 'Handbollsligan Herr', 'https://cdn.x/6e3', 'https://cdn.x/849', '1 h 42 min', []]);
});

test('the window gets the data with "<" escaped', () => {
  const sheet = new FakeSheet('EPG', [HEADERS, show({ ...base, 'Date': '2026-11-27', 'Start time': '17:00',
    'End time': '18:00', 'Episode title': '</script><b>' })]);
  const g = load(FILES, { spreadsheet: fakeSpreadsheet([sheet]) });
  const template = {};
  const html = { setWidth: () => html, setHeight: () => html };
  g.HtmlService = { createTemplateFromFile: () => Object.assign(template, { evaluate: () => html }) };
  g.SpreadsheetApp.getUi = () => ({ showModalDialog() {} });
  g.previewEpg();
  assert.ok(!template.data.includes('<'));
  assert.equal(JSON.parse(template.data).days[0].shows[0].title, '</script><b>');
  assert.equal(g.epgLength_(120), '2 h');
  assert.equal(g.epgLength_(0), '0 min');
});

test('the Excel file for Simply has columns A to P, as displayed', () => {
  const sheet = new FakeSheet('EPG', [[...HEADERS, 'Q'], [...show({ ...base, 'Cast/credits role': 'Kommentator: A B' }), 'extra']]);
  const g = load(FILES, { spreadsheet: fakeSpreadsheet([sheet]), properties: { EPG_MAIL_ENDPOINT: 'epg@example.com' } });
  let sent = null;
  g.valuesToXlsx_ = (values) => { sent = values; return 'xlsx'; };
  g.exportEPGAndEmail();
  assert.deepEqual(plain(sent[0]), HEADERS);
  assert.equal(sent[1][15], 'Kommentator: A B');
  assert.deepEqual(plain(g.fakes.mail[0].options.attachments), ['xlsx']);
});
