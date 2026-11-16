// Unit tests for the master schedule app. Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, scriptsIn, plain, FakeSheet, fakeSpreadsheet } = require('../../../tests/harness');

// lib.js plus every app file, in one scope, as Apps Script deploys them.
const FILES = ['lib/lib.js', ...scriptsIn('apps/master-schedule/src')];

const KALENDER_HEADERS = ['Datum', 'Börjar', 'Slutar', 'Genre', 'Ligans namn', 'Programtitel', 'Avsnitt',
  'Live/Rerun', 'Avsnittsbeskrivning', 'Programbild', 'Avsnittsbild'];
const KEYS = ['date', 'start', 'end', 'genre', 'league', 'series', 'episode', 'live', 'description',
  'seriesImage', 'episodeImage'];

// A Kalender row from { date, start, ... } (keys as in COLUMNS).
const row = (o) => KEYS.map((k) => o[k] || '');

// An "Elmo images" row: [id, name, modified, size, path, url, hash, upMod, status].
const imageRow = (path, status) => ['id-' + path, path.split('/').pop(), '', '', path, '', '', '', status];
const IMAGE_HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Elmo path', 'Elmo URL',
  'Uploaded hash', 'Uploaded modified', 'Status'];

function app({ kalender = [], images = [], headers = KALENDER_HEADERS, properties } = {}) {
  const sheets = [
    new FakeSheet('Kalender', [headers, ...kalender]),
    new FakeSheet('Elmo images', [IMAGE_HEADERS, ...images]),
  ];
  return load(FILES, {
    spreadsheet: fakeSpreadsheet(sheets),
    properties: Object.assign({ ELMO_FOLDER: 'images' }, properties),
  });
}

const NOW = new Date('2026-11-10T12:00:00Z'); // 13:00 in Stockholm

// --- Loading -----------------------------------------------------------------------

test('lib.js and all app files load together without clashing names', () => {
  const gas = app();
  assert.equal(typeof gas.onOpen, 'function');
  assert.equal(typeof gas.handleEdit, 'function');
  assert.equal(typeof gas.runPostDeployTests, 'function');
  assert.equal(gas.get('SCHEDULE_SHEET'), 'Kalender');
});

// --- readSchedule ---------------------------------------------------------------------

test('readSchedule finds the columns by header', () => {
  const gas = app({ kalender: [row({ date: '2026-11-11', start: '18:00' })] });
  const s = gas.readSchedule();
  assert.equal(s.cols.date, 0);
  assert.equal(s.cols.episodeImage, 10);
  assert.equal(s.rows.length, 1);
});

test('readSchedule accepts the misspelt "Progamtitel" header', () => {
  const headers = KALENDER_HEADERS.map((h) => (h === 'Programtitel' ? 'Progamtitel' : h));
  const gas = app({ headers });
  assert.equal(gas.readSchedule().cols.series, 5);
});

test('readSchedule fails clearly when Datum or Börjar is missing', () => {
  const gas = app({ headers: KALENDER_HEADERS.filter((h) => h !== 'Börjar') });
  assert.throws(() => gas.readSchedule(), /The sheet "Kalender" has no column "Börjar"/);
});

test('readSchedule fails clearly when there is no Kalender sheet', () => {
  const gas = load(FILES, { spreadsheet: fakeSpreadsheet([]) });
  assert.throws(() => gas.readSchedule(), /There is no sheet named "Kalender"/);
});

// --- buildComingUp ---------------------------------------------------------------------

function comingUp(kalender, images = []) {
  const gas = app({ kalender, images });
  return plain(gas.buildComingUp(gas.readSchedule(), gas.uploadedImages(), 'proj', NOW)).shows;
}

test('buildComingUp leaves out shows that have started and rows it cannot read', () => {
  const shows = comingUp([
    row({ date: '2026-11-10', start: '12:59', series: 'Started a minute ago' }),
    row({ date: '2026-11-10', start: '13:00', series: 'Starts now' }),
    row({ date: '2026-11-10', start: '13:01', series: 'Next' }),
    row({ date: 'TBD', start: '18:00', series: 'No date' }),
    row({ date: '2026-11-12', start: '', series: 'No time' }),
  ]);
  assert.deepEqual(shows.map((s) => s.title), ['Next']);
});

test('buildComingUp sorts by start and keeps at most MAX_SHOWS', () => {
  const kalender = [];
  for (let day = 25; day >= 11; day--) kalender.push(row({ date: `2026-11-${day}`, start: '18:00', series: 'Day ' + day }));
  const shows = comingUp(kalender);
  assert.equal(shows.length, 10);
  assert.deepEqual(shows.map((s) => s.title).slice(0, 3), ['Day 11', 'Day 12', 'Day 13']);
});

test('buildComingUp fills in a show', () => {
  const [show] = comingUp([row({
    date: '11.11.2026', start: '18.30', end: '20:15', genre: 'Handboll', league: 'Handbollsligan',
    series: 'Handbollsligan', episode: 'IFK – Redbergslid', live: 'Live', description: 'Omgång 12',
  })], [imageRow('images/Genre/Handboll.png', 'Up to date')]);
  assert.deepEqual(show, {
    id: '11.11.2026 18.30 IFK – Redbergslid',
    start: '2026-11-11T17:30:00.000Z',
    end: '2026-11-11T19:15:00.000Z',
    sport: 'Handboll',
    league: 'Handbollsligan',
    title: 'IFK – Redbergslid',
    description: 'Omgång 12',
    image: '/api/files/projects/proj/images/Genre/Handboll.png',
    live: true,
  });
});

test('buildComingUp: the series title is used without an episode, end is null without an end time', () => {
  const [show] = comingUp([row({ date: '2026-11-11', start: '18:00', series: 'Magasinet', live: 'Rerun' })]);
  assert.equal(show.title, 'Magasinet');
  assert.equal(show.end, null);
  assert.equal(show.live, false);
  assert.equal(show.image, '');
});

test('buildComingUp: a show ending after midnight ends the next day', () => {
  const [show] = comingUp([row({ date: '2026-11-11', start: '23:30', end: '01:00', series: 'Natt' })]);
  assert.equal(show.start, '2026-11-11T22:30:00.000Z');
  assert.equal(show.end, '2026-11-12T00:00:00.000Z');
});

test('buildComingUp: "live" is not case sensitive', () => {
  const [show] = comingUp([row({ date: '2026-11-11', start: '18:00', series: 'x', live: ' LIVE ' })]);
  assert.equal(show.live, true);
});

// --- Images -------------------------------------------------------------------------------

test('uploadedImages counts only rows that are uploaded or up to date', () => {
  const gas = app({
    images: [
      imageRow('images/a.png', 'Uploaded 2026-11-01T10:00:00Z'),
      imageRow('images/b.png', 'Up to date'),
      imageRow('images/c.png', 'Error: upload to S3 failed (403)'),
      imageRow('images/d.png', 'New'),
      imageRow('images/e.png', 'Missing in Drive'),
    ],
  });
  const up = plain(gas.uploadedImages());
  assert.deepEqual(Object.keys(up.byPath).sort(), ['images/a.png', 'images/b.png']);
  assert.equal(up.byName['a.png'], 'images/a.png');
  assert.equal(up.byName.a, 'images/a.png');
});

test('uploadedImages marks a file name used in two folders as ambiguous', () => {
  const gas = app({ images: [imageRow('images/A/logo.png', 'Up to date'), imageRow('images/B/logo.png', 'Up to date')] });
  const up = gas.uploadedImages();
  assert.equal(up.byName['logo.png'], null);
  assert.equal(up.byName.logo, null);
});

test('resolveImage takes a URL, an uploaded path or an unambiguous file name', () => {
  const gas = app({
    images: [
      imageRow('images/Serier/Magasinet.png', 'Up to date'),
      imageRow('images/A/logo.png', 'Up to date'),
      imageRow('images/B/logo.png', 'Up to date'),
    ],
  });
  const up = gas.uploadedImages();
  assert.equal(gas.resolveImage('https://cdn/x.png', up), 'https://cdn/x.png');
  assert.equal(gas.resolveImage('/images/Serier/Magasinet.png', up), 'images/Serier/Magasinet.png');
  assert.equal(gas.resolveImage('magasinet', up), 'images/Serier/Magasinet.png');
  assert.equal(gas.resolveImage('Magasinet.PNG', up), 'images/Serier/Magasinet.png');
  assert.equal(gas.resolveImage('logo.png', up), '');                 // in two folders
  assert.equal(gas.resolveImage('images/Serier/Other.png', up), '');  // not uploaded
  assert.equal(gas.resolveImage('Serier/Magasinet.png', up), '');     // a path must be complete
});

function pick(o, images) {
  const gas = app({ kalender: [row(o)], images });
  const s = gas.readSchedule();
  return plain(gas.pickImage(s, s.rows[0], gas.uploadedImages()));
}

const ALL_IMAGES = [
  imageRow('images/ep.png', 'Up to date'),
  imageRow('images/series.png', 'Up to date'),
  imageRow('images/Genre/Handboll.png', 'Up to date'),
];

test('pickImage prefers the episode image, then the series image, then the genre image', () => {
  assert.deepEqual(pick({ genre: 'Handboll', seriesImage: 'series.png', episodeImage: 'ep.png' }, ALL_IMAGES),
    { image: 'images/ep.png', warnings: [] });
  assert.deepEqual(pick({ genre: 'Handboll', seriesImage: 'series.png' }, ALL_IMAGES),
    { image: 'images/series.png', warnings: [] });
  assert.deepEqual(pick({ genre: 'Handboll' }, ALL_IMAGES),
    { image: 'images/Genre/Handboll.png', warnings: [] });
});

test('pickImage warns about image cells that are not in Elmo and falls back', () => {
  assert.deepEqual(pick({ genre: 'Handboll', seriesImage: 'series.png', episodeImage: 'saknas.png' }, ALL_IMAGES),
    { image: 'images/series.png', warnings: ['Avsnittsbild not in Elmo: saknas.png'] });
});

test('pickImage warns when it falls back to a genre image that is not uploaded', () => {
  assert.deepEqual(pick({ genre: 'Fotboll' }, ALL_IMAGES),
    { image: '', warnings: ['Genre image missing: images/Genre/Fotboll.png'] });
});

test('pickImage does not need the genre image when another image is found', () => {
  assert.deepEqual(pick({ genre: 'Fotboll', episodeImage: 'ep.png' }, ALL_IMAGES), { image: 'images/ep.png', warnings: [] });
});

test('genreImage follows ELMO_FOLDER and replaces "/" in the genre name', () => {
  const gas = app({ properties: { ELMO_FOLDER: '/bilder/' } });
  assert.equal(gas.genreImage('Fotboll/Futsal'), 'bilder/Genre/Fotboll-Futsal.png');
  assert.equal(gas.genreImage(''), '');
});

// --- Körschema tab names -------------------------------------------------------------------

test('getKalenderTabName_ removes characters tabs cannot have and keeps 100 characters', () => {
  const gas = app();
  assert.equal(gas.getKalenderTabName_('IFK: Final [1/2]?'), 'IFK  Final  1 2');
  assert.equal(gas.getKalenderTabName_('  '), 'Körschema');
  assert.equal(gas.getKalenderTabName_('x'.repeat(150)).length, 100);
});

test('getUniqueKalenderTabName_ adds (2), (3)… when the name is taken by another tab', () => {
  const gas = app();
  const own = new FakeSheet('Copy of Mall', [], 1);
  const ss = fakeSpreadsheet([own, new FakeSheet('Final', [], 2), new FakeSheet('Final (2)', [], 3)]);
  assert.equal(gas.getUniqueKalenderTabName_(ss, own, 'Final'), 'Final (3)');
  assert.equal(gas.getUniqueKalenderTabName_(ss, own, 'Semi'), 'Semi');
});

test('getUniqueKalenderTabName_ ignores the tab being renamed', () => {
  const gas = app();
  const own = new FakeSheet('Final', [], 1);
  assert.equal(gas.getUniqueKalenderTabName_(fakeSpreadsheet([own]), own, 'Final'), 'Final');
});

test('getUniqueKalenderTabName_ keeps a long name with its suffix within 100 characters', () => {
  const gas = app();
  const long = 'x'.repeat(100);
  const own = new FakeSheet('own', [], 1);
  const name = gas.getUniqueKalenderTabName_(fakeSpreadsheet([own, new FakeSheet(long, [], 2)]), own, long);
  assert.equal(name.length, 100);
  assert.ok(name.endsWith(' (2)'));
});

// --- Kalender IDs ----------------------------------------------------------------------------

test('newKalenderId_ writes a new ID (and the header) in column S, findKalenderRowById_ finds it', () => {
  const gas = app();
  const sheet = new FakeSheet('Kalender', [KALENDER_HEADERS, row({}), row({})]);
  const id = gas.newKalenderId_(sheet, 3);
  assert.equal(sheet.data[0][18], 'Körschema-ID');
  assert.equal(sheet.data[2][18], id);
  assert.equal(gas.findKalenderRowById_(sheet, id), 3);
  assert.equal(gas.findKalenderRowById_(sheet, 'unknown'), 0);
});

// --- Configuration -----------------------------------------------------------------------------

test('config_ requires the Elmo properties and tidies URL and folder', () => {
  const props = { DRIVE_FOLDER_ID: 'f', ELMO_URL: 'https://elmo/', ELMO_KEY: 'k', ELMO_PROJECT: 'p', ELMO_FOLDER: '/img/' };
  assert.deepEqual(plain(app({ properties: props }).config_()),
    { folderId: 'f', elmoUrl: 'https://elmo', key: 'k', project: 'p', folder: 'img' });
  const missing = Object.assign({}, props, { ELMO_KEY: '' });
  assert.throws(() => app({ properties: missing }).config_(), /ELMO_KEY is not set/);
});

// --- Post-deploy tests ------------------------------------------------------------------------
// runPostDeployTests runs against the live project; here it runs against a fake, correctly set up
// one, so a mistake in the checks themselves shows up before deploying.

const { fakeFolder } = require('../../../tests/harness');

// Kalender as the dropdown, Live and archive code expect it by position (A, B, D, E, J, S).
const LIVE_HEADERS = ['Datum', 'Börjar', 'Slutar', 'Genre', 'Ligans namn', 'Hemma', 'Borta', 'Programtitel',
  'Avsnitt', 'Live/Rerun', 'Körschema', 'Länk', 'Avsnittsbeskrivning', 'Programbild', 'Avsnittsbild',
  'Warning', '', '', 'Körschema-ID'];

function liveProject({ headers = LIVE_HEADERS, ids = [''] } = {}) {
  const at = (name) => headers.indexOf(name);
  const upcoming = Array(headers.length).fill('');
  upcoming[at('Datum')] = '2099-01-01';
  upcoming[at('Börjar')] = '18:00';
  upcoming[at('Genre')] = 'Handboll';
  upcoming[at('Programtitel')] = 'Handbollsligan';
  upcoming[18] = ids[0]; // column S
  const ss = fakeSpreadsheet(['Program', 'Ligor&Lag', 'EPG', 'Playout', 'KalenderArkiv']
    .map((n) => new FakeSheet(n, [['header'], ['row']]))
    .concat([
      new FakeSheet('Kalender', [headers, upcoming]),
      new FakeSheet('Inställningar', [[], [], ['', new Date('2026-10-01T00:00:00Z')]]),
      new FakeSheet('Elmo images', [IMAGE_HEADERS, imageRow('images/Genre/Handboll.png', 'Up to date')]),
    ]));
  const assets = fakeFolder({ name: 'Assets', folders: [{ id: 'IMG', name: 'Bilder' }, { name: 'Video' }] });
  const gas = load(FILES, {
    spreadsheet: ss,
    properties: {
      DRIVE_FOLDER_ID: 'IMG', ELMO_URL: 'https://elmo', ELMO_KEY: 'k', ELMO_PROJECT: 'p',
      ELMO_CACHEAROO_URL: 'https://cachearoo', EPG_MAIL_ENDPOINT: 'epg@example.com',
    },
    folders: { IMG: assets.getFolders().next(), '1T_ekwOZNoPo3LNAGV8xCbYCQpaZ46-5Z': fakeFolder({ name: 'Watched' }) },
    fetch: () => ({ code: 200, body: { list: [{ name: 'images/Genre/Handboll.png', url: 'https://cdn/h' }] } }),
    globals: { Sheets: { Spreadsheets: { get: () => ({ spreadsheetId: 'x' }) } } },
  });
  gas.ensureEditTrigger_('handleEdit', ss);
  gas.fakes.triggers.push({ getHandlerFunction: () => 'checkForNewAssets', getTriggerSourceId: () => null });
  return gas;
}

test('runPostDeployTests passes on a correctly set up project and logs every check', () => {
  const gas = liveProject();
  const summary = gas.runPostDeployTests();
  const failed = gas.fakes.logs.filter((l) => !l.startsWith('PASS'));
  assert.equal(summary, '12 passed, 0 warnings, 0 failed', failed.join('\n'));
  assert.ok(gas.fakes.logs.some((l) => /^PASS +Elmo: file listing +1 files in "p"/.test(l)));
});

test('runPostDeployTests fails, naming the check, when a Kalender column has moved', () => {
  const headers = LIVE_HEADERS.slice();
  [headers[3], headers[4]] = [headers[4], headers[3]]; // Genre and Ligans namn swapped
  const gas = liveProject({ headers });
  assert.throws(() => gas.runPostDeployTests(), /Post-deploy tests failed: Kalender columns used by position/);
  assert.ok(gas.fakes.logs.some((l) => l.includes('column D should be "Genre", is "Ligans namn"')));
});

test('runPostDeployTests accepts "ID" as the header of the körschema ID column', () => {
  const headers = LIVE_HEADERS.slice();
  headers[18] = 'ID';
  const gas = liveProject({ headers, ids: ['3f2b1c9e-8a7d-4e6f-9b0a-1c2d3e4f5a6b'] });
  assert.equal(gas.runPostDeployTests(), '12 passed, 0 warnings, 0 failed');
  assert.ok(gas.fakes.logs.some((l) => l.includes('S "ID" holds 1 körschema IDs')));
});

test('runPostDeployTests fails when column S holds something other than körschema IDs', () => {
  const gas = liveProject({ ids: ['Hemmamatch'] });
  assert.throws(() => gas.runPostDeployTests(), /Kalender columns used by position/);
  assert.ok(gas.fakes.logs.some((l) => l.includes('column S row 2 is "Hemmamatch", not a körschema ID')));
});

test('runPostDeployTests reports every failure in an empty project instead of stopping at the first', () => {
  const gas = load(FILES, { spreadsheet: fakeSpreadsheet([]) });
  assert.throws(() => gas.runPostDeployTests(), /Post-deploy tests failed: Script Properties, Sheets exist/);
  assert.equal(gas.fakes.logs.filter((l) => l.startsWith('FAIL')).length, 11);
});
