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
const VIDEO_HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Elmo path', 'Duration',
  'Skicka till DMC', 'DMC-status'];

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

// The DMC job's service account, and the job as the Cloud Run Admin API describes it.
const SA = 'dmc-transfer@sesport.iam.gserviceaccount.com';
const DMC_JOB = {
  name: 'projects/sesport/locations/europe-north1/jobs/dmc-transfer',
  template: { template: { serviceAccount: SA, containers: [{ env: [
    { name: 'FTP_HOST', value: 'ftp-no1.dmcproduction.se' }, { name: 'FTP_USER', value: 'u' },
    { name: 'FTP_PASSWORD', valueSource: { secretKeyRef: { secret: 'dmc-ftp-password', version: 'latest' } } },
    { name: 'SPREADSHEET_ID', value: 'spreadsheet-id' },
  ] }] } },
};

// Answers the Google Cloud requests of the DMC code (reading the job, the caller's permissions on
// it, the password secret's access, starting it), and Elmo's file listing for anything else.
function cloudFetch({ job = DMC_JOB, granted = true, runCode = 200,
  secretPolicy = { bindings: [{ role: 'roles/secretmanager.secretAccessor', members: ['serviceAccount:' + SA] }] } } = {}) {
  return (url) => {
    if (url.endsWith(':testIamPermissions')) return { code: 200, body: granted ? { permissions: ['run.jobs.runWithOverrides'] } : {} };
    if (url.endsWith(':run')) {
      return runCode === 200 ? { code: 200, body: { metadata: { name: 'executions/e1' } } } : { code: runCode, body: 'Permission denied' };
    }
    if (url.startsWith('https://secretmanager.googleapis.com/')) {
      return secretPolicy ? { code: 200, body: secretPolicy } : { code: 403, body: 'Permission denied' };
    }
    if (url.startsWith('https://run.googleapis.com/')) return job ? { code: 200, body: job } : { code: 404, body: 'Job not found' };
    return { code: 200, body: { list: [{ name: 'images/Genre/Handboll.png', url: 'https://cdn/h' }] } };
  };
}

// The Drive advanced service's sharing: sharing maps a Drive ID (or '*', any other) to its
// permissions; reading the sharing of an ID not in it fails, as for an item one may not manage.
function fakeDriveSharing(sharing) {
  return {
    Permissions: {
      list(id) {
        const permissions = id in sharing ? sharing[id] : sharing['*'];
        if (!permissions) throw new Error('The user does not have sufficient permissions for file ' + id);
        return { permissions };
      },
    },
  };
}
const sharedWith = (role, email = SA) => [{ type: 'user', emailAddress: email, role }];
// The spreadsheet ('spreadsheet-id') and the Video folder ('Video') shared as the job needs them.
const SHARED = { 'spreadsheet-id': sharedWith('writer'), Video: sharedWith('reader') };

function liveProject({ headers = LIVE_HEADERS, ids = [''], dmcJob = DMC_JOB, properties = {}, cloud = {},
  sharing = SHARED } = {}) {
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
      new FakeSheet('Video Assets', [VIDEO_HEADERS]),
    ]));
  const assets = fakeFolder({ name: 'Assets', folders: [{ id: 'IMG', name: 'Bilder' }, { name: 'Video' }] });
  const gas = load(FILES, {
    spreadsheet: ss,
    properties: {
      DRIVE_FOLDER_ID: 'IMG', ELMO_URL: 'https://elmo', ELMO_KEY: 'k', ELMO_PROJECT: 'p',
      ELMO_CACHEAROO_URL: 'https://cachearoo', EPG_MAIL_ENDPOINT: 'epg@example.com',
      WATCH_FOLDER_ID: 'WATCHED', WATCH_RECIPIENTS: 'watch@example.com', WATCH_RUN_AS: 'User@example.com',
      DMC_GCP_PROJECT: 'sesport',
      ...properties,
    },
    folders: { IMG: assets.getFolders().next(), WATCHED: fakeFolder({ id: 'WATCHED', name: 'Watched' }) },
    fetch: cloudFetch(Object.assign({ job: dmcJob }, cloud)),
    globals: { Sheets: { Spreadsheets: { get: () => ({ spreadsheetId: 'x' }) } }, Drive: fakeDriveSharing(sharing) },
  });
  gas.ensureEditTrigger_('handleEdit', ss);
  gas.fakes.triggers.push({ getHandlerFunction: () => 'checkForNewAssets', getTriggerSourceId: () => null });
  return gas;
}

test('runPostDeployTests passes on a correctly set up project and logs every check', () => {
  const gas = liveProject();
  const summary = gas.runPostDeployTests();
  const failed = gas.fakes.logs.filter((l) => !l.startsWith('PASS'));
  assert.equal(summary, '14 passed, 0 warnings, 0 failed', failed.join('\n'));
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
  assert.equal(gas.runPostDeployTests(), '14 passed, 0 warnings, 0 failed');
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

// --- Asset watcher ------------------------------------------------------------------------------

test('checkForNewAssets watches the folder from WATCH_FOLDER_ID and mails WATCH_RECIPIENTS', () => {
  const now = Date.now();
  const folder = fakeFolder({ id: 'W', name: 'Images', files: [{ name: 'new.png', created: now - 60000 }] });
  const gas = load(FILES, {
    folders: { W: folder },
    properties: { WATCH_FOLDER_ID: 'W', WATCH_RECIPIENTS: 'a@example.com, b@example.com' },
  });
  gas.checkForNewAssets();
  const [m] = gas.fakes.mail;
  assert.equal(m.to, 'a@example.com, b@example.com');
  assert.equal(m.subject, '[Drive Alert] New Asset Uploaded: 1 new asset(s) in [Root / Parent Folder]');
  assert.ok(gas.fakes.properties.store.LAST_CHECK_TIME, 'keeps its state under LAST_CHECK_TIME');
});

test('checkForNewAssets uses WATCH_SUBJECT_PREFIX when it is set', () => {
  const folder = fakeFolder({ id: 'W', name: 'Images', files: [{ name: 'new.png', created: Date.now() - 60000 }] });
  const gas = load(FILES, {
    folders: { W: folder },
    properties: { WATCH_FOLDER_ID: 'W', WATCH_RECIPIENTS: 'a@example.com', WATCH_SUBJECT_PREFIX: '[Nytt]' },
  });
  gas.checkForNewAssets();
  assert.match(gas.fakes.mail[0].subject, /^\[Nytt\]: 1 new asset/);
});

test('checkForNewAssets fails clearly, and saves no check time, when its properties are not set', () => {
  const gas = load(FILES, { properties: { WATCH_RECIPIENTS: 'a@example.com' } });
  assert.throws(() => gas.checkForNewAssets(), /The script property WATCH_FOLDER_ID is not set/);
  assert.equal(gas.fakes.properties.store.LAST_CHECK_TIME, undefined);
});

test('runPostDeployTests: the DMC check warns when DMC_GCP_PROJECT is not set', () => {
  const gas = liveProject({ properties: { DMC_GCP_PROJECT: '' } });
  assert.equal(gas.runPostDeployTests(), '12 passed, 2 warnings, 0 failed');
  assert.ok(gas.fakes.logs.some((l) => /^WARN +DMC transfer: Cloud Run job +DMC_GCP_PROJECT is not set/.test(l)));
});

test('runPostDeployTests: the DMC check names the job, the server, and that the password is a secret', () => {
  const gas = liveProject();
  gas.runPostDeployTests();
  assert.ok(gas.fakes.logs.some((l) => l.includes('dmc-transfer in sesport/europe-north1 sends to ftp-no1.dmcproduction.se; never run')));
  const req = gas.fakes.urlFetch.requests.find((r) => r.url.startsWith('https://run.googleapis.com/'));
  assert.equal(req.url, 'https://run.googleapis.com/v2/projects/sesport/locations/europe-north1/jobs/dmc-transfer');
  assert.equal(req.options.method, 'get', 'only reads');
});

test('runPostDeployTests: the DMC check fails when the job writes to another spreadsheet', () => {
  const job = JSON.parse(JSON.stringify(DMC_JOB));
  job.template.template.containers[0].env[3].value = 'another-sheet';
  const gas = liveProject({ dmcJob: job });
  assert.throws(() => gas.runPostDeployTests(), /Post-deploy tests failed: DMC transfer/);
  assert.ok(gas.fakes.logs.some((l) => l.includes('writes its status to spreadsheet another-sheet, not this one')));
});

// --- Video Assets and sending to DMC ---------------------------------------------------------

test('"Video Assets" G and H are the columns the DMC code uses', () => {
  const gas = app();
  const headers = plain(gas.get('VIDEO_HEADERS'));
  assert.deepEqual(headers, VIDEO_HEADERS);
  assert.equal(headers[gas.get('DMC_SEND_COL') - 1], 'Skicka till DMC');
  assert.equal(headers[gas.get('DMC_STATUS_COL') - 1], 'DMC-status');
});

// A project with a Video folder (next to the images folder IMG) holding videos, and the
// "Video Assets" sheet holding rows (A-F, and G-H when given).
function videoProject({ videos, rows }) {
  const sheet = new FakeSheet('Video Assets', [VIDEO_HEADERS, ...rows]);
  const assets = fakeFolder({
    name: 'Assets',
    folders: [{ id: 'IMG', name: 'Bilder' }, { name: 'Video', files: videos.map((v) => Object.assign({ mimeType: 'video/mp4' }, v)) }],
  });
  const gas = load(FILES, {
    spreadsheet: fakeSpreadsheet([sheet]),
    properties: { DRIVE_FOLDER_ID: 'IMG', ELMO_URL: 'https://elmo', ELMO_KEY: 'k', ELMO_PROJECT: 'p', DMC_GCP_PROJECT: 'sesport' },
    folders: { IMG: assets.getFolders().next() },
    fetch: () => ({ code: 200, body: { videoMediaMetadata: { durationMillis: 1000 } } }),
  });
  return { gas, sheet };
}

const sixColumns = (id, name) => [id, name, '', '1', 'videos/' + name, '0:00:01:00'];

test('listDriveVideos keeps each row\'s "Skicka till DMC" and "DMC-status" with its video', () => {
  const { gas, sheet } = videoProject({
    videos: [{ id: 'a', name: 'a.mp4' }, { id: 'b', name: 'b.mp4' }, { id: 'c', name: 'c.mp4' }],
    rows: [
      [...sixColumns('b', 'b.mp4'), true, 'Skickad 2026-11-01 10:00'],
      [...sixColumns('gone', 'gone.mp4'), false, 'Fel: x'],
      sixColumns('a', 'a.mp4'), // listed before G-H existed
      [...sixColumns('x', 'x.mp4'), false, ''],
    ],
  });
  sheet.checkboxes.add('4,6'); // the old 4th row's checkbox
  gas.listDriveVideos();
  const rows = plain(sheet.data.slice(1, 5)).map((r) => [r[0], r[6], r[7]]);
  assert.deepEqual(rows, [
    ['a', false, ''],
    ['b', true, 'Skickad 2026-11-01 10:00'],
    ['c', false, ''],
    ['', '', ''], // the rows that are gone are cleared, checkboxes and all
  ]);
  assert.deepEqual([...sheet.checkboxes].sort(), ['1,6', '2,6', '3,6']);
});

// "Video Assets" rows A-H: [id, name, ..., ticked, status]
const videoRow = (id, name, ticked, status) => [...sixColumns(id, name), ticked, status || ''];

test('pickDmcClips_ picks the ticked rows that have a Drive file ID', () => {
  const gas = app();
  const pick = gas.pickDmcClips_([
    videoRow('a', 'a.mp4', true),
    videoRow('b', 'b.mp4', false),
    videoRow('', 'no-id.mp4', true),
    videoRow('c', 'c.mp4', true, 'Skickad 2026-11-01 10:00'), // sending again is fine
    videoRow('d', 'd.mp4', 'TRUE'), // text, not a ticked checkbox
  ]);
  assert.deepEqual(plain(pick), {
    clips: [{ row: 2, id: 'a', name: 'a.mp4' }, { row: 5, id: 'c', name: 'c.mp4' }],
    problems: [],
  });
});

test('pickDmcClips_ refuses clips on their way and two clips with the same name', () => {
  const gas = app();
  const { problems } = plain(gas.pickDmcClips_([
    videoRow('a', 'a.mp4', true, 'Köad 2026-11-17 14:00'),
    videoRow('b', 'b.mp4', true, 'Skickar 1.2 GB av 4.8 GB (25 %)'),
    videoRow('c1', 'c.mp4', true),
    videoRow('c2', 'c.mp4', true),
  ]));
  assert.equal(problems.length, 3);
  assert.match(problems[0], /^a.mp4 \(rad 2\) är redan på väg: "Köad 2026-11-17 14:00"/);
  assert.match(problems[1], /^b.mp4 \(rad 3\) är redan på väg/);
  assert.match(problems[2], /^c.mp4 finns på rad 4 och 5: samma namn hos DMC/);
});

function dmcProject(rows, { cloud = {}, sharing = { '*': sharedWith('writer') } } = {}) {
  const sheet = new FakeSheet('Video Assets', [VIDEO_HEADERS, ...rows]);
  const alerts = [];
  const gas = load(FILES, {
    spreadsheet: fakeSpreadsheet([sheet]),
    properties: { DMC_GCP_PROJECT: 'sesport' },
    fetch: cloudFetch(cloud),
    globals: {
      Session: { getActiveUser: () => ({ getEmail: () => 'producer@example.com' }) },
      Drive: fakeDriveSharing(sharing),
    },
  });
  // The UI, for alerts (the harness spreadsheet has toast but no getUi).
  gas.SpreadsheetApp.getUi = () => ({ alert: (...a) => alerts.push(a), ButtonSet: { OK: 'OK' } });
  return { gas, sheet, alerts };
}

test('sendCheckedToDmc starts the job with the ticked clips, one task each, then queues and unticks them', () => {
  const { gas, sheet } = dmcProject([
    videoRow('a', 'a.mp4', true), videoRow('b', 'b.mp4', false), videoRow('c', 'c.mp4', true),
  ]);
  gas.sendCheckedToDmc();
  const req = gas.fakes.urlFetch.requests.find((r) => r.url.endsWith(':run'));
  assert.equal(req.url, 'https://run.googleapis.com/v2/projects/sesport/locations/europe-north1/jobs/dmc-transfer:run');
  assert.deepEqual(JSON.parse(req.options.payload).overrides, {
    containerOverrides: [{ env: [{ name: 'FILE_IDS', value: 'a,c' }, { name: 'REQUESTED_BY', value: 'producer@example.com' }] }],
    taskCount: 2,
  });
  const gh = plain(sheet.data.slice(1)).map((r) => [r[6], r[7]]);
  assert.equal(gh[1].join('|'), 'false|', 'the unticked row is left alone');
  [gh[0], gh[2]].forEach(([ticked, status]) => {
    assert.equal(ticked, false);
    assert.match(status, /^Köad \d{4}-\d{2}-\d{2} \d{2}:\d{2}$/);
  });
});

test('sendCheckedToDmc uses DMC_REGION and DMC_JOB when they are set', () => {
  const { gas } = dmcProject([videoRow('a', 'a.mp4', true)]);
  gas.fakes.properties.setProperty('DMC_REGION', 'europe-west1').setProperty('DMC_JOB', 'dmc-test');
  gas.sendCheckedToDmc();
  assert.ok(gas.fakes.urlFetch.requests.some((r) => /\/locations\/europe-west1\/jobs\/dmc-test:run$/.test(r.url)));
});

test('sendCheckedToDmc keeps the clips ticked, with the reason, when the job does not start', () => {
  const { gas, sheet } = dmcProject([videoRow('a', 'a.mp4', true)], { cloud: { runCode: 403 } });
  assert.throws(() => gas.sendCheckedToDmc(), /Starting Cloud Run job dmc-transfer failed \(403\)/);
  assert.equal(sheet.data[1][6], true);
  assert.match(sheet.data[1][7], /^Fel: överföringen startade inte: Starting Cloud Run job dmc-transfer failed \(403\)/);
});

test('sendCheckedToDmc shows the problems and sends nothing when a clip cannot be sent', () => {
  const { gas, sheet, alerts } = dmcProject([
    videoRow('a', 'a.mp4', true), videoRow('b', 'b.mp4', true, 'Köad 2026-11-17 14:00'),
  ]);
  gas.sendCheckedToDmc();
  assert.equal(gas.fakes.urlFetch.requests.length, 0);
  assert.equal(alerts.length, 1);
  assert.match(alerts[0][1], /b.mp4 \(rad 3\) är redan på väg[\s\S]*Inget skickades\.$/);
  assert.equal(sheet.data[1][6], true, 'nothing was unticked');
});

test('sendCheckedToDmc does nothing when no clip is ticked, and fails clearly without a video list', () => {
  const { gas } = dmcProject([videoRow('a', 'a.mp4', false)]);
  gas.sendCheckedToDmc();
  assert.equal(gas.fakes.urlFetch.requests.length, 0);
  assert.ok(gas.fakes.logs.some((l) => l.includes('Bocka i "Skicka till DMC"')));
  const empty = load(FILES, { spreadsheet: fakeSpreadsheet([]) });
  assert.throws(() => empty.sendCheckedToDmc(), /Inga videor listade i "Video Assets"/);
});

test('runPostDeployTests: the DMC check shows the job\'s last run once it has run', () => {
  const never = liveProject({ dmcJob: Object.assign({}, DMC_JOB, { latestCreatedExecution: {} }) });
  never.runPostDeployTests();
  assert.ok(never.fakes.logs.some((l) => l.endsWith('; never run')));
  const ran = liveProject({ dmcJob: Object.assign({}, DMC_JOB, {
    latestCreatedExecution: { createTime: '2026-11-17T14:20:00Z', completionStatus: 'EXECUTION_SUCCEEDED' },
  }) });
  ran.runPostDeployTests();
  assert.ok(ran.fakes.logs.some((l) => l.endsWith('; last run 2026-11-17T14:20:00Z EXECUTION_SUCCEEDED')));
});

// --- DMC permissions ------------------------------------------------------------------------------

const TYPO = 'dmc-transfer@project.iam.gserviceaccount.com'; // the mistake made when setting up

test('sendCheckedToDmc sends nothing, and says what to share, when the spreadsheet is not shared with the job', () => {
  const { gas, sheet, alerts } = dmcProject([videoRow('a', 'a.mp4', true)], {
    sharing: { 'spreadsheet-id': sharedWith('writer', TYPO), a: sharedWith('reader') },
  });
  gas.sendCheckedToDmc();
  assert.ok(!gas.fakes.urlFetch.requests.some((r) => r.url.endsWith(':run')), 'the job was not started');
  assert.equal(alerts[0][1], `"Master schedule" är inte delad med ${SA} som redigerare. ` +
    'Dela den så att jobbet kan skriva DMC-status.\n\nInget skickades.');
  assert.deepEqual(plain(sheet.data[1].slice(6)), [true, ''], 'still ticked, not queued');
});

test('sendCheckedToDmc names each clip the job cannot read', () => {
  const { gas, alerts } = dmcProject([videoRow('a', 'a.mp4', true), videoRow('b', 'b.mp4', true)], {
    sharing: { 'spreadsheet-id': sharedWith('writer'), a: sharedWith('reader'), b: sharedWith('reader', TYPO) },
  });
  gas.sendCheckedToDmc();
  assert.match(alerts[0][1], new RegExp(`^"b.mp4" är inte delad med ${SA} som läsare\\. Dela den så att jobbet kan läsa klippet\\.`));
  assert.ok(!alerts[0][1].includes('"a.mp4"'));
});

test('sendCheckedToDmc tells a sender without permission to start the job which role they need', () => {
  const { gas, alerts } = dmcProject([videoRow('a', 'a.mp4', true)], { cloud: { granted: false } });
  gas.sendCheckedToDmc();
  assert.ok(!gas.fakes.urlFetch.requests.some((r) => r.url.endsWith(':run')));
  assert.match(alerts[0][1], /^producer@example.com får inte starta överföringen\. .*"Cloud Run Jobs Executor With Overrides" på jobbet dmc-transfer\./);
});

test('sendCheckedToDmc reports a job it cannot read', () => {
  const { gas, alerts } = dmcProject([videoRow('a', 'a.mp4', true)], { cloud: { job: null } });
  gas.sendCheckedToDmc();
  assert.match(alerts[0][1], /^Kan inte läsa överföringsjobbet dmc-transfer: Reading Cloud Run job dmc-transfer failed \(404\)/);
});

test('sendCheckedToDmc still sends when the sender may not see a clip\'s sharing', () => {
  const { gas, alerts } = dmcProject([videoRow('a', 'a.mp4', true)], {
    sharing: { 'spreadsheet-id': sharedWith('writer') }, // a: sharing not readable
  });
  gas.sendCheckedToDmc();
  assert.equal(alerts.length, 0);
  assert.ok(gas.fakes.urlFetch.requests.some((r) => r.url.endsWith(':run')));
});

test('dmcAccess_ accepts link sharing and a higher role than needed, and matches addresses in any case', () => {
  const { gas } = dmcProject([], {
    sharing: {
      'spreadsheet-id': [{ type: 'user', emailAddress: SA.toUpperCase(), role: 'owner' }],
      a: [{ type: 'anyone', role: 'reader' }],
      b: [{ type: 'user', emailAddress: SA, role: 'commenter' }],
    },
  });
  const access = gas.dmcAccess_([{ id: 'a', name: 'a.mp4', role: 'reader' }, { id: 'b', name: 'b.mp4', role: 'reader' }]);
  assert.deepEqual(plain(access.problems), []);
  assert.equal(access.account, SA);
});

test('runPostDeployTests: the DMC permissions check lists what it verified', () => {
  const gas = liveProject();
  gas.runPostDeployTests();
  assert.ok(gas.fakes.logs.some((l) => /^PASS +DMC transfer: permissions +you may start the job; /.test(l) &&
    l.endsWith(`${SA} can edit this spreadsheet, read "Video" and read the password`)), gas.fakes.logs.join('\n'));
});

test('runPostDeployTests: the DMC permissions check fails on a folder shared with the wrong address', () => {
  const gas = liveProject({ sharing: { 'spreadsheet-id': sharedWith('writer'), Video: sharedWith('reader', TYPO) } });
  assert.throws(() => gas.runPostDeployTests(), /Post-deploy tests failed: DMC transfer: permissions/);
  assert.ok(gas.fakes.logs.some((l) => l.includes(`"Video" is not shared with ${SA} as Viewer`)));
});

test('runPostDeployTests: the DMC permissions check fails when the job may not read the password', () => {
  const gas = liveProject({ cloud: { secretPolicy: { bindings: [
    { role: 'roles/secretmanager.secretAccessor', members: ['serviceAccount:' + TYPO] },
    { role: 'roles/secretmanager.viewer', members: ['serviceAccount:' + SA] },
  ] } } });
  assert.throws(() => gas.runPostDeployTests(), /DMC transfer: permissions/);
  assert.ok(gas.fakes.logs.some((l) => l.includes(`${SA} may not read the secret dmc-ftp-password`)));
});

test('runPostDeployTests: the DMC permissions check warns when it cannot see the password\'s access', () => {
  const gas = liveProject({ cloud: { secretPolicy: null } });
  assert.equal(gas.runPostDeployTests(), '13 passed, 1 warnings, 0 failed');
  assert.ok(gas.fakes.logs.some((l) => /^WARN +DMC transfer: permissions .*read "Video"; secret dmc-ftp-password not checked/.test(l)));
});

test('runPostDeployTests: the DMC permissions check fails when the runner may not start the job', () => {
  const gas = liveProject({ cloud: { granted: false } });
  assert.throws(() => gas.runPostDeployTests(), /DMC transfer: permissions/);
  assert.ok(gas.fakes.logs.some((l) => l.includes('user@example.com may not start the job; grant roles/run.jobsExecutorWithOverrides')));
});

// --- Who runs the watcher (WATCH_RUN_AS) ------------------------------------------------------------

// A project whose watcher runs as `runAs`, run by `me`, who has a checkForNewAssets trigger.
function watcherProject({ runAs, me = 'editor@example.com' }) {
  const folder = fakeFolder({ id: 'W', name: 'Images', files: [{ name: 'new.png', created: Date.now() - 60000 }] });
  const properties = { WATCH_FOLDER_ID: 'W', WATCH_RECIPIENTS: 'a@example.com' };
  if (runAs !== undefined) properties.WATCH_RUN_AS = runAs;
  const gas = load(FILES, {
    folders: { W: folder },
    properties,
    globals: { Session: { getScriptTimeZone: () => 'Europe/Stockholm', getEffectiveUser: () => ({ getEmail: () => me }) } },
  });
  gas.fakes.triggers.push({ getHandlerFunction: () => 'checkForNewAssets', getTriggerSourceId: () => null });
  gas.fakes.triggers.push({ getHandlerFunction: () => 'handleEdit', getTriggerSourceId: () => 'sheet' });
  return gas;
}
const handlers = (gas) => gas.fakes.triggers.map((t) => t.getHandlerFunction());

test('checkForNewAssets run by another account than WATCH_RUN_AS removes that account\'s trigger and does nothing else', () => {
  const gas = watcherProject({ runAs: 'owner@example.com' });
  gas.checkForNewAssets();
  assert.deepEqual(handlers(gas), ['handleEdit'], 'only the checkForNewAssets trigger is removed');
  assert.equal(gas.fakes.mail.length, 0);
  assert.equal(gas.fakes.properties.store.LAST_CHECK_TIME, undefined, 'the watcher\'s state is untouched');
  assert.ok(gas.fakes.logs.some((l) => l.includes("runs as owner@example.com, not editor@example.com: removed editor@example.com's checkForNewAssets trigger")));
});

test('checkForNewAssets runs the watcher for WATCH_RUN_AS (any case, spaces trimmed)', () => {
  const gas = watcherProject({ runAs: ' Editor@Example.com ' });
  gas.checkForNewAssets();
  assert.equal(gas.fakes.mail.length, 1);
  assert.deepEqual(handlers(gas), ['checkForNewAssets', 'handleEdit']);
});

test('checkForNewAssets runs for everyone without WATCH_RUN_AS, and when Google hides the address', () => {
  const unset = watcherProject({});
  unset.checkForNewAssets();
  assert.equal(unset.fakes.mail.length, 1);
  const hidden = watcherProject({ runAs: 'owner@example.com', me: '' });
  hidden.checkForNewAssets();
  assert.equal(hidden.fakes.mail.length, 1);
  assert.deepEqual(handlers(hidden), ['checkForNewAssets', 'handleEdit'], 'the trigger is kept');
});

test('runPostDeployTests: the trigger check follows WATCH_RUN_AS', () => {
  const other = liveProject({ properties: { WATCH_RUN_AS: 'owner@example.com' } });
  other.fakes.triggers.splice(1); // keep handleEdit only: the watcher's trigger is the owner's
  assert.equal(other.runPostDeployTests(), '14 passed, 0 warnings, 0 failed');
  assert.ok(other.fakes.logs.some((l) => l.includes('handleEdit and not checkForNewAssets (the watcher runs as owner@example.com)')));

  const leftover = liveProject({ properties: { WATCH_RUN_AS: 'owner@example.com' } });
  leftover.runPostDeployTests();
  assert.ok(leftover.fakes.logs.some((l) => /^WARN +Triggers +you have a checkForNewAssets trigger, but the watcher runs as owner@example.com/.test(l)));

  const unset = liveProject({ properties: { WATCH_RUN_AS: '' } });
  unset.runPostDeployTests();
  assert.ok(unset.fakes.logs.some((l) => /^WARN +Triggers +WATCH_RUN_AS is not set/.test(l)));
});
