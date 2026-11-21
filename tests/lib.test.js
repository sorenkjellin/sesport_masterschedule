// Unit tests for lib/lib.js. Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, plain, fakeBlob, fakeFolder, fakeSpreadsheet } = require('./harness');

const lib = (opts) => load(['lib/lib.js'], opts);

// --- Script Properties ---------------------------------------------------------

test('requiredProperty_ returns a set property and throws for a missing one', () => {
  const gas = lib({ properties: { ELMO_KEY: 'k1' } });
  assert.equal(gas.requiredProperty_('ELMO_KEY'), 'k1');
  assert.throws(() => gas.requiredProperty_('NOPE'), /The script property NOPE is not set/);
});

test('optionalProperty_ falls back when the property is missing or empty', () => {
  const gas = lib({ properties: { EMPTY: '', SET: 'x' } });
  assert.equal(gas.optionalProperty_('SET', 'd'), 'x');
  assert.equal(gas.optionalProperty_('EMPTY', 'd'), 'd');
  assert.equal(gas.optionalProperty_('NOPE', 'd'), 'd');
});

test('trimSlashes_ removes leading and trailing slashes only', () => {
  const gas = lib();
  assert.equal(gas.trimSlashes_('//images/genre//'), 'images/genre');
  assert.equal(gas.trimSlashes_('images'), 'images');
});

// --- HTTP ------------------------------------------------------------------------

test('checkResponse_ passes 2xx and throws on anything else, naming the request', () => {
  const gas = lib();
  const res = (code, body) => ({ getResponseCode: () => code, getContentText: () => body });
  const ok = res(204, '');
  assert.equal(gas.checkResponse_(ok, 'x'), ok);
  assert.throws(() => gas.checkResponse_(res(500, 'boom'), 'list project'), /list project failed \(500\): boom/);
  assert.throws(() => gas.checkResponse_(res(302, ''), 'x'), /x failed \(302\)/);
});

// --- Cachearoo ---------------------------------------------------------------------

test('cachearooUrl_ builds <base>/_data/<bucket>/<key>, encoded', () => {
  const gas = lib();
  assert.equal(gas.cachearooUrl_('https://c.example/', 'sesport', 'comingup'),
    'https://c.example/_data/sesport/comingup');
  assert.equal(gas.cachearooUrl_('https://c.example', 'se sport', 'a/b'),
    'https://c.example/_data/se%20sport/a%2Fb');
});

test('cachearooUrl_ uses a base that already points at a value as it is', () => {
  const gas = lib();
  const full = 'https://c.example/_data/other/key';
  assert.equal(gas.cachearooUrl_(full, 'sesport', 'comingup'), full);
});

test('cachearooWrite_ PUTs the value as JSON with both API key headers', () => {
  const gas = lib({ fetch: () => ({ code: 200 }) });
  gas.cachearooWrite_('https://c.example', 'key1', 'sesport', 'comingup', { shows: [1] });
  const [req] = gas.fakes.urlFetch.requests;
  assert.equal(req.url, 'https://c.example/_data/sesport/comingup');
  assert.equal(req.options.method, 'put');
  assert.equal(req.options.contentType, 'application/json');
  assert.equal(req.options.headers['API-Key'], 'key1');
  assert.equal(req.options.headers['x-api-key'], 'key1');
  assert.deepEqual(JSON.parse(req.options.payload), { shows: [1] });
});

test('cachearooWrite_ throws when Cachearoo refuses', () => {
  const gas = lib({ fetch: () => ({ code: 403, body: 'bad key' }) });
  assert.throws(() => gas.cachearooWrite_('https://c', 'k', 'b', 'k', {}), /Cachearoo write failed \(403\): bad key/);
});

// --- Elmo file API -------------------------------------------------------------------

const elmo = { elmoUrl: 'https://elmo.example', key: 'ek', project: 'my proj' };

test('getElmoFiles_ maps file paths to hash and URL, skipping folders', () => {
  const gas = lib({
    fetch: () => ({
      code: 200,
      body: {
        list: [
          { name: 'images', folder: true },
          { name: 'images/a.png', url: 'https://cdn/x/abc123' },
          { name: 'images/no-url.png' },
        ],
      },
    }),
  });
  assert.deepEqual(plain(gas.getElmoFiles_(elmo)), {
    'images/a.png': { hash: 'abc123', url: 'https://cdn/x/abc123' },
  });
  const [req] = gas.fakes.urlFetch.requests;
  assert.equal(req.url, 'https://elmo.example/api/files/projects/my%20proj/?recurse=1');
  assert.equal(req.options.headers['x-api-key'], 'ek');
});

test('getElmoFiles_ treats a missing project (404) as empty', () => {
  const gas = lib({ fetch: () => ({ code: 404 }) });
  assert.deepEqual(plain(gas.getElmoFiles_(elmo)), {});
});

test('uploadToElmo_ opens a session, PUTs to S3 without the API key, completes and returns the URL', () => {
  const gas = lib({
    fetch(url) {
      if (url.endsWith('/upload-session')) return { code: 200, body: { id: 's1', url: 'https://s3/put?sig' } };
      if (url.startsWith('https://s3/')) return { code: 200 };
      if (url.endsWith('/upload-session/s1/complete')) return { code: 200, body: { url: 'https://cdn/hash' } };
      return { code: 500 };
    },
  });
  const url = gas.uploadToElmo_(elmo, fakeBlob('img', 'image/png'), 'images/a.png');
  assert.equal(url, 'https://cdn/hash');

  const [open, s3, done] = gas.fakes.urlFetch.requests;
  assert.deepEqual(JSON.parse(open.options.payload), { project: 'my proj', path: 'images/a.png', contentType: 'image/png' });
  assert.equal(s3.url, 'https://s3/put?sig');
  assert.equal(s3.options.method, 'put');
  assert.equal(s3.options.contentType, 'image/png'); // part of the presigned signature
  assert.equal(s3.options.headers, undefined);
  assert.equal(done.url, 'https://elmo.example/api/files/upload-session/s1/complete');
});

test('uploadToElmo_ sends octet-stream when the blob has no content type', () => {
  const gas = lib({
    fetch: (url) => (url.endsWith('complete') ? { code: 200, body: { url: 'u' } }
      : { code: 200, body: { id: 's', url: 'https://s3/x' } }),
  });
  gas.uploadToElmo_(elmo, fakeBlob('x', null), 'a.bin');
  assert.equal(gas.fakes.urlFetch.requests[1].options.contentType, 'application/octet-stream');
});

test('uploadToElmo_ stops when the S3 upload fails', () => {
  const gas = lib({
    fetch: (url) => (url.startsWith('https://s3/') ? { code: 403, body: 'denied' }
      : { code: 200, body: { id: 's', url: 'https://s3/x' } }),
  });
  assert.throws(() => gas.uploadToElmo_(elmo, fakeBlob('x', 'image/png'), 'a.png'), /upload to S3 failed \(403\)/);
  assert.equal(gas.fakes.urlFetch.requests.length, 2);
});

test('elmoFileUrl_ gives the file API path, keeps full URLs, and is empty for no path', () => {
  const gas = lib();
  assert.equal(gas.elmoFileUrl_('', 'p'), '');
  assert.equal(gas.elmoFileUrl_('https://x/y.png', 'p'), 'https://x/y.png');
  assert.equal(gas.elmoFileUrl_('/images/Genre/Hand boll.png', 'my proj'),
    '/api/files/projects/my%20proj/images/Genre/Hand%20boll.png');
});

// --- Mail ---------------------------------------------------------------------------

test('sendMail_ sends through GmailApp and leaves out empty options', () => {
  const gas = lib();
  const blob = fakeBlob('x', 'text/plain');
  gas.sendMail_('a@example.com', 'Subject', 'Body', { cc: '', bcc: null, name: 'Sender', attachments: [blob] });
  const [m] = gas.fakes.mail;
  assert.equal(m.to, 'a@example.com');
  assert.equal(m.subject, 'Subject');
  assert.deepEqual(Object.keys(m.options).sort(), ['attachments', 'name']);
});

test('sendMail_ throws when there is no recipient', () => {
  const gas = lib();
  assert.throws(() => gas.sendMail_('', 'Report', 'x'), /No recipient for the email "Report"/);
  assert.equal(gas.fakes.mail.length, 0);
});

// --- Spreadsheets -----------------------------------------------------------------------

test('findColumns_ matches headers ignoring case and spaces, first alternative wins, -1 when missing', () => {
  const gas = lib();
  const cols = gas.findColumns_([' Datum ', 'PROGAMTITEL', 'Programtitel'], {
    date: ['Datum'], series: ['Programtitel', 'Progamtitel'], typo: ['Progamtitel'], none: ['Nope'],
  });
  assert.deepEqual(plain(cols), { date: 0, series: 2, typo: 1, none: -1 });
});

test('columnLetter_ gives spreadsheet column letters', () => {
  const gas = lib();
  assert.deepEqual([1, 19, 26, 27, 52, 703].map((n) => gas.columnLetter_(n)), ['A', 'S', 'Z', 'AA', 'AZ', 'AAA']);
});

test('readTextRows_ / writeTextRows_ / writeTextRow_ work below the header within the width', () => {
  const { FakeSheet } = require('./harness');
  const sheet = new FakeSheet('S', [['A', 'B', 'keep'], ['1', '2', 'x'], ['3', '4', 'y']]);
  const gas = lib({ spreadsheet: fakeSpreadsheet([sheet]) });
  assert.deepEqual(plain(gas.readTextRows_(sheet, 2)), [['1', '2'], ['3', '4']]);
  gas.writeTextRows_(sheet, [['5', '6']], 2);
  assert.deepEqual(plain(gas.readTextRows_(sheet, 3)), [['5', '6', 'x'], ['', '', 'y']]);
  gas.writeTextRow_(sheet, 1, ['7', '8']);
  assert.deepEqual(sheet.data[2].slice(0, 2), ['7', '8']);
});

test('readTextRows_ is empty for a sheet with only a header', () => {
  const { FakeSheet } = require('./harness');
  const gas = lib();
  assert.deepEqual(plain(gas.readTextRows_(new FakeSheet('S', [['A']]), 1)), []);
});

test('getOrCreateSheet_ creates a missing sheet with its headers', () => {
  const ss = fakeSpreadsheet([]);
  const gas = lib({ spreadsheet: ss });
  const sheet = gas.getOrCreateSheet_('New', ['H1', 'H2']);
  assert.equal(ss.getSheetByName('New'), sheet);
  assert.deepEqual(sheet.data[0], ['H1', 'H2']);
});

test('valuesToXlsx_ refuses an empty table', () => {
  const gas = lib();
  assert.throws(() => gas.valuesToXlsx_([], 'S', 'x.xlsx'), /Nothing to export to x.xlsx/);
});

// --- Drive ------------------------------------------------------------------------------

test('extractDriveId reads /d/<id> and id=<id> links', () => {
  const gas = lib();
  assert.equal(gas.extractDriveId('https://docs.google.com/spreadsheets/d/1Ab_c-D/edit#gid=0'), '1Ab_c-D');
  assert.equal(gas.extractDriveId('https://drive.google.com/open?id=XyZ_9'), 'XyZ_9');
  assert.equal(gas.extractDriveId('https://example.com/page'), null);
});

test('listDriveFiles_ recurses, filters by type and replaces "/" in names', () => {
  const gas = lib();
  const folder = fakeFolder({
    name: 'Bilder',
    files: [
      { name: 'b.png', mimeType: 'image/png', size: 10 },
      { name: 'notes.txt', mimeType: 'text/plain' },
    ],
    folders: [{
      name: 'Genre/Old',
      files: [{ name: 'x/y.jpg', mimeType: 'image/jpeg' }],
      folders: [{ name: 'Deep', files: [{ name: 'a.png', mimeType: 'image/png' }] }],
    }],
  });
  const found = plain(gas.listDriveFiles_(folder, 'image/'));
  assert.deepEqual(found.map((f) => f.relPath).sort(), ['Genre-Old/Deep/a.png', 'Genre-Old/x-y.jpg', 'b.png']);
  const b = found.find((f) => f.relPath === 'b.png');
  assert.equal(b.name, 'b.png');
  assert.equal(b.size, 10);
  assert.equal(plain(gas.listDriveFiles_(folder, '')).length, 4);
});

test('findSiblingFolder_ finds a folder next to the given one, or null', () => {
  const gas = lib();
  const parent = fakeFolder({ name: 'Assets', folders: [{ name: 'Bilder' }, { name: 'Video' }] });
  const bilder = parent.getFolders().next();
  assert.equal(gas.findSiblingFolder_(bilder, 'Video').getName(), 'Video');
  assert.equal(gas.findSiblingFolder_(bilder, 'Ljud'), null);
});

test('getGrandparentFolder_ goes two levels up, else falls back to My Drive', () => {
  const gas = lib();
  const top = fakeFolder({ name: 'Produktion', folders: [{ name: 'Mallar' }] });
  const mallar = top.getFolders().next();
  const file = { getParents: () => ({ hasNext: () => true, next: () => mallar }) };
  assert.equal(gas.getGrandparentFolder_(file).getName(), 'Produktion');
  const orphan = { getParents: () => ({ hasNext: () => false }) };
  assert.equal(gas.getGrandparentFolder_(orphan).getName(), 'My Drive');
});

test('driveVideoDurationsMs_ reads durations and gives null for unprocessed videos or errors', () => {
  const gas = lib({
    fetch(url) {
      if (url.includes('/v1?')) return { code: 200, body: { videoMediaMetadata: { durationMillis: '61040' } } };
      if (url.includes('/v2?')) return { code: 200, body: {} };
      return { code: 404 };
    },
  });
  assert.deepEqual(plain(gas.driveVideoDurationsMs_(['v1', 'v2', 'v3'])), { v1: 61040, v2: null, v3: null });
});

test('driveVideoDurationsMs_ asks in batches of 50', () => {
  const gas = lib({ fetch: () => ({ code: 200, body: {} }) });
  let calls = 0;
  const fetchAll = gas.UrlFetchApp.fetchAll;
  gas.UrlFetchApp.fetchAll = (list) => { calls++; assert.ok(list.length <= 50); return fetchAll(list); };
  gas.driveVideoDurationsMs_(Array.from({ length: 120 }, (_, i) => 'id' + i));
  assert.equal(calls, 3);
});

// --- Drive folder watcher --------------------------------------------------------------

function watched(now) {
  const ago = (min) => now - min * 60000;
  return fakeFolder({
    id: 'F1',
    name: 'Assets',
    files: [
      { name: 'new-root.png', mimeType: 'image/png', created: ago(30), size: 1048576 },
      { name: 'old-root.png', mimeType: 'image/png', created: ago(120) },
    ],
    folders: [{
      name: 'Promo',
      files: [{ name: 'new-promo.mp4', mimeType: 'video/mp4', created: ago(5) }],
      folders: [{ name: 'Deep', files: [{ name: 'too-deep.png', created: ago(1) }] }],
    }],
  });
}

test('watchDriveFolder_ on its first run looks back an hour, mails the new files and saves the check time', () => {
  const now = Date.now();
  const gas = lib({ folders: { F1: watched(now) } });
  const found = plain(gas.watchDriveFolder_({ folderId: 'F1', recipients: 'a@example.com', subjectPrefix: '[Alert]' }));

  assert.deepEqual(found.map((f) => f.name).sort(), ['new-promo.mp4', 'new-root.png']); // not old, not too deep
  assert.equal(found.find((f) => f.name === 'new-root.png').folderName, 'Root / Parent Folder');
  assert.equal(found.find((f) => f.name === 'new-root.png').sizeMb, '1.00');

  const [m] = gas.fakes.mail;
  assert.equal(m.to, 'a@example.com');
  assert.equal(m.subject, '[Alert]: 2 new asset(s) in [Root / Parent Folder, Promo]');
  assert.match(m.body, /"Assets"/);
  assert.match(m.body, /SUBFOLDER: Promo \(1 new asset\)/);
  assert.match(m.body, /File link: https:\/\/drive\/new-promo.mp4/);

  const saved = Date.parse(gas.fakes.properties.store.DRIVE_WATCH_LAST_CHECK_F1);
  assert.ok(Math.abs(saved - now) < 5000);
});

test('watchDriveFolder_ only reports files created after the last check, and sends nothing when none are', () => {
  const now = Date.now();
  const since = new Date(now - 10 * 60000).toISOString();
  const gas = lib({ folders: { F1: watched(now) }, properties: { MY_KEY: since } });
  const found = plain(gas.watchDriveFolder_({ folderId: 'F1', recipients: 'a@example.com', stateKey: 'MY_KEY' }));
  assert.deepEqual(found.map((f) => f.name), ['new-promo.mp4']);
  assert.equal(gas.fakes.mail[0].subject, 'New files: 1 new asset(s) in [Promo]');

  const again = plain(gas.watchDriveFolder_({ folderId: 'F1', recipients: 'a@example.com', stateKey: 'MY_KEY' }));
  assert.deepEqual(again, []);
  assert.equal(gas.fakes.mail.length, 1);
});

test('watchDriveFolder_ runs beforeMail with the new files first and puts its text above the list', () => {
  const now = Date.now();
  const gas = lib({ folders: { F1: watched(now) } });
  const seen = [];
  gas.watchDriveFolder_({ folderId: 'F1', recipients: 'a@example.com', beforeMail: (files) => {
    seen.push(files.length, gas.fakes.mail.length);
    return 'Updated: https://sheet\n';
  } });
  assert.deepEqual(seen, [2, 0], 'called with the new files, before any mail');
  assert.match(gas.fakes.mail[0].body, /in "Assets"\.\n\nUpdated: https:\/\/sheet\n\n=====/);

  const quiet = lib({ folders: { F1: watched(now) }, properties: { DRIVE_WATCH_LAST_CHECK_F1: new Date(now).toISOString() } });
  let called = false;
  quiet.watchDriveFolder_({ folderId: 'F1', recipients: 'a@example.com', beforeMail: () => { called = true; } });
  assert.equal(called, false, 'not called when nothing is new');
  const plainMail = lib({ folders: { F1: watched(now) } });
  plainMail.watchDriveFolder_({ folderId: 'F1', recipients: 'a@example.com' });
  assert.match(plainMail.fakes.mail[0].body, /in "Assets"\.\n\n=====/, 'no note without beforeMail');
});

// --- Hashing, time and dates -------------------------------------------------------------

test('sha256Hex_ gives the lowercase hex digest', () => {
  const gas = lib();
  const bytes = Array.from(Buffer.from('abc'));
  assert.equal(gas.sha256Hex_(bytes), 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad');
});

test('formatTimecode_ formats H:MM:SS:FF at 25 fps by default, dropping partial frames', () => {
  const gas = lib();
  assert.equal(gas.formatTimecode_(0), '0:00:00:00');
  assert.equal(gas.formatTimecode_(3723960), '1:02:03:24');
  assert.equal(gas.formatTimecode_(39), '0:00:00:00');
  assert.equal(gas.formatTimecode_(36000000), '10:00:00:00');
  assert.equal(gas.formatTimecode_(1999, 30), '0:00:01:29');
});

test('parseLocalDateTime_ reads both date formats and both time separators in the given time zone', () => {
  const gas = lib();
  const iso = (d) => d.toISOString();
  assert.equal(iso(gas.parseLocalDateTime_('2026-11-07', '17:50', 'Europe/Stockholm')), '2026-11-07T16:50:00.000Z');
  assert.equal(iso(gas.parseLocalDateTime_('7.11.2026', '9.05', 'Europe/Stockholm')), '2026-11-07T08:05:00.000Z');
  assert.equal(iso(gas.parseLocalDateTime_('2026-7-1', '17.50', 'Europe/Stockholm')), '2026-07-01T15:50:00.000Z'); // summer time
  assert.equal(iso(gas.parseLocalDateTime_('2026-11-07', '17:50:30', 'Europe/Stockholm')), '2026-11-07T16:50:00.000Z');
});

test('parseLocalDateTime_ gives null for what it cannot read', () => {
  const gas = lib();
  for (const [d, t] of [['2026/11/07', '17:50'], ['2026-11-07', 'kväll'], ['', ''], ['7.11.26', '17:50'], ['2026-11-07', '']]) {
    assert.equal(gas.parseLocalDateTime_(d, t, 'Europe/Stockholm'), null, `${d} ${t}`);
  }
});

// --- UI ------------------------------------------------------------------------------------

test('escapeHtml_ escapes the five HTML characters', () => {
  const gas = lib();
  assert.equal(gas.escapeHtml_(`<a href="x">Tom & 'Jerry'</a>`),
    '&lt;a href=&quot;x&quot;&gt;Tom &amp; &#39;Jerry&#39;&lt;/a&gt;');
});

test('toast_ logs, and does not fail where there is no UI', () => {
  const gas = lib({ globals: { SpreadsheetApp: { getActiveSpreadsheet() { throw new Error('no UI'); } } } });
  gas.toast_('hello');
  assert.deepEqual(plain(gas.fakes.logs), ['hello']);
});

// --- Triggers -------------------------------------------------------------------------------

test('ensureEditTrigger_ adds one trigger per handler and spreadsheet', () => {
  const gas = lib();
  const a = fakeSpreadsheet([], 'A');
  const b = fakeSpreadsheet([], 'B');
  gas.ensureEditTrigger_('handleEdit', a);
  gas.ensureEditTrigger_('handleEdit', a);
  gas.ensureEditTrigger_('handleEdit', b);
  gas.ensureEditTrigger_('other', a);
  assert.equal(gas.fakes.triggers.length, 3);
});

test('deleteTriggers_ removes every trigger for the handler and nothing else', () => {
  const gas = lib();
  gas.ensureEditTrigger_('handleEdit', fakeSpreadsheet([], 'A'));
  gas.ensureEditTrigger_('handleEdit', fakeSpreadsheet([], 'B'));
  gas.ensureEditTrigger_('other', fakeSpreadsheet([], 'A'));
  gas.deleteTriggers_('handleEdit');
  assert.deepEqual(gas.fakes.triggers.map((t) => t.getHandlerFunction()), ['other']);
});

// --- Cloud Run jobs -----------------------------------------------------------------

const JOB = { project: 'sesport', region: 'europe-north1', name: 'dmc-transfer' };
const JOB_URL = 'https://run.googleapis.com/v2/projects/sesport/locations/europe-north1/jobs/dmc-transfer';

test('runCloudRunJob_ starts the job with the user\'s token and overrides, and returns the execution', () => {
  const gas = lib({ fetch: () => ({ code: 200, body: { name: 'operations/1', metadata: { name: JOB_URL + '/executions/x1' } } }) });
  const execution = gas.runCloudRunJob_(JOB, { env: { FILE_IDS: 'a,b', N: 2 }, taskCount: 2 });
  assert.equal(execution, JOB_URL + '/executions/x1');
  const [req] = gas.fakes.urlFetch.requests;
  assert.equal(req.url, JOB_URL + ':run');
  assert.equal(req.options.method, 'post');
  assert.equal(req.options.headers.Authorization, 'Bearer token');
  assert.deepEqual(JSON.parse(req.options.payload), {
    overrides: { containerOverrides: [{ env: [{ name: 'FILE_IDS', value: 'a,b' }, { name: 'N', value: '2' }] }], taskCount: 2 },
  });
});

test('runCloudRunJob_ without options runs the job as it is deployed', () => {
  const gas = lib({ fetch: () => ({ code: 200, body: { name: 'operations/1' } }) });
  assert.equal(gas.runCloudRunJob_(JOB), 'operations/1');
  assert.deepEqual(JSON.parse(gas.fakes.urlFetch.requests[0].options.payload), { overrides: {} });
});

test('runCloudRunJob_ and getCloudRunJob_ fail with Cloud Run\'s answer', () => {
  const gas = lib({ fetch: () => ({ code: 403, body: 'Permission \'run.jobs.runWithOverrides\' denied' }) });
  assert.throws(() => gas.runCloudRunJob_(JOB), /Starting Cloud Run job dmc-transfer failed \(403\): Permission 'run.jobs.runWithOverrides' denied/);
  assert.throws(() => gas.getCloudRunJob_(JOB), /Reading Cloud Run job dmc-transfer failed \(403\)/);
});

test('getCloudRunJob_ reads the job', () => {
  const gas = lib({ fetch: (url) => ({ code: 200, body: { name: url } }) });
  assert.equal(gas.getCloudRunJob_(JOB).name, JOB_URL);
  assert.equal(gas.fakes.urlFetch.requests[0].options.method, 'get');
});

test('cloudRunJobPermissions_ asks Cloud Run which permissions the user has on the job', () => {
  const gas = lib({ fetch: () => ({ code: 200, body: { permissions: ['run.jobs.get'] } }) });
  assert.deepEqual(plain(gas.cloudRunJobPermissions_(JOB, ['run.jobs.get', 'run.jobs.runWithOverrides'])), ['run.jobs.get']);
  const [req] = gas.fakes.urlFetch.requests;
  assert.equal(req.url, JOB_URL + ':testIamPermissions');
  assert.deepEqual(JSON.parse(req.options.payload), { permissions: ['run.jobs.get', 'run.jobs.runWithOverrides'] });
  const none = lib({ fetch: () => ({ code: 200, body: {} }) });
  assert.deepEqual(plain(none.cloudRunJobPermissions_(JOB, ['run.jobs.run'])), []);
});

test('getSecretIamPolicy_ reads the policy of the secret', () => {
  const gas = lib({ fetch: () => ({ code: 200, body: { bindings: [] } }) });
  assert.deepEqual(plain(gas.getSecretIamPolicy_('sesport', 'dmc-ftp-password')), { bindings: [] });
  assert.equal(gas.fakes.urlFetch.requests[0].url,
    'https://secretmanager.googleapis.com/v1/projects/sesport/secrets/dmc-ftp-password:getIamPolicy');
});

test('driveRoleOf_ gives the highest role an address has, directly or by link, over every page', () => {
  const pages = {
    undefined: { permissions: [{ type: 'user', emailAddress: 'a@x.com', role: 'reader' }], nextPageToken: 'p2' },
    p2: { permissions: [{ type: 'user', emailAddress: 'A@x.com', role: 'writer' }, { type: 'domain', role: 'owner' }] },
  };
  const calls = [];
  const gas = lib({ globals: { Drive: { Permissions: { list: (id, opts) => { calls.push(opts); return pages[opts.pageToken]; } } } } });
  assert.equal(gas.driveRoleOf_('F', 'a@x.com'), 'writer');
  assert.equal(calls.length, 2);
  assert.equal(calls[0].supportsAllDrives, true);
  assert.equal(gas.driveRoleOf_('F', 'b@x.com'), null, 'a domain share does not count');
  const link = lib({ globals: { Drive: { Permissions: { list: () => ({ permissions: [{ type: 'anyone', role: 'reader' }] }) } } } });
  assert.equal(link.driveRoleOf_('F', 'b@x.com'), 'reader');
});

test('driveRoleRank_ orders the Drive roles', () => {
  const gas = lib();
  const ranks = [null, 'reader', 'commenter', 'writer', 'fileOrganizer', 'organizer', 'owner'].map(gas.driveRoleRank_);
  assert.deepEqual(ranks, [0, 1, 1, 2, 3, 3, 4]);
});
