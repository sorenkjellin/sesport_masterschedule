// Tests for the DMC transfer job: real basic-ftp against a fake FTPS server (explicit TLS on the
// control and data connections), with Drive and the sheet faked. Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { Readable } = require('node:stream');
const { FakeFtpServer } = require('./fakeFtpServer');
const { configFromEnv, taskFileIds, ftpName, progressText, stamp, sendClip } = require('../src/transfer');
const { googleApi, columnLetter } = require('../src/google');

const CLIP = Buffer.alloc(3 * 1024 * 1024 + 17, 7); // a little over 3 MB
const NOW = new Date('2026-11-17T13:32:00Z'); // 14:32 in Stockholm

// Drive with one clip per entry: { id: { name, content } }.
function fakeDrive(files) {
  return {
    metadata: async (id) => {
      if (!files[id]) throw new Error('File not found: ' + id);
      return { name: files[id].name, size: files[id].content.length };
    },
    download: async (id) => Readable.from([files[id].content]),
  };
}

// Starts a fake server, sends fileId with the given settings and returns what happened.
async function send({ server = {}, ftp = {}, files = { f1: { name: 'Match 1.mp4', content: CLIP } },
  fileId = 'f1', progressEveryMs } = {}) {
  const srv = new FakeFtpServer(Object.assign({ users: { dmc: 'secret' }, dirs: ['/', '/incoming'] }, server));
  const port = await srv.start();
  const statuses = [];
  let result;
  let error;
  try {
    result = await sendClip(fileId, {
      drive: fakeDrive(files),
      status: async (id, text) => statuses.push(id + ': ' + text),
      ftp: Object.assign({ host: 'localhost', port, user: 'dmc', password: 'secret', dir: '/incoming', tempSuffix: '.part' }, ftp),
      tlsOptions: { ca: FakeFtpServer.cert },
      now: () => NOW,
      log: () => {},
      progressEveryMs,
    });
  } catch (e) {
    error = e;
  } finally {
    await srv.stop();
  }
  return { srv, statuses, result, error };
}

// --- Sending ------------------------------------------------------------------------------

test('sends a clip over explicit TLS into the folder, under its Drive name', async () => {
  const { srv, statuses, result, error } = await send();
  assert.equal(error, undefined);
  assert.deepEqual(result, { name: 'Match 1.mp4', size: CLIP.length });
  assert.ok(srv.files.get('/incoming/Match 1.mp4').equals(CLIP));
  assert.deepEqual([...srv.files.keys()], ['/incoming/Match 1.mp4'], 'no .part file is left');
  assert.deepEqual(statuses, ['f1: Skickar', 'f1: Skickad 2026-11-17 14:32']);
  // TLS came first; the password was sent after it, and data went over TLS (PROT P).
  assert.equal(srv.commands[0], 'AUTH TLS');
  assert.ok(srv.commands.indexOf('PASS ***') > 0);
  assert.ok(srv.commands.includes('PROT P'));
});

test('uploads under a temporary name and renames it when complete', async () => {
  const { srv } = await send();
  assert.ok(srv.commands.includes('STOR Match 1.mp4.part'));
  assert.ok(srv.commands.includes('RNTO Match 1.mp4'));
});

test('FTP_TEMP_SUFFIX "none" uploads under the final name directly', async () => {
  const { srv, error } = await send({ ftp: { tempSuffix: configFromEnv(env({ FTP_TEMP_SUFFIX: 'none' })).ftp.tempSuffix } });
  assert.equal(error, undefined);
  assert.ok(srv.commands.includes('STOR Match 1.mp4'));
  assert.ok(!srv.commands.some((c) => c.startsWith('RNFR')));
});

test('overwrites a file with the same name at DMC', async () => {
  const { srv, error } = await send({ server: { files: { '/incoming/Match 1.mp4': 'old version' } } });
  assert.equal(error, undefined);
  assert.ok(srv.files.get('/incoming/Match 1.mp4').equals(CLIP));
});

test('overwrites on a server that will not rename onto an existing file', async () => {
  const { srv, error } = await send({
    server: { renameOverwrites: false, files: { '/incoming/Match 1.mp4': 'old version' } },
  });
  assert.equal(error, undefined);
  assert.ok(srv.files.get('/incoming/Match 1.mp4').equals(CLIP));
  assert.ok(srv.commands.includes('DELE Match 1.mp4'));
});

test('a "/" in the Drive name becomes "-" (no subfolders at DMC)', async () => {
  const { srv } = await send({ files: { f1: { name: 'Final 1/2.mp4', content: CLIP } } });
  assert.ok(srv.files.has('/incoming/Final 1-2.mp4'));
});

test('writes progress while sending', async () => {
  const { statuses } = await send({ progressEveryMs: 0 });
  assert.ok(statuses.some((s) => /^f1: Skickar \d+ MB av 3 MB \(\d+ %\)$/.test(s)), statuses.join('\n'));
  assert.equal(statuses[statuses.length - 1], 'f1: Skickad 2026-11-17 14:32');
});

// --- Failing ------------------------------------------------------------------------------

test('a wrong password fails with the server\'s answer in the status, nothing is stored', async () => {
  const { srv, statuses, error } = await send({ ftp: { password: 'wrong' } });
  assert.match(error.message, /530 Login incorrect/);
  assert.equal(statuses[statuses.length - 1], 'f1: Fel: 530 Login incorrect.');
  assert.equal(srv.files.size, 0);
});

test('a server without TLS is refused before the password is sent', async () => {
  const { srv, statuses, error } = await send({ server: { tls: false } });
  assert.ok(error);
  assert.match(statuses[statuses.length - 1], /^f1: Fel: /);
  assert.ok(!srv.commands.some((c) => c.startsWith('PASS')), 'the password never went out in the clear');
});

test('a size mismatch after the upload fails and removes the partial file', async () => {
  const { srv, statuses, error } = await send({ server: { truncateUploads: true } });
  assert.match(error.message, /DMC has \d+ bytes of Match 1.mp4, Drive has 3145745/);
  assert.match(statuses[statuses.length - 1], /^f1: Fel: DMC has/);
  assert.equal(srv.files.size, 0);
});

test('a missing folder at DMC fails clearly', async () => {
  const { statuses, error } = await send({ ftp: { dir: '/nowhere' } });
  assert.match(error.message, /550/);
  assert.match(statuses[statuses.length - 1], /^f1: Fel: 550/);
});

test('a file Drive cannot find fails without connecting', async () => {
  const { srv, statuses, error } = await send({ fileId: 'gone' });
  assert.match(error.message, /File not found: gone/);
  assert.equal(statuses[statuses.length - 1], 'gone: Fel: File not found: gone');
  assert.equal(srv.commands.length, 0);
});

test('a failing status write does not stop the transfer', async () => {
  const srv = new FakeFtpServer({ dirs: ['/'] });
  const port = await srv.start();
  const logs = [];
  try {
    await sendClip('f1', {
      drive: fakeDrive({ f1: { name: 'a.mp4', content: CLIP } }),
      status: async () => { throw new Error('Sheets is down'); },
      ftp: { host: 'localhost', port, user: 'dmc', password: 'secret', dir: '', tempSuffix: '' },
      tlsOptions: { ca: FakeFtpServer.cert },
      log: (m) => logs.push(m),
    });
  } finally {
    await srv.stop();
  }
  assert.ok(srv.files.get('/a.mp4').equals(CLIP));
  assert.ok(logs.some((l) => l.includes('Could not write the status of f1 ("Skickar"): Sheets is down')));
});

// --- Settings and helpers ---------------------------------------------------------------------

function env(extra) {
  return Object.assign({
    FTP_HOST: 'ftp-no1.dmcproduction.se', FTP_USER: 'u', FTP_PASSWORD: 'p', SPREADSHEET_ID: 's', FILE_IDS: 'a',
  }, extra);
}

test('configFromEnv names every missing setting, and has defaults for the rest', () => {
  assert.throws(() => configFromEnv({ FTP_HOST: 'h' }), /Not set: FTP_USER, FTP_PASSWORD, SPREADSHEET_ID, FILE_IDS/);
  const cfg = configFromEnv(env());
  assert.deepEqual(cfg.ftp, { host: 'ftp-no1.dmcproduction.se', port: 21, user: 'u', password: 'p', dir: '', tempSuffix: '.part' });
  assert.equal(cfg.sheetName, 'Video Assets');
  assert.equal(cfg.statusHeader, 'DMC-status');
  assert.equal(cfg.timeZone, 'Europe/Stockholm');
});

test('taskFileIds shares the clips between the tasks, all of them to one task by default', () => {
  assert.deepEqual(taskFileIds({ FILE_IDS: 'a, b,c,' }), ['a', 'b', 'c']);
  const tasks = [0, 1].map((i) => taskFileIds({ FILE_IDS: 'a,b,c', CLOUD_RUN_TASK_INDEX: String(i), CLOUD_RUN_TASK_COUNT: '2' }));
  assert.deepEqual(tasks, [['a', 'c'], ['b']]);
  assert.deepEqual(taskFileIds({ FILE_IDS: 'a,b,c', CLOUD_RUN_TASK_INDEX: '1', CLOUD_RUN_TASK_COUNT: '3' }), ['b']);
});

test('ftpName, progressText and stamp', () => {
  assert.equal(ftpName(' a/b\\c.mp4 '), 'a-b-c.mp4');
  assert.equal(progressText(1.2e9, 4.8e9), 'Skickar 1.2 GB av 4.8 GB (25 %)');
  assert.equal(progressText(300e6, 4.8e9), 'Skickar 300 MB av 4.8 GB (6 %)');
  assert.equal(progressText(300e6, NaN), 'Skickar 300 MB');
  assert.equal(stamp(new Date('2026-07-01T10:05:00Z'), 'Europe/Stockholm'), '2026-07-01 12:05');
});

// --- Google APIs ----------------------------------------------------------------------------

// A fetch answering from routes: [[regexp, { status, json, text, body }]]; requests are recorded.
function fakeFetch(routes) {
  const requests = [];
  const fn = async (url, init = {}) => {
    requests.push({ url: decodeURIComponent(url), init });
    const route = routes.find(([re]) => re.test(decodeURIComponent(url)));
    if (!route) throw new Error('Unexpected request ' + url);
    const r = route[1];
    return {
      ok: (r.status || 200) < 300,
      status: r.status || 200,
      json: async () => r.json,
      text: async () => r.text || JSON.stringify(r.json || ''),
      body: r.body,
    };
  };
  fn.requests = requests;
  return fn;
}

test('sheetStatus writes to the status column of the row holding the file ID', async () => {
  const fetch = fakeFetch([
    [/:batchGet/, { json: { valueRanges: [
      { values: [['Drive file ID', 'Name', 'Drive modified', 'Size', 'Elmo path', 'Duration', 'Skicka till DMC', 'DMC-status']] },
      { values: [['Drive file ID'], ['x'], ['f1']] },
    ] } }],
    [/values\/'Video Assets'!H3/, { json: {} }],
  ]);
  const status = googleApi({ fetch, token: async () => 'T' })
    .sheetStatus({ spreadsheetId: 'S', sheetName: 'Video Assets', header: 'DMC-status' });
  await status('f1', 'Skickar');
  const put = fetch.requests[1];
  assert.equal(put.url, "https://sheets.googleapis.com/v4/spreadsheets/S/values/'Video Assets'!H3?valueInputOption=RAW");
  assert.equal(put.init.method, 'PUT');
  assert.equal(put.init.headers.Authorization, 'Bearer T');
  assert.deepEqual(JSON.parse(put.init.body), { values: [['Skickar']] });
});

test('sheetStatus fails clearly without the column or the row', async () => {
  const api = (head, ids) => googleApi({ token: async () => 'T', fetch: fakeFetch([
    [/:batchGet/, { json: { valueRanges: [{ values: [head] }, { values: ids }] } }],
  ]) }).sheetStatus({ spreadsheetId: 'S', sheetName: 'Video Assets', header: 'DMC-status' });
  await assert.rejects(api(['Drive file ID'], [['f1']])('f1', 'x'), /No "DMC-status" column in Video Assets/);
  await assert.rejects(api(['Drive file ID', 'DMC-status'], [['Drive file ID'], ['other']])('f1', 'x'),
    /f1 is not in Video Assets/);
});

test('drive.metadata and drive.download ask Drive for the file, also in shared drives', async () => {
  const fetch = fakeFetch([
    [/alt=media/, { body: Readable.toWeb(Readable.from([Buffer.from('video')])) }],
    [/fields=name,size/, { json: { name: 'a.mp4', size: '5' } }],
  ]);
  const { drive } = googleApi({ fetch, token: async () => 'T' });
  assert.deepEqual(await drive.metadata('F'), { name: 'a.mp4', size: 5 });
  const chunks = [];
  for await (const c of await drive.download('F')) chunks.push(c);
  assert.equal(Buffer.concat(chunks).toString(), 'video');
  assert.ok(fetch.requests.every((r) => r.url.includes('supportsAllDrives=true')));
});

test('a failed Google request names the request and the answer', async () => {
  const { drive } = googleApi({ token: async () => 'T', fetch: fakeFetch([[/files/, { status: 404, text: 'File not found: F' }]]) });
  await assert.rejects(drive.metadata('F'), /GET https:\/\/www.googleapis.com\/drive\/v3\/files\/F failed \(404\): File not found: F/);
});

test('columnLetter', () => {
  assert.deepEqual([1, 8, 26, 27].map(columnLetter), ['A', 'H', 'Z', 'AA']);
});
