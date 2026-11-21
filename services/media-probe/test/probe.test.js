// Tests for the media probe job. The fixtures are real ffprobe output for clips made with ffmpeg in
// the formats the inbox gets (ProRes 422/4444 in MOV and MXF, H.264 and H.265 in MP4), and real
// idet output. The end-to-end tests run the real ffprobe against a local HTTP server standing in
// for Drive; they are skipped where ffprobe is not installed.
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const probe = require('../src/probe');
const { probeAll, taskFileIds, stamp } = require('../src/main');
const { googleApi } = require('../src/google');

const FIXTURES = path.join(__dirname, 'fixtures');
const fixture = (name) => JSON.parse(fs.readFileSync(path.join(FIXTURES, name + '.json'), 'utf8'));
const row = (name) => Object.values(probe.describe(fixture(name)).values);

// --- What each kind of file is described as ------------------------------------------------------

test('ProRes: the flavour, 10 or 12 bits, alpha, and the scan from the first frame', () => {
  assert.deepEqual(row('prores422hq_p25.mov'), [2 / 86400, 25, 1920, 1080, 'ProRes 422 HQ', 62.4, 'Progressive', 10, 'No', 'mono (1 ch)']);
  assert.deepEqual(row('prores422hq_bff.mov'), [2 / 86400, 25, 1920, 1080, 'ProRes 422 HQ', 73.1, 'Interlaced BFF', 10, 'No', 'none']);
  assert.deepEqual(row('prores4444_alpha.mxf'), [2 / 86400, 25, 1920, 1080, 'ProRes 4444', 123.2, 'Progressive', 12, 'Yes', 'none']);
});

test('interlaced ProRes in MXF is interlaced, although the container says progressive', () => {
  const json = fixture('prores422hq_i25.mxf');
  assert.equal(json.streams[0].field_order, 'progressive', 'what FFmpeg reports for the stream');
  assert.deepEqual(row('prores422hq_i25.mxf'), [2 / 86400, 25, 1920, 1080, 'ProRes 422 HQ', 73.5, 'Interlaced TFF', 10, 'No', '2 × mono (2 ch)']);
});

test('H.264: the profile, and the scan from the stream\'s field order', () => {
  assert.deepEqual(row('h264_i2997.mp4'), [2 / 86400, 29.97, 1920, 1080, 'H.264 High', 6.3, 'Interlaced TFF', 8, 'No', 'stereo (2 ch)']);
  assert.deepEqual(row('h264_p50.mp4'), [2 / 86400, 50, 1280, 720, 'H.264 High', 4.0, 'Progressive', 8, 'No', 'stereo (2 ch)']);
});

test('H.265 is always progressive', () => {
  assert.deepEqual(row('hevc_main10.mp4'), [1 / 86400, 25, 3840, 2160, 'H.265 Main 10', 15.0, 'Progressive', 10, 'No', 'none']);
  const flaggedInterlaced = fixture('hevc_main10.mp4');
  flaggedInterlaced.streams[0].field_order = 'tt';
  assert.equal(probe.describe(flaggedInterlaced).values.Scan, 'Progressive');
  assert.equal(probe.describe(flaggedInterlaced).needsMeasuring, false);
});

test('a file that does not say how it is scanned is measured; tb and bt are measured too', () => {
  ['unknown', undefined, 'tb', 'bt'].forEach((order) => {
    const json = fixture('h264_p50.mp4');
    json.streams[0].field_order = order;
    const d = probe.describe(json);
    assert.equal(d.needsMeasuring, true, String(order));
    assert.equal(d.values.Scan, '');
  });
});

test('a file without video is an error naming what ffprobe saw', () => {
  assert.throws(() => probe.describe({ streams: [{ codec_type: 'audio', channels: 2 }], format: { format_name: 'wav' } }),
    /no video stream \(ffprobe sees: wav\)/);
  const coverArt = { streams: [{ index: 0, codec_type: 'video', disposition: { attached_pic: 1 } }], format: {} };
  assert.throws(() => probe.describe(coverArt), /no video stream/, 'cover art is not the video');
});

test('duration: whole seconds as a fraction of a day, frames past the last second dropped', () => {
  const sec = (x) => Math.round(x * 86400 * 1e6) / 1e6;
  assert.equal(sec(probe.durationDays({ avg_frame_rate: '25/1', nb_frames: '90061' }, {})), 3602, '1:00:02 and 11 frames');
  assert.equal(sec(probe.durationDays({ avg_frame_rate: '24/1', nb_frames: '2160' }, {})), 90, 'exactly 90 s stays 90');
  assert.equal(sec(probe.durationDays({ avg_frame_rate: '30000/1001', nb_frames: '107892' }, {})), 3599,
    '29.97: real time (3599.6 s), not timecode (non-drop counts these frames as 0:59:56)');
  // MXF: no frame count, the duration on the container.
  assert.equal(sec(probe.durationDays({ avg_frame_rate: '0/0', r_frame_rate: '25/1' }, { duration: '65.04' })), 65);
  assert.equal(probe.durationDays({ avg_frame_rate: '25/1' }, {}), '', 'unknown');
});

test('bitrate: Mbit/s as a number, the video\'s own, else the whole file\'s minus the audio', () => {
  const v = { codec_type: 'video' };
  assert.equal(probe.bitrate([v], {}, { bit_rate: '73112468' }), 73.1);
  // MXF: 8 tracks of 24-bit 48 kHz mono, one reporting its bitrate and the rest only their format.
  const audio = [{ codec_type: 'audio', bit_rate: '1152000' }, ...Array(7).fill({ codec_type: 'audio', sample_rate: '48000', bits_per_sample: 24, channels: 1 })];
  assert.equal(probe.bitrate([v, ...audio], { bit_rate: String(120e6 + 8 * 1152000) }, v), 120.0);
  assert.equal(probe.bitrate([v], {}, v), '', 'unknown');
});

test('frame rates, codec names, bit depth, alpha and audio layouts', () => {
  assert.deepEqual(['30000/1001', '24000/1001', '25/1', '60000/1001', '0/0', ''].map(probe.frameRate),
    [29.97, 23.976, 25, 59.94, '', '']);
  assert.equal(probe.codecName({ codec_name: 'prores', profile: 'XQ' }), 'ProRes 4444 XQ');
  assert.equal(probe.codecName({ codec_name: 'prores', profile: 'Proxy' }), 'ProRes 422 Proxy');
  assert.equal(probe.codecName({ codec_name: 'prores', profile: 'Standard' }), 'ProRes 422');
  assert.equal(probe.codecName({ codec_name: 'dnxhd', codec_long_name: 'VC3/DNxHD', profile: 'DNXHR HQX' }), 'VC3/DNxHD DNXHR HQX');
  assert.equal(probe.bitDepth({ pix_fmt: 'yuv422p10le' }), 10);
  assert.equal(probe.bitDepth({ pix_fmt: 'yuv420p' }), 8);
  assert.equal(probe.hasAlpha('yuva444p12le'), true);
  assert.equal(probe.hasAlpha('yuv444p12le'), false);
  const audio = (...s) => probe.audioLayout(s.map((x) => Object.assign({ codec_type: 'audio' }, x)));
  assert.equal(audio(...Array(8).fill({ channels: 1, channel_layout: 'unknown' })), '8 × mono (8 ch)');
  assert.equal(audio({ channels: 6, channel_layout: '5.1(side)' }, { channels: 2, channel_layout: 'stereo' }), '5.1 + stereo (8 ch)');
  assert.equal(audio({ channels: 16 }), '16 ch (16 ch)');
});

test('idet: the report that counted frames wins over the empty one ffmpeg also prints', () => {
  const read = (f) => fs.readFileSync(path.join(FIXTURES, f), 'utf8');
  assert.equal(probe.measuredScan(read('idet-progressive.txt')), 'Progressive (measured)');
  assert.equal(probe.measuredScan(read('idet-tff.txt')), 'Interlaced TFF (measured)');
  assert.equal(probe.measuredScan('Multi frame detection: TFF: 2 BFF: 40 Progressive: 5 Undetermined: 3'), 'Interlaced BFF (measured)');
  assert.equal(probe.measuredScan('no idet here'), 'Unknown');
});

test('probeFile measures only when it has to, and never puts the token in an error', async () => {
  const calls = [];
  const needsIt = fixture('h264_p50.mp4');
  needsIt.streams[0].field_order = 'unknown';
  const run = async (cmd, args) => {
    calls.push(cmd);
    if (cmd === 'ffprobe') return { stdout: JSON.stringify(needsIt), stderr: '' };
    return { stdout: '', stderr: fs.readFileSync(path.join(FIXTURES, 'idet-tff.txt'), 'utf8') };
  };
  const { values } = await probe.probeFile({ url: 'https://x', token: 'SECRET', run });
  assert.deepEqual(calls, ['ffprobe', 'ffmpeg']);
  assert.equal(values.Scan, 'Interlaced TFF (measured)');

  const once = [];
  await probe.probeFile({ url: 'https://x', token: 'SECRET', run: async (cmd) => { once.push(cmd); return { stdout: JSON.stringify(fixture('h264_p50.mp4')) }; } });
  assert.deepEqual(once, ['ffprobe'], 'no measuring when the file says how it is scanned');

  // The real runCommand, with a command that fails: the error has the tool's message, not its arguments.
  await assert.rejects(probe.runCommand('node', ['-e', 'console.error("boom"); process.exit(3)', 'Authorization: Bearer SECRET']),
    (e) => /node failed: boom/.test(e.message) && !e.message.includes('SECRET'));
});

// --- Writing to the sheet ---------------------------------------------------------------------------

// A fake Sheets API holding one tab: headers and rows, as values.
function fakeSheets(headers, rows) {
  const data = [headers, ...rows.map((r) => r.slice())];
  const requests = [];
  const fetch = async (url, init = {}) => {
    requests.push({ url, init });
    const ok = (body) => ({ ok: true, json: async () => body, text: async () => JSON.stringify(body) });
    if (url.includes(':batchGet')) {
      return ok({ valueRanges: [{ values: [data[0]] }, { values: data.map((r) => [r[0]]) }] });
    }
    if (url.includes(':batchUpdate')) {
      JSON.parse(init.body).data.forEach(({ range, values }) => {
        const [, col, rowNo] = /!([A-Z]+)(\d+)$/.exec(range);
        const c = col.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
        data[Number(rowNo) - 1][c] = values[0][0];
      });
      return ok({});
    }
    throw new Error('unexpected ' + url);
  };
  const api = googleApi({ fetch, token: async () => 'TOKEN' });
  const stored = {};
  api.setFileProperties = async (id, props) => { stored[id] = props; };
  return { data, requests, api, stored };
}

const INBOX_HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size',
  ...probe.RESULT_HEADERS, 'Probe status'];

test('probeAll writes each video\'s values and "Probed <time>" to its row, in one write per step', async () => {
  const sheets = fakeSheets(INBOX_HEADERS, [['a', 'a.mov'], ['b', 'b.mp4']]);
  const run = async () => ({ stdout: JSON.stringify(fixture('prores422hq_i25.mxf')) });
  const failed = await probeAll(['b'], {
    api: sheets.api, sheet: { spreadsheetId: 'S', sheetName: 'Video Inbox' }, run,
    now: () => new Date('2026-11-21T16:02:00Z'),
  });
  assert.equal(failed, 0);
  assert.deepEqual(sheets.data[2].slice(4), [2 / 86400, 25, 1920, 1080, 'ProRes 422 HQ', 73.5, 'Interlaced TFF', 10, 'No', '2 × mono (2 ch)', 'Probed 2026-11-21 17:02']);
  assert.equal(sheets.data[1].length, 2, 'the other row is untouched');
  const writes = sheets.requests.filter((r) => r.url.includes(':batchUpdate'));
  assert.equal(writes.length, 2, '"Probing", then the results');
  assert.match(sheets.requests[0].url, /ranges='Video%20Inbox'!1%3A1&ranges='Video%20Inbox'!A%3AA$/);
});

test('probeAll writes the error for a video it cannot probe and goes on with the next', async () => {
  const sheets = fakeSheets(INBOX_HEADERS, [['a', 'a.txt'], ['b', 'b.mp4']]);
  const run = async (cmd, args) => (args[args.length - 1].includes('/files/a?')
    ? Promise.reject(new Error('ffprobe failed: Invalid data found when processing input'))
    : { stdout: JSON.stringify(fixture('h264_p50.mp4')) });
  const failed = await probeAll(['a', 'b'], { api: sheets.api, sheet: { spreadsheetId: 'S', sheetName: 'Video Inbox' }, run });
  assert.equal(failed, 1);
  assert.equal(sheets.data[1][14], 'Error: ffprobe failed: Invalid data found when processing input');
  assert.equal(sheets.data[1][8], undefined, 'no values for the failed one');
  assert.match(sheets.data[2][14], /^Probed /);
});

test('probeAll fails clearly when a column is missing from the tab', async () => {
  const sheets = fakeSheets(['Drive file ID', 'Name'], [['a', 'a.mp4']]);
  const failed = await probeAll(['a'], { api: sheets.api, sheet: { spreadsheetId: 'S', sheetName: 'Video Inbox' }, run: async () => ({}) });
  assert.equal(failed, 1);
});

test('taskFileIds shares the IDs between tasks, and stamp writes Stockholm time', () => {
  assert.deepEqual(taskFileIds({ FILE_IDS: 'a, b,c,,d', CLOUD_RUN_TASK_INDEX: '1', CLOUD_RUN_TASK_COUNT: '2' }), ['b', 'd']);
  assert.deepEqual(taskFileIds({ FILE_IDS: 'a,b' }), ['a', 'b']);
  assert.equal(stamp(new Date('2026-07-01T10:00:00Z'), 'Europe/Stockholm'), '2026-07-01 12:00');
});

// --- End to end, with the real ffprobe ------------------------------------------------------------

const hasFfprobe = (() => { try { execFileSync('ffprobe', ['-version']); return true; } catch (e) { return false; } })();

// Serves the fixtures like Drive's download URL: only with the token, and with range requests.
function fakeDrive() {
  const seen = { ranges: 0, unauthorized: 0 };
  const server = http.createServer((req, res) => {
    if (req.headers.authorization !== 'Bearer TOKEN') { seen.unauthorized++; res.writeHead(401).end(); return; }
    const file = path.join(FIXTURES, decodeURIComponent(new URL(req.url, 'http://x').pathname.split('/').pop()));
    const body = fs.readFileSync(file);
    const range = /^bytes=(\d+)-(\d*)$/.exec(req.headers.range || '');
    if (range) {
      seen.ranges++;
      const start = Number(range[1]);
      const end = range[2] ? Math.min(Number(range[2]), body.length - 1) : body.length - 1;
      res.writeHead(206, { 'Content-Range': `bytes ${start}-${end}/${body.length}`, 'Content-Length': end - start + 1, 'Accept-Ranges': 'bytes' });
      res.end(body.subarray(start, end + 1));
    } else {
      res.writeHead(200, { 'Content-Length': body.length, 'Accept-Ranges': 'bytes' });
      res.end(body);
    }
  });
  return new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve({ server, seen, url: (f) => `http://127.0.0.1:${server.address().port}/files/${f}` })));
}

test('end to end: ffprobe reads interlaced ProRes in MXF from the server, with the token', { skip: !hasFfprobe && 'ffprobe is not installed' }, async () => {
  const drive = await fakeDrive();
  try {
    const { values } = await probe.probeFile({ url: drive.url('prores-tff.mxf'), token: 'TOKEN' });
    const v = Object.values(values);
    assert.equal(typeof v[5], 'number', 'MXF: estimated, still a number');
    assert.ok(v[5] > 0);
    assert.deepEqual(v.slice(0, 5).concat(v.slice(6)), [0, 25, 320, 180, 'ProRes 422 HQ', 'Interlaced TFF', 10, 'No', '2 × mono (2 ch)']);
    assert.equal(drive.seen.unauthorized, 0);
    await assert.rejects(probe.probeFile({ url: drive.url('prores-tff.mxf'), token: 'WRONG' }), (e) => !e.message.includes('WRONG'));
  } finally {
    drive.server.close();
  }
});

test('end to end: an MP4 with its index at the end is read with range requests', { skip: !hasFfprobe && 'ffprobe is not installed' }, async () => {
  const drive = await fakeDrive();
  try {
    const { values } = await probe.probeFile({ url: drive.url('h264-mp4-moov-last.mp4'), token: 'TOKEN' });
    assert.deepEqual(Object.values(values).filter((x, i) => i !== 5).slice(0, 6), [0, 25, 320, 180, 'H.264 High', 'Progressive']);
    assert.ok(drive.seen.ranges > 0, 'ffprobe asked for parts of the file');
  } finally {
    drive.server.close();
  }
});

// --- Results stored on the Drive file (TARGET=drive) -------------------------------------------------

const { toProperties, PROPERTY_KEYS } = require('../src/properties');

test('toProperties: every value as text, duration in whole seconds, when it was probed; empty values left out', () => {
  const { values } = probe.describe(fixture('prores422hq_i25.mxf'));
  values.Audio = '';
  assert.deepEqual(toProperties(values, new Date('2026-11-21T20:00:00Z')), {
    'sesport.duration': '2', 'sesport.frameRate': '25', 'sesport.width': '1920', 'sesport.height': '1080',
    'sesport.codec': 'ProRes 422 HQ', 'sesport.mbit': '73.5', 'sesport.scan': 'Interlaced TFF',
    'sesport.bitDepth': '10', 'sesport.alpha': 'No', 'sesport.probed': '2026-11-21T20:00:00.000Z',
  });
  assert.deepEqual(Object.keys(PROPERTY_KEYS), probe.RESULT_HEADERS, 'a key for every column the job fills');
  const long = toProperties({ Codec: 'x'.repeat(300) }, new Date(0))['sesport.codec'];
  assert.equal(long.length, 100, 'within Drive\'s 124 bytes for key and value');
});

test('probeToDrive stores each video\'s results on the file, and goes on after a failure', async () => {
  const requests = [];
  const fetch = async (url, init = {}) => {
    requests.push({ url, init });
    return { ok: true, json: async () => ({}), text: async () => '' };
  };
  const api = googleApi({ fetch, token: async () => 'TOKEN' });
  const run = async (cmd, args) => (args[args.length - 1].includes('/files/bad?')
    ? Promise.reject(new Error('ffprobe failed: 403 Forbidden'))
    : { stdout: JSON.stringify(fixture('h264_p50.mp4')) });
  const failed = await require('../src/main').probeToDrive(['a', 'bad'], { api, run, now: () => new Date('2026-11-21T20:00:00Z') });
  assert.equal(failed, 1);
  assert.equal(requests.length, 1, 'one write, for the one that could be probed');
  assert.equal(requests[0].url, 'https://www.googleapis.com/drive/v3/files/a?supportsAllDrives=true&fields=id');
  assert.equal(requests[0].init.method, 'PATCH');
  assert.equal(JSON.parse(requests[0].init.body).properties['sesport.codec'], 'H.264 High');
});

test('probing for the sheet also stores the results on the file; a failure there is logged, not an error', async () => {
  const sheets = fakeSheets(INBOX_HEADERS, [['a', 'a.mov'], ['b', 'b.mov']]);
  const run = async () => ({ stdout: JSON.stringify(fixture('prores422hq_i25.mxf')) });
  const sheet = { spreadsheetId: 'S', sheetName: 'Video Inbox' };
  await probeAll(['a'], { api: sheets.api, sheet, run, now: () => new Date('2026-11-21T20:00:00Z') });
  assert.equal(sheets.stored.a['sesport.codec'], 'ProRes 422 HQ');
  assert.equal(sheets.stored.a['sesport.probed'], '2026-11-21T20:00:00.000Z');

  sheets.api.setFileProperties = async () => { throw new Error('403 insufficientFilePermissions'); };
  const failed = await probeAll(['b'], { api: sheets.api, sheet, run });
  assert.equal(failed, 0, 'still probed');
  assert.match(sheets.data[2][14], /^Probed /);
});
