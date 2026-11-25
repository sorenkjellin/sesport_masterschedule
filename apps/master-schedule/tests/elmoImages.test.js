// Unit tests for listing images and uploading them to Elmo (elmoImages.js). Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { load, scriptsIn, plain, fakeBlob, FakeSheet, fakeSpreadsheet } = require('../../../tests/harness');

const FILES = ['lib/lib.js', ...scriptsIn('apps/master-schedule/src')];
const HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Elmo path', 'Elmo URL',
  'Uploaded hash', 'Uploaded modified', 'Status'];
const FOLDER = 'application/vnd.google-apps.folder';
const T = '2026-11-01T10:00:00.000Z';
const sha = (s) => crypto.createHash('sha256').update(s).digest('hex');
const img = (id, parent, content, o) => ({ id, name: id + '.png', mimeType: 'image/png', parents: [parent],
  modifiedTime: T, size: '3', sha256Checksum: sha(content), content, ...o });
const row = (id, { hash = '', mod = '', path = `images/${id}.png`, status = 'New' } = {}) =>
  [id, id + '.png', '', '3', path, '', hash, mod, status];

// drive: files and folders. elmo: path -> content.
function images({ drive = [], rows = [], elmo = {}, pageSize = 1000, incomplete = false, failUpload } = {}) {
  const sheet = new FakeSheet('Elmo images', [HEADERS, ...rows]);
  const state = { lists: 0, downloads: [], uploads: [], cache: {}, writes: 0, scriptLocks: 0, toasts: [] };
  const setValues = sheet.getRange(1, 1).constructor.prototype.setValues;
  const getRange = sheet.getRange.bind(sheet);
  sheet.getRange = (r, ...a) => {
    const range = getRange(r, ...a);
    if (r > 1) range.setValues = (v) => { state.writes++; return setValues.call(range, v); };
    return range;
  };
  const ss = fakeSpreadsheet([sheet, new FakeSheet('Kalender', [['Datum', 'Börjar', 'Warning']])]);
  ss.toast = (m) => state.toasts.push(m);
  const gas = load(FILES, {
    spreadsheet: ss,
    properties: { DRIVE_FOLDER_ID: 'ROOT', ELMO_URL: 'https://elmo', ELMO_KEY: 'k', ELMO_PROJECT: 'p' },
    fetch: (url, o) => {
      if (url.startsWith('https://www.googleapis.com/drive')) {
        state.lists++;
        const q = new URL(url).searchParams;
        const parents = [...q.get('q').matchAll(/'([^']+)' in parents/g)].map((m) => m[1]);
        const all = drive.filter((f) => parents.includes(f.parents[0]));
        const from = Number(q.get('pageToken') || 0);
        return { code: 200, body: { files: all.slice(from, from + pageSize), incompleteSearch: incomplete,
          nextPageToken: from + pageSize < all.length ? String(from + pageSize) : undefined } };
      }
      if (url.includes('?recurse=1')) {
        return { code: 200, body: { list: Object.entries(elmo).map(([name, c]) => ({ name, url: 'https://cdn/' + sha(c) })) } };
      }
      if (url.endsWith('/upload-session')) {
        const path = JSON.parse(o.payload).path;
        if (failUpload === path) return { code: 500, body: 'no' };
        state.uploads.push(path);
        return { code: 200, body: { id: 's', url: 'https://s3/x' } };
      }
      return { code: 200, body: { url: 'https://cdn/new' } };
    },
    globals: {
      CacheService: { getScriptCache: () => ({ get: (k) => state.cache[k] || null,
        put: (k, v) => { state.cache[k] = v; }, remove: (k) => { delete state.cache[k]; } }) },
      LockService: {
        getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }),
        getScriptLock: () => { state.scriptLocks++; return { tryLock: () => true, releaseLock() {} }; },
      },
      DriveApp: { getFileById: (id) => {
        const f = drive.find((x) => x.id === id);
        if (!f) throw new Error('No such file');
        return { getLastUpdated: () => new Date(f.modifiedTime),
          getBlob: () => { state.downloads.push(id); return fakeBlob(f.content, 'image/png'); } };
      } },
    },
  });
  return { gas, sheet, state, status: () => sheet.data.slice(1).map((r) => r[8].split(' 20')[0]) };
}

test('images are listed with one request per level of folders, over pages', () => {
  const p = images({ pageSize: 2, drive: [
    img('b', 'ROOT', 'b'), { id: 'F', name: 'A/B', mimeType: FOLDER, parents: ['ROOT'] }, img('a', 'ROOT', 'a'),
    { id: 'G', name: 'Genre', mimeType: FOLDER, parents: ['ROOT'] }, img('c', 'F', 'c'), img('d', 'G', 'd'),
    img('e', 'G', 'e', { mimeType: 'application/vnd.google-apps.shortcut' })] });
  p.gas.listDriveImages();
  assert.deepEqual(p.sheet.data.slice(1).map((r) => r[4]),
    ['images/A-B/c.png', 'images/a.png', 'images/b.png', 'images/Genre/d.png']);
  assert.deepEqual(plain(p.sheet.data[1]), ['c', 'c.png', T, '3', 'images/A-B/c.png', '', '', '', 'New']);
  assert.equal(p.state.lists, 4);
  assert.deepEqual(p.state.cache, {});
});

test('an incomplete listing changes nothing', () => {
  const p = images({ rows: [row('a')], incomplete: true });
  assert.throws(() => p.gas.listDriveImages(), /did not list all/);
  assert.deepEqual([p.status(), p.state.cache], [['New'], {}]);
});

test('nothing new: no downloads or uploads, and the rows are written together, without the script lock', () => {
  const ids = ['a', 'b', 'c'];
  // Drive's modified time differs from the last upload's; the content does not.
  const p = images({ drive: ids.map((i) => img(i, 'ROOT', i)), rows: ids.map((i) => row(i, { hash: 'old', mod: 'x' })),
    elmo: { 'images/a.png': 'a', 'images/b.png': 'b', 'images/c.png': 'c' } });
  p.gas.uploadChangedImages();
  assert.deepEqual(p.status(), ['Up to date', 'Up to date', 'Up to date']);
  assert.deepEqual(p.sheet.data[1].slice(5, 8), ['https://cdn/' + sha('a'), sha('a'), T]);
  assert.deepEqual([p.state.downloads, p.state.uploads, p.state.writes, p.state.scriptLocks, p.state.lists], [[], [], 1, 0, 1]);
  p.gas.uploadChangedImages();
  assert.equal(p.state.writes, 1, 'nothing to write');
  assert.match(p.gas.fakes.logs.join('\n'), /Images: 3 rows, 0 downloaded; sheet .* s, Elmo/);
});

test('new and changed images are uploaded; failures and duplicates only affect themselves', () => {
  const p = images({
    drive: ['new', 'changed', 'same', 'bad', 'dup1', 'dup2'].map((i) => img(i, 'ROOT', i)),
    rows: [row('new'), row('changed', { hash: sha('before'), mod: T }), row('same'), row('bad'),
      row('dup1', { path: 'images/d.png' }), row('dup2', { path: 'images/d.png' }), row('gone'),
      row('m', { status: 'Missing in Drive' })],
    elmo: { 'images/changed.png': 'before', 'images/same.png': 'same' }, failUpload: 'images/bad.png' });
  p.gas.uploadChangedImages();
  assert.deepEqual(p.status(), ['Uploaded', 'Uploaded', 'Up to date', 'Error: open upload session failed (500): no',
    'Duplicate Elmo path, rename in Drive', 'Duplicate Elmo path, rename in Drive', 'Error: No such file', 'Missing in Drive']);
  assert.deepEqual(p.state.uploads, ['images/new.png', 'images/changed.png']);
  assert.deepEqual(p.state.downloads, ['new', 'changed', 'bad']);
  assert.deepEqual(p.state.toasts, ['Uploaded 2, unchanged 1, failed 4']);
});

test('without a checksum in Drive, an untouched image is known by its time, and others are downloaded', () => {
  const o = { sha256Checksum: undefined };
  const p = images({ drive: [img('a', 'ROOT', 'a', o), img('b', 'ROOT', 'b', o), img('moved', 'ELSEWHERE', 'moved')],
    rows: [row('a', { hash: sha('a'), mod: T }), row('b'), row('moved', { hash: sha('moved'), mod: T })],
    elmo: { 'images/a.png': 'a', 'images/b.png': 'b', 'images/moved.png': 'moved' } });
  p.gas.uploadChangedImages();
  assert.deepEqual(p.status(), ['Up to date', 'Up to date', 'Up to date']);
  assert.deepEqual(p.state.downloads, ['b']);
});

test('only one listing or upload at a time, and a failed one lets the next in', () => {
  const p = images({ rows: [row('a')] });
  p.state.cache.ELMO_IMAGES_BUSY = 'yes';
  p.gas.uploadChangedImages();
  p.gas.listDriveImages();
  assert.equal(p.gas.fakes.urlFetch.requests.length, 0);
  assert.match(p.state.toasts[0], /already being listed or uploaded/);
  assert.equal(p.state.cache.ELMO_IMAGES_BUSY, 'yes', 'the other run still has it');
  delete p.state.cache.ELMO_IMAGES_BUSY;
  p.gas.getElmoFiles_ = () => { throw new Error('Elmo is down'); };
  assert.throws(() => p.gas.uploadChangedImages(), /Elmo is down/);
  assert.deepEqual(p.state.cache, {});
});
