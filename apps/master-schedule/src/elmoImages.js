// Drive -> Elmo: lists the Drive folder's images ("Elmo images") and uploads new and changed
// images to the Elmo project. (Videos are listed in the separate "Video Assets" spreadsheet: see
// apps/video-assets in the repo.)

const SHEET_NAME = 'Elmo images';
const HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Elmo path',
                 'Elmo URL', 'Uploaded hash', 'Uploaded modified', 'Status'];
const COL = { id: 0, name: 1, modified: 2, size: 3, path: 4, url: 5, hash: 6, upMod: 7, status: 8 };
const MISSING = 'Missing in Drive';
// Apps Script stops a run at 6 minutes; leave room to write the last row.
// Rows not reached are picked up by the next run.
const MAX_RUNTIME_MS = 5 * 60 * 1000;
// While uploading, finished rows are written at least this often, so that progress survives a timeout.
const IMAGE_ROWS_WRITE_MS = 10 * 1000;
// Set in the script cache while images are listed or uploaded. Not the script lock, which would stop
// checkProgramLengths and the körschema work for as long. It expires by itself (seconds, longer than
// a run can be) if the run is stopped.
const IMAGES_BUSY_KEY = 'ELMO_IMAGES_BUSY';
const IMAGES_BUSY_SECONDS = 400;
const DRIVE_FOLDERS_PER_REQUEST = 20;

// --- Step 1: Drive folder -> sheet ------------------------------------------

function listDriveImages() {
  withImagesBusy_(listDriveImagesNow_);
}

function listDriveImagesNow_() {
  const cfg = config_();
  const sheet = getSheet_();

  const existing = {};
  readTextRows_(sheet, HEADERS.length).forEach(function (r) { if (r[COL.id]) existing[r[COL.id]] = r; });

  const found = driveImages_(cfg.folderId);
  found.sort(function (a, b) { return a.relPath.localeCompare(b.relPath); });

  const seen = {};
  const rows = found.map(function (f) {
    seen[f.id] = true;
    const old = existing[f.id];
    return [
      f.id, f.name, f.modified, String(f.size), cfg.folder + '/' + f.relPath,
      old ? old[COL.url] : '',
      old ? old[COL.hash] : '',
      old ? old[COL.upMod] : '',
      old && old[COL.status] !== MISSING ? old[COL.status] : 'New',
    ];
  });

  // Keep rows whose Drive file is gone, flagged, so their Elmo URLs stay visible.
  Object.keys(existing).forEach(function (id) {
    if (seen[id]) return;
    const r = existing[id].slice();
    r[COL.status] = MISSING;
    rows.push(r);
  });

  writeTextRows_(sheet, rows, HEADERS.length);
  toast_('Listed ' + found.length + ' images from Drive', 'Elmo');
}

// --- Step 2: sheet -> Elmo, only what changed ------------------------------
function uploadChangedImages() {
  withImagesBusy_(uploadChangedImagesNow_);
}

function uploadChangedImagesNow_() {
  const started = Date.now();
  const times = [];
  let last = started;
  const lap = function (what) {
    times.push(what + ' ' + ((Date.now() - last) / 1000).toFixed(1) + ' s');
    last = Date.now();
  };

  const cfg = config_();
  const sheet = getSheet_();
  const rows = readTextRows_(sheet, HEADERS.length);
  lap('sheet');
  const elmo = getElmoFiles_(cfg);   // path -> { hash, url }
  ensureElmoProject_(cfg);
  lap('Elmo');
  const inDrive = {};
  driveImages_(cfg.folderId).forEach(function (f) { inDrive[f.id] = f; });
  lap('Drive');

  const pathCount = {};
  rows.forEach(function (r) {
    if (r[COL.id] && r[COL.status] !== MISSING) pathCount[r[COL.path]] = (pathCount[r[COL.path]] || 0) + 1;
  });

  const stats = { uploaded: 0, unchanged: 0, failed: 0, remaining: 0, downloaded: 0 };
  let changed = []; // indexes of rows not yet written
  let written = Date.now();
  const write = function () {
    writeChangedRows_(sheet, rows, changed);
    changed = [];
    written = Date.now();
  };

  try {
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r[COL.id] || r[COL.status] === MISSING) continue;
      if (Date.now() - started > MAX_RUNTIME_MS) { stats.remaining++; continue; }

      const before = r.join('\u0000');
      try {
        if (pathCount[r[COL.path]] > 1) {
          r[COL.status] = 'Duplicate Elmo path, rename in Drive';
          stats.failed++;
        } else {
          uploadImageRow_(cfg, r, inDrive[r[COL.id]], elmo, stats);
        }
      } catch (e) {
        r[COL.status] = 'Error: ' + e.message;
        stats.failed++;
      }
      if (r.join('\u0000') !== before) changed.push(i);
      if (changed.length && Date.now() - written > IMAGE_ROWS_WRITE_MS) write();
    }
  } finally {
    write();
  }
  lap('compare and upload');

  markMissingGenreImages();
  lap('warnings');
  // Rows without a file ID are not counted: a formula beside the list makes the sheet longer.
  const listed = rows.filter(function (r) { return r[COL.id]; });
  const missing = listed.filter(function (r) { return r[COL.status] === MISSING; }).length;
  Logger.log('Images: ' + listed.length + ' in the list (' + missing + ' missing in Drive); ' +
             stats.downloaded + ' downloaded; ' + times.join(', '));
  toast_('Uploaded ' + stats.uploaded + ', unchanged ' + stats.unchanged +
         ', failed ' + stats.failed +
         (stats.remaining ? ', ' + stats.remaining + ' left for next run' : ''), 'Elmo');
}

/**
 * Uploads a row's image if Elmo does not have it, and updates the row. listed: the image as from
 * driveImages_, absent when it is no longer in the folder.
 */
function uploadImageRow_(cfg, r, listed, elmo, stats) {
  const path = r[COL.path];
  let file = null;
  const driveFile = function () { return file || (file = DriveApp.getFileById(r[COL.id])); };
  const modified = listed ? listed.modified : driveFile().getLastUpdated().toISOString();
  const inElmo = elmo[path];

  // What Drive holds, without downloading it: the hash of the last upload if the file is untouched
  // since; for a file never handled, Drive's checksum. Not that after a change: for a new version of a
  // file, Drive can still give the old version's checksum.
  let hash = '';
  if (r[COL.upMod]) {
    if (r[COL.upMod] === modified) hash = r[COL.hash];
  } else if (listed) {
    hash = listed.sha256;
  }
  let blob = null;
  if (!hash || !inElmo || inElmo.hash !== hash) {
    blob = driveFile().getBlob();
    stats.downloaded++;
    hash = sha256Hex_(blob.getBytes());
  }
  if (inElmo && inElmo.hash === hash) {
    r[COL.url] = inElmo.url;
    r[COL.status] = 'Up to date';
    stats.unchanged++;
  } else {
    r[COL.url] = uploadToElmo_(cfg, blob, path);
    r[COL.status] = 'Uploaded ' + new Date().toISOString();
    elmo[path] = { hash: hash, url: r[COL.url] };
    stats.uploaded++;
  }
  r[COL.hash] = hash;
  r[COL.modified] = modified;
  r[COL.upMod] = modified;
}

/** Writes the rows at indexes (ascending), neighbours together. */
function writeChangedRows_(sheet, rows, indexes) {
  for (let i = 0; i < indexes.length;) {
    let n = 1;
    while (indexes[i + n] === indexes[i] + n) n++;
    sheet.getRange(indexes[i] + 2, 1, n, HEADERS.length).setNumberFormat('@')
      .setValues(rows.slice(indexes[i], indexes[i] + n));
    i += n;
  }
}

/** Runs fn unless images are being listed or uploaded already (by anyone). */
function withImagesBusy_(fn) {
  const cache = CacheService.getScriptCache();
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  let busy;
  try {
    busy = cache.get(IMAGES_BUSY_KEY);
    if (!busy) cache.put(IMAGES_BUSY_KEY, 'yes', IMAGES_BUSY_SECONDS);
  } finally {
    lock.releaseLock();
  }
  if (busy) {
    toast_('Images are already being listed or uploaded. If a run was stopped, wait a few minutes.', 'Elmo', 8);
    return;
  }
  try {
    fn();
  } finally {
    cache.remove(IMAGES_BUSY_KEY);
  }
}

// --- Drive ------------------------------------------------------------------

/**
 * Every image in the folder and its subfolders, as listDriveFiles_ (lib.js) gives them, and with
 * sha256: Drive's checksum ('' when it has none; possibly the previous version's after a new one). One request per level of folders (and 1000 items), where
 * listDriveFiles_ asks about each file.
 */
function driveImages_(folderId) {
  const out = [];
  const seen = new Set([folderId]);
  let level = [{ id: folderId, prefix: '' }];
  while (level.length) {
    const next = [];
    for (let i = 0; i < level.length; i += DRIVE_FOLDERS_PER_REQUEST) {
      const prefixes = new Map(level.slice(i, i + DRIVE_FOLDERS_PER_REQUEST)
        .map(function (f) { return [f.id, f.prefix]; }));
      driveChildren_(Array.from(prefixes.keys())).forEach(function (f) {
        const parent = (f.parents || []).find(function (id) { return prefixes.has(id); });
        if (parent === undefined || seen.has(f.id)) return;
        seen.add(f.id);
        const relPath = prefixes.get(parent) + f.name.replace(/\//g, '-');
        if (f.mimeType === 'application/vnd.google-apps.folder') {
          next.push({ id: f.id, prefix: relPath + '/' });
        } else if (f.mimeType.indexOf('image/') === 0) {
          out.push({ id: f.id, name: f.name, relPath: relPath, modified: new Date(f.modifiedTime).toISOString(),
                     size: Number(f.size || 0), mimeType: f.mimeType, sha256: f.sha256Checksum || '' });
        }
      });
    }
    level = next;
  }
  return out;
}

/** The folders and images directly in the folders (IDs), as Drive's API gives them. */
function driveChildren_(folderIds) {
  const items = [];
  let pageToken = '';
  do {
    const query = {
      q: '(' + folderIds.map(function (id) { return "'" + id + "' in parents"; }).join(' or ') + ') and trashed = false' +
        " and (mimeType = 'application/vnd.google-apps.folder' or mimeType contains 'image/')",
      corpora: 'allDrives', includeItemsFromAllDrives: 'true', supportsAllDrives: 'true', pageSize: '1000',
      fields: 'nextPageToken,incompleteSearch,files(id,name,mimeType,parents,modifiedTime,size,sha256Checksum)',
    };
    if (pageToken) query.pageToken = pageToken;
    const url = 'https://www.googleapis.com/drive/v3/files?' + Object.keys(query).map(function (k) {
      return k + '=' + encodeURIComponent(query[k]);
    }).join('&');
    const page = googleApiRequest_(url, { method: 'get' }, 'Listing images');
    // Files left out would be flagged as missing.
    if (page.incompleteSearch) throw new Error('Drive did not list all images; try again');
    items.push.apply(items, page.files || []);
    pageToken = page.nextPageToken;
  } while (pageToken);
  return items;
}

// --- Config -----------------------------------------------------------------

/** The Elmo settings from Script Properties; also the `elmo` argument of the Elmo API in lib.js. */
function config_() {
  return {
    folderId: requiredProperty_('DRIVE_FOLDER_ID'),
    elmoUrl: requiredProperty_('ELMO_URL').replace(/\/+$/, ''),
    key: requiredProperty_('ELMO_KEY'),
    project: requiredProperty_('ELMO_PROJECT'),
    folder: trimSlashes_(optionalProperty_('ELMO_FOLDER', 'images')),
  };
}

function getSheet_() {
  return getOrCreateSheet_(SHEET_NAME, HEADERS);
}
