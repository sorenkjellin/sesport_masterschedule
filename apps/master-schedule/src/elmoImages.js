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

// --- Step 1: Drive folder -> sheet ------------------------------------------

function listDriveImages() {
  const cfg = config_();
  const sheet = getSheet_();

  const existing = {};
  readTextRows_(sheet, HEADERS.length).forEach(function (r) { if (r[COL.id]) existing[r[COL.id]] = r; });

  const found = listDriveFiles_(DriveApp.getFolderById(cfg.folderId), 'image/');
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
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { toast_('Another upload is already running', 'Elmo'); return; }
  const started = Date.now();

  try {
    const cfg = config_();
    const sheet = getSheet_();
    const rows = readTextRows_(sheet, HEADERS.length);
    const elmo = getElmoFiles_(cfg);   // path -> { hash, url }
    ensureElmoProject_(cfg);

    const pathCount = {};
    rows.forEach(function (r) {
      if (r[COL.id] && r[COL.status] !== MISSING) pathCount[r[COL.path]] = (pathCount[r[COL.path]] || 0) + 1;
    });

    const stats = { uploaded: 0, unchanged: 0, failed: 0, remaining: 0 };

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r[COL.id] || r[COL.status] === MISSING) continue;
      if (Date.now() - started > MAX_RUNTIME_MS) { stats.remaining++; continue; }

      const path = r[COL.path];
      const before = r.join('\u0000');
      try {
        if (pathCount[path] > 1) {
          r[COL.status] = 'Duplicate Elmo path, rename in Drive';
          stats.failed++;
        } else {
          const file = DriveApp.getFileById(r[COL.id]);
          const modified = file.getLastUpdated().toISOString();
          const inElmo = elmo[path];

          if (r[COL.hash] && r[COL.upMod] === modified && inElmo && inElmo.hash === r[COL.hash]) {
            // Fast path: untouched in Drive since our last upload, and Elmo
            // still has that version. No download needed.
            r[COL.url] = inElmo.url;
            r[COL.status] = 'Up to date';
            stats.unchanged++;
          } else {
            const blob = file.getBlob();
            const hash = sha256Hex_(blob.getBytes());
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
          }
          r[COL.modified] = modified;
          r[COL.upMod] = modified;
        }
      } catch (e) {
        r[COL.status] = 'Error: ' + e.message;
        stats.failed++;
      }

      // Write each changed row as we go so progress survives a timeout.
      if (r.join('\u0000') !== before) writeTextRow_(sheet, i, r);
    }

    markMissingGenreImages();
    toast_('Uploaded ' + stats.uploaded + ', unchanged ' + stats.unchanged +
           ', failed ' + stats.failed +
           (stats.remaining ? ', ' + stats.remaining + ' left for next run' : ''), 'Elmo');
  } finally {
    lock.releaseLock();
  }
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
