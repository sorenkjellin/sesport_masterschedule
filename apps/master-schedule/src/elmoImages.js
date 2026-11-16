// Drive -> Elmo: lists the Drive folder's images ("Elmo images") and videos ("Video Assets"),
// and uploads new and changed images to the Elmo project.

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

// --- Videos: Drive folder -> "Video Assets" columns A-D ----------------------
// The "Video" folder is the sibling of the images folder (DRIVE_FOLDER_ID). Lists every video in
// it and its subfolders with the same details as "Elmo images" columns A-D.
// Like "Elmo images", the list is rebuilt on every run, sorted by folder path and name.
// Only columns A-F are cleared and written.
const VIDEO_SHEET_NAME = 'Video Assets';
const VIDEO_FOLDER_NAME = 'Video';
// Drive file ID, Name, Drive modified, Size, Elmo path, Duration
const VIDEO_HEADERS = HEADERS.slice(0, 5).concat(['Duration']);

function listDriveVideos() {
  const cfg = config_();
  const imagesFolder = DriveApp.getFolderById(cfg.folderId);
  const videoFolder = findSiblingFolder_(imagesFolder, VIDEO_FOLDER_NAME);
  if (!videoFolder) {
    throw new Error('Hittade ingen mapp "' + VIDEO_FOLDER_NAME + '" bredvid "' + imagesFolder.getName() + '"');
  }

  const sheet = getOrCreateSheet_(VIDEO_SHEET_NAME, VIDEO_HEADERS);
  const found = listDriveFiles_(videoFolder, 'video/');
  found.sort(function (a, b) { return a.relPath.localeCompare(b.relPath); });
  // The folder inside the Elmo project: ELMO_VIDEO_FOLDER, default "videos".
  const elmoFolder = trimSlashes_(optionalProperty_('ELMO_VIDEO_FOLDER', 'videos'));
  // Duration as H:MM:SS:FF (25 fps); '' for a video Drive has not processed or cannot read.
  const durations = driveVideoDurationsMs_(found.map(function (f) { return f.id; }));
  const rows = found.map(function (f) {
    const ms = durations[f.id];
    return [f.id, f.name, f.modified, String(f.size), elmoFolder + '/' + f.relPath,
            ms == null ? '' : formatTimecode_(ms, 25)];
  });

  // Plain text, as in "Elmo images", so the timestamps stay as written.
  writeTextRows_(sheet, rows, VIDEO_HEADERS.length);
  toast_('Listed ' + found.length + ' videos from Drive', 'Elmo');
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
