// Drive "Video" folder -> the "Video Assets" sheet.
// Lists every video in the folder (VIDEO_FOLDER_ID) and its subfolders, sorted by folder path and
// name. The list is rebuilt on every run: columns A-F come from Drive, and what is in G-H (sending
// to DMC, see dmcTransfer.js) moves along with its video, matched by Drive file ID. A video gone
// from Drive loses its row.

const VIDEO_SHEET_NAME = 'Video Assets';
// Drive file ID, Name, Drive modified, Size, Path, Duration (A-F, from Drive),
// Skicka till DMC, DMC-status (G-H, kept). The DMC job finds its rows by column A and its status
// column by the header "DMC-status".
const VIDEO_DRIVE_HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Path', 'Duration'];
const VIDEO_HEADERS = VIDEO_DRIVE_HEADERS.concat(['Skicka till DMC', 'DMC-status']);
// What a new video gets in G-H.
const VIDEO_KEPT_DEFAULTS = [false, ''];

/** Menu: Video → Lista videor i Google Drive. */
function listDriveVideos() {
  const videoFolder = DriveApp.getFolderById(requiredProperty_('VIDEO_FOLDER_ID'));
  const sheet = getOrCreateSheet_(VIDEO_SHEET_NAME, VIDEO_HEADERS);
  const kept = keptVideoColumns_(sheet);
  const found = listDriveFiles_(videoFolder, 'video/');
  found.sort(function (a, b) { return a.relPath.localeCompare(b.relPath); });
  // Duration as H:MM:SS:FF (25 fps); '' for a video Drive has not processed or cannot read.
  const durations = driveVideoDurationsMs_(found.map(function (f) { return f.id; }));
  const rows = found.map(function (f) {
    const ms = durations[f.id];
    return [f.id, f.name, f.modified, String(f.size), f.relPath,
            ms == null ? '' : formatTimecode_(ms, 25)];
  });

  // Plain text, so the timestamps stay as written.
  const lastRow = sheet.getLastRow();
  writeTextRows_(sheet, rows, VIDEO_DRIVE_HEADERS.length);
  writeKeptVideoColumns_(sheet, found.map(function (f) { return kept[f.id] || VIDEO_KEPT_DEFAULTS; }), lastRow);
  toast_('Listade ' + found.length + ' videor från Drive', 'Video');
}

/** Each row's G-H (sent to DMC? and its status), by Drive file ID: [true/false, text]. */
function keptVideoColumns_(sheet) {
  const kept = keptColumnsByKey_(sheet, VIDEO_DRIVE_HEADERS.length + 1, VIDEO_KEPT_DEFAULTS.length);
  Object.keys(kept).forEach(function (id) { kept[id] = [kept[id][0] === true, String(kept[id][1])]; });
  return kept;
}

/** Writes G-H below the header (G as checkboxes, H as text) and clears them under the last row. */
function writeKeptVideoColumns_(sheet, rows, oldLastRow) {
  writeKeptColumns_(sheet, VIDEO_DRIVE_HEADERS.length + 1, VIDEO_KEPT_DEFAULTS.length, rows, oldLastRow,
    { checkboxes: [0], text: [1] });
}
