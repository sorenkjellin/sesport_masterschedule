// Drive "Video" folder -> the "Video Assets" sheet.
// Lists every video in the folder (VIDEO_FOLDER_ID) and its subfolders, sorted by folder path and
// name. The list is rebuilt on every run: columns A-F come from Drive, and Recipient status (H, see
// recipients.js) moves along with its video, matched by Drive file ID. A video gone
// from Drive loses its row. The Action dropdown (G) sends a clip to a recipient, such as Playout, or
// deletes it (actions.js).

const VIDEO_SHEET_NAME = 'Video Assets';
// Drive file ID, Name, Drive modified, Size, Path, Duration (A-F, from Drive), then Action (G, a
// dropdown holding no value of its own) and Recipient status (H, kept with its video):
// videoAssetHeaders_() in recipients.js. A recipient's job finds its rows by column A and the status
// column by header.
const VIDEO_DRIVE_HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Path', 'Duration'];

/** Whether a Drive file is a video: video/…, or MXF, which Drive calls application/mxf. */
function isVideoFile_(mimeType) {
  return /^video\//.test(mimeType || '') || mimeType === 'application/mxf';
}

/** Menu: Video → Lista videor i Google Drive. */
function listDriveVideos() {
  const videoFolder = DriveApp.getFolderById(requiredProperty_('VIDEO_FOLDER_ID'));
  // Recipient status belongs to its video; read by header, so it also survives a change of layout.
  // A recipient's former own column (e.g. DMC-status) carries over, with its name put first.
  const formers = [];
  recipients_().forEach(function (r) {
    (r.formerHeaders || []).forEach(function (h) { formers.push({ header: h, prefix: r.label + ': ' }); });
  });
  const kept = keptColumnsByHeader_(getOrCreateSheet_(VIDEO_SHEET_NAME),
    [RECIPIENT_STATUS_HEADER].concat(formers.map(function (f) { return f.header; })));
  const statusOf = function (id) {
    const values = kept[id] || [];
    if (String(values[0] || '') !== '') return String(values[0]);
    for (let i = 0; i < formers.length; i++) {
      const v = String(values[i + 1] || '');
      if (v) return v.indexOf(formers[i].prefix) === 0 ? v : formers[i].prefix + v;
    }
    return '';
  };
  const headers = videoAssetHeaders_();
  const sheet = getOrCreateSheet_(VIDEO_SHEET_NAME, headers);
  clearColumnsAfter_(sheet, headers.length);
  const found = listDriveFiles_(videoFolder, '').filter(function (f) { return isVideoFile_(f.mimeType); });
  found.sort(function (a, b) { return a.relPath.localeCompare(b.relPath); });
  // Duration in whole seconds, as a Sheets duration (shown as h:mm:ss, and summable): the probed one
  // stored on the file (mediaProbe.js), else Drive's; '' when neither is known (Drive has none for MXF).
  let stored = {};
  try {
    stored = driveFolderProperties_(videoFolder);
  } catch (e) {
    console.log('Stored durations not read: ' + e.message);
  }
  const durations = driveVideoDurationsMs_(found.map(function (f) { return f.id; }));
  const rows = found.map(function (f) { return [f.id, f.name, f.modified, String(f.size), f.relPath]; });

  // Plain text, so the timestamps stay as written; then the durations, as durations.
  const lastRow = sheet.getLastRow();
  const durationCol = VIDEO_DRIVE_HEADERS.indexOf('Duration') + 1;
  writeTextRows_(sheet, rows, durationCol - 1);
  writeKeptColumns_(sheet, durationCol, 1, found.map(function (f) {
    const seconds = (stored[f.id] || {})[PROBE_PROPERTY_KEYS.Duration];
    if (seconds !== undefined && seconds !== '') return [durationDays_(Number(seconds))];
    const ms = durations[f.id];
    return [ms == null ? '' : durationDays_(Math.floor(ms / 1000))];
  }), lastRow, { formats: { 0: durationFormat_() } });
  // Each file's modified time links to its preview in Drive (the name carries notes; see validation.js).
  writeLinkedCells_(sheet, 2, 3, found.map(function (f) { return { text: f.modified, url: driveFileUrl_(f.id) }; }));
  // G: the Action dropdown, empty (the old checkboxes go); H: each video's Recipient status.
  writeKeptColumns_(sheet, ACTION_COL, 2, found.map(function (f) { return ['', statusOf(f.id)]; }), lastRow, { text: [1] });
  setVideoActionDropdowns_(sheet, found.length, lastRow);
  // Path in dark grey for a video without stored probe results (Video → Analysera videor utan
  // metadata fills them in), black for one with them.
  const pathCol = VIDEO_DRIVE_HEADERS.indexOf('Path') + 1;
  if (found.length) {
    sheet.getRange(2, pathCol, found.length, 1).setFontColors(found.map(function (f) {
      return [(stored[f.id] || {})[PROBED_PROPERTY] ? null : VALIDATION_COLORS.greyText];
    }));
  }
  if (lastRow > found.length + 1) sheet.getRange(found.length + 2, pathCol, lastRow - found.length - 1, 1).setFontColor(null);
  toast_('Listade ' + found.length + ' videor från Drive', 'Video');
}

/**
 * The Action dropdown (G) on every listed video (rows 2 .. count + 1): a choice per recipient, and
 * Delete; none below, down to oldLastRow.
 */
function setVideoActionDropdowns_(sheet, count, oldLastRow) {
  if (count) {
    const choices = recipients_().map(function (r) { return r.choice; }).concat([DELETE_CHOICE]);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(choices, true).setAllowInvalid(false).build();
    // Clearing first takes away the checkboxes column G held before it was Action.
    sheet.getRange(2, ACTION_COL, count, 1).clearDataValidations()
      .setDataValidations(Array.from({ length: count }, function () { return [rule]; }));
  }
  if (oldLastRow > count + 1) sheet.getRange(count + 2, ACTION_COL, oldLastRow - count - 1, 1).clearDataValidations();
}
