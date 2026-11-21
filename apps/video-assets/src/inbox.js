// The Inbox folder (WATCH_FOLDER_ID) -> the "Video Inbox" and "Image Inbox" tabs.
// When the asset watcher finds new files in Inbox/Video, it lists everything in Inbox/Video (and
// its subfolders) in "Video Inbox"; new files in Inbox/Images likewise refresh "Image Inbox". This
// happens before the mail goes out, and the mail links to each tab that was updated.
// Inkorg → Uppdatera inkorgarna refreshes both by hand, e.g. after files were moved away or deleted
// (the watcher only notices new files).
// Like "Video Assets", a tab is rebuilt on every run, sorted by folder path and name. Every file is
// listed, whatever its type, so nothing in the inbox is hidden.

// Each Inbox subfolder that is listed, the tab it is listed in, and whether to look up durations.
const INBOX_LISTS = [
  { folder: 'Video', sheet: 'Video Inbox', duration: true, probe: true }, // mediaProbe.js
  { folder: 'Images', sheet: 'Image Inbox', duration: false },
];
const INBOX_HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Path'];

/** Menu: Inkorg → Uppdatera inkorgarna. Lists both Inbox subfolders. */
function refreshInboxLists() {
  const results = listInboxes_(INBOX_LISTS.map(function (l) { return l.folder; }));
  toast_(results.map(function (r) {
    if (r.error) return `${r.sheet}: ${r.error}`;
    let text = `${r.sheet}: ${r.count} filer`;
    if (r.probe && r.probe.started) text += ` (${r.probe.started} analyseras)`;
    if (r.probe && r.probe.error) text += ` (analysen startade inte: ${r.probe.error})`;
    return text;
  }).join(', '), 'Inkorg');
}

/**
 * Lists the Inbox subfolders named in folderNames (as in INBOX_LISTS) in their tabs. Returns one
 * result per tab: { sheet, count, url } when it was rebuilt, { sheet, error } when it was not (the
 * subfolder is missing or listing failed). One tab failing does not stop the other.
 */
function listInboxes_(folderNames) {
  const inbox = DriveApp.getFolderById(requiredProperty_('WATCH_FOLDER_ID'));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return INBOX_LISTS.filter(function (l) { return folderNames.indexOf(l.folder) >= 0; }).map(function (l) {
    try {
      const found = inbox.getFoldersByName(l.folder);
      if (!found.hasNext()) return { sheet: l.sheet, error: `ingen mapp "${l.folder}" i "${inbox.getName()}"` };
      const listed = listInboxFolder_(found.next(), l);
      return { sheet: l.sheet, count: listed.count, probe: listed.probe,
               url: ss.getUrl() + '#gid=' + ss.getSheetByName(l.sheet).getSheetId() };
    } catch (e) {
      return { sheet: l.sheet, error: e.message };
    }
  });
}

/** The lines the watcher's mail gets about the inbox tabs (results of listInboxes_), or ''. */
function inboxMailNote_(results) {
  const done = results.filter(function (r) { return !r.error; });
  const failed = results.filter(function (r) { return r.error; });
  const lines = [];
  if (done.length) {
    lines.push(`Updated in ${SpreadsheetApp.getActiveSpreadsheet().getName()}:`);
    done.forEach(function (r) {
      let line = `  ${r.sheet} (${r.count} file${r.count === 1 ? '' : 's'}): ${r.url}`;
      if (r.probe && r.probe.started) {
        line += ` (probing ${r.probe.started} video${r.probe.started === 1 ? '' : 's'}; the results appear in the tab)`;
      }
      lines.push(line);
    });
  }
  done.filter(function (r) { return r.probe && r.probe.error; }).forEach(function (r) {
    lines.push(`The videos in ${r.sheet} could not be probed: ${r.probe.error}`);
  });
  failed.forEach(function (r) { lines.push(`${r.sheet} could not be updated: ${r.error}`); });
  return lines.join('\n');
}

/**
 * Rebuilds list.sheet from the files in folder and its subfolders, and for a list with probe,
 * starts probing the videos not probed yet. Returns { count } (files listed), with probe: the result
 * of startProbes_ for such a list.
 */
function listInboxFolder_(folder, list) {
  const driveHeaders = list.duration ? INBOX_HEADERS.concat(['Duration']) : INBOX_HEADERS;
  const headers = list.probe ? driveHeaders.concat(PROBE_HEADERS) : driveHeaders;
  const sheet = getOrCreateSheet_(list.sheet, headers);
  // The probe columns belong to their video, so they move with it (by Drive file ID).
  const kept = list.probe ? keptColumnsByKey_(sheet, driveHeaders.length + 1, PROBE_HEADERS.length) : {};
  const oldLastRow = sheet.getLastRow();
  const found = listDriveFiles_(folder, '');
  found.sort(function (a, b) { return a.relPath.localeCompare(b.relPath); });
  // Duration as H:MM:SS:FF (25 fps); '' for a file that is not a video or not processed yet.
  const durations = list.duration ? driveVideoDurationsMs_(found.map(function (f) { return f.id; })) : {};
  const rows = found.map(function (f) {
    const row = [f.id, f.name, f.modified, String(f.size), f.relPath];
    if (list.duration) row.push(durations[f.id] == null ? '' : formatTimecode_(durations[f.id], 25));
    return row;
  });
  writeTextRows_(sheet, rows, driveHeaders.length);
  if (!list.probe) return { count: rows.length };
  const blank = PROBE_HEADERS.map(function () { return ''; });
  const probeRows = found.map(function (f) { return kept[f.id] || blank; });
  writeKeptColumns_(sheet, driveHeaders.length + 1, PROBE_HEADERS.length, probeRows, oldLastRow,
    { text: [PROBE_HEADERS.length - 1] });
  const statuses = probeRows.map(function (r) { return r[PROBE_HEADERS.length - 1]; });
  return { count: rows.length, probe: startProbes_(sheet, found, statuses) };
}
