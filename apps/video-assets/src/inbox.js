// The Inbox folder (WATCH_FOLDER_ID) -> the "Video Inbox" and "Image Inbox" tabs.
// When the asset watcher finds new files in Inbox/Video, it lists everything in Inbox/Video (and
// its subfolders) in "Video Inbox"; new files in Inbox/Images likewise refresh "Image Inbox". This
// happens before the mail goes out, and the mail links to each tab that was updated.
// Inkorg → Uppdatera inkorgarna refreshes both by hand, e.g. after files were moved away or deleted
// (the watcher only notices new files).
// Like "Video Assets", a tab is rebuilt on every run, sorted by name. Every file is
// listed, whatever its type, so nothing in the inbox is hidden.

// Each Inbox subfolder that is listed, the tab it is listed in, and whether to look up durations.
const INBOX_LISTS = [
  { folder: 'Video', sheet: 'Video Inbox', duration: true, probe: true }, // mediaProbe.js
  { folder: 'Images', sheet: 'Image Inbox', duration: false },
];
// The inbox folders are flat, so no path: the name says it all.
const INBOX_HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size'];

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
 *
 * In a list with probe, Duration is one of the probe columns: the probe writes it for every video
 * (Drive has none for MXF). Until then it shows Drive's duration, if any. It is a Sheets duration
 * (whole seconds, shown as h:mm:ss), so durations can be summed.
 */
function listInboxFolder_(folder, list) {
  const kept = list.probe ? PROBE_HEADERS : [];
  const headers = INBOX_HEADERS.concat(list.probe ? kept : list.duration ? ['Duration'] : []);
  // The probe columns belong to their video, so they move with it (by Drive file ID). They are
  // read by header, before the header row is rewritten, so they survive a change of layout.
  const keptValues = kept.length ? keptColumnsByHeader_(getOrCreateSheet_(list.sheet), kept, ['Duration']) : {};
  const sheet = getOrCreateSheet_(list.sheet, headers);
  clearColumnsAfter_(sheet, headers.length);
  const oldLastRow = sheet.getLastRow();
  const found = listDriveFiles_(folder, '');
  found.sort(function (a, b) { return a.relPath.localeCompare(b.relPath); });
  // Drive's duration in whole seconds, as a Sheets duration; '' for a file that is not a video, an
  // MXF, or one Drive has not processed yet.
  const durations = list.duration ? driveVideoDurationsMs_(found.map(function (f) { return f.id; })) : {};
  const driveDuration = function (f) { return durations[f.id] == null ? '' : durationDays_(Math.floor(durations[f.id] / 1000)); };
  const rows = found.map(function (f) {
    const row = [f.id, f.name, f.modified, String(f.size)];
    if (list.duration && !list.probe) row.push(driveDuration(f));
    return row;
  });
  writeTextRows_(sheet, rows, list.probe ? INBOX_HEADERS.length : headers.length);
  // Each name links to its file in Drive.
  writeLinkedCells_(sheet, 2, 2, found.map(function (f) { return { text: f.name, url: driveFileUrl_(f.id) }; }));
  if (!list.probe) return { count: rows.length };

  const durationAt = kept.indexOf('Duration');
  const statusAt = kept.indexOf('Probe status');
  const keptRows = found.map(function (f) {
    const r = (keptValues[f.id] || kept.map(function () { return ''; })).slice();
    // Read as displayed (h:mm:ss, or an older timecode h:mm:ss:ff) and stored as a duration again.
    r[durationAt] = parseDuration_(r[durationAt]);
    if (r[durationAt] === '') r[durationAt] = driveDuration(f);
    return r;
  });
  const formats = {};
  formats[durationAt] = durationFormat_();
  writeKeptColumns_(sheet, INBOX_HEADERS.length + 1, kept.length, keptRows, oldLastRow,
    { text: [statusAt], formats: formats });
  const statuses = keptRows.map(function (r) { return r[statusAt]; });
  return { count: rows.length, probe: startProbes_(sheet, found, statuses) };
}
