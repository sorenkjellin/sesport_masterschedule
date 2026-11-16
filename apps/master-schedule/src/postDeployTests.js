/**
 * Post-deploy tests: run after each deploy (scripts/push.sh) and whenever something looks wrong.
 * Open this file in the Apps Script editor and press Run; runPostDeployTests is preselected because
 * it is the first function in the file. Results go to the execution log, and the run is marked
 * failed when a check fails.
 *
 * Every check only reads: nothing is written to the sheets, Elmo or Cachearoo, and no mail is sent.
 *
 * The unit tests (npm test) cover the logic; these cover what only the live project can tell:
 * settings, sheets and their columns, access to Drive and Elmo, and the triggers.
 */
function runPostDeployTests() {
  const results = [];
  const check = function (name, fn) {
    try {
      const r = fn();
      results.push(r && r.warn ? { status: 'WARN', name: name, detail: r.warn }
                               : { status: 'PASS', name: name, detail: r || '' });
    } catch (e) {
      results.push({ status: 'FAIL', name: name, detail: e.message });
    }
  };
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Settings ---
  check('Script Properties', function () {
    const required = ['DRIVE_FOLDER_ID', 'ELMO_URL', 'ELMO_KEY', 'ELMO_PROJECT', 'ELMO_CACHEAROO_URL',
                      'EPG_MAIL_ENDPOINT'];
    const missing = required.filter(function (k) { return !optionalProperty_(k, ''); });
    if (missing.length) throw new Error('Not set: ' + missing.join(', '));
    ['ELMO_URL', 'ELMO_CACHEAROO_URL'].forEach(function (k) {
      if (!/^https?:\/\//.test(requiredProperty_(k))) throw new Error(k + ' is not a URL');
    });
    return 'all ' + required.length + ' set';
  });

  // --- Sheets ---
  check('Sheets exist', function () {
    const required = [SCHEDULE_SHEET, 'Program', 'Ligor&Lag', 'EPG', 'Playout',
                      CONFIG.SETTINGS_SHEET, CONFIG.DEST_SHEET];
    const missing = required.filter(function (n) { return !ss.getSheetByName(n); });
    if (missing.length) throw new Error('Missing: ' + missing.join(', '));
    if (!ss.getSheetByName(SHEET_NAME)) return { warn: `"${SHEET_NAME}" is missing; run Bilder → Lista bilder` };
    return required.length + 1 + ' sheets';
  });

  check('Kalender columns found by header (coming up)', function () {
    const schedule = readSchedule(); // throws without Datum and Börjar
    const missing = Object.keys(COLUMNS).filter(function (k) { return schedule.cols[k] < 0; })
      .map(function (k) { return COLUMNS[k][0]; });
    if (missing.length) return { warn: 'No column for: ' + missing.join(', ') + ' (left empty in the feed)' };
    return Object.keys(COLUMNS).length + ' columns, ' + schedule.rows.length + ' rows';
  });

  check('Kalender columns used by position (dropdowns, Live, archive)', function () {
    const sheet = ss.getSheetByName(SCHEDULE_SHEET);
    const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), KALENDER_ID_COL)).getDisplayValues()[0]
      .map(function (h) { return String(h).trim(); });
    const expected = [[1, 'date'], [2, 'start'], [4, 'genre'], [5, 'league'], [10, 'live']];
    const wrong = expected.filter(function ([col, key]) {
      return COLUMNS[key].map(function (n) { return n.toLowerCase(); }).indexOf(headers[col - 1].toLowerCase()) < 0;
    }).map(function ([col, key]) {
      return `column ${columnLetter_(col)} should be "${COLUMNS[key][0]}", is "${headers[col - 1]}"`;
    });
    const idHeader = headers[KALENDER_ID_COL - 1];
    if (idHeader && idHeader !== KALENDER_ID_HEADER) {
      wrong.push(`column ${columnLetter_(KALENDER_ID_COL)} should be "${KALENDER_ID_HEADER}" or empty, is "${idHeader}"`);
    }
    if (wrong.length) throw new Error(wrong.join('; '));
    return 'A, B, D, E, J and S as expected';
  });

  check('Archive cutoff date', function () {
    const cell = CONFIG.SETTINGS_SHEET + '!' + CONFIG.DATE_CELL;
    const value = ss.getSheetByName(CONFIG.SETTINGS_SHEET).getRange(CONFIG.DATE_CELL).getValue();
    if (isNaN(new Date(value).getTime()) || value === '') throw new Error(cell + ' is not a date: "' + value + '"');
    return cell + ' = ' + Utilities.formatDate(new Date(value), TIME_ZONE, 'yyyy-MM-dd');
  });

  check('Program has körschema templates', function () {
    const program = ss.getSheetByName('Program');
    const rows = program.getLastRow() - 1;
    if (rows < 1) return { warn: 'Program has no rows' };
    return rows + ' rows';
  });

  // --- Services and access ---
  check('Sheets advanced service', function () {
    Sheets.Spreadsheets.get(ss.getId(), { fields: 'spreadsheetId' });
    return 'enabled and authorized';
  });

  check('Drive: image and video folders', function () {
    const images = DriveApp.getFolderById(config_().folderId);
    const video = findSiblingFolder_(images, VIDEO_FOLDER_NAME);
    if (!video) return { warn: `"${images.getName()}" opens, but no "${VIDEO_FOLDER_NAME}" folder next to it` };
    return `"${images.getName()}" and "${video.getName()}"`;
  });

  check('Drive: watched asset folder', function () {
    return '"' + DriveApp.getFolderById(DRIVE_WATCH_CONFIG.PARENT_FOLDER_ID).getName() + '"';
  });

  check('Elmo: file listing', function () {
    const cfg = config_();
    const files = getElmoFiles_(cfg);
    const n = Object.keys(files).length;
    if (!n) return { warn: `project "${cfg.project}" answers but has no files` };
    return n + ' files in "' + cfg.project + '"';
  });

  check('Coming up: dry run (nothing published)', function () {
    const project = requiredProperty_('ELMO_PROJECT');
    const schedule = readSchedule();
    const uploaded = uploadedImages();
    const shows = buildComingUp(schedule, uploaded, project, new Date()).shows;
    if (!shows.length) return { warn: 'no upcoming shows in Kalender' };
    const noImage = shows.filter(function (s) { return !s.image; }).length;
    const detail = shows.length + ' shows, first ' + shows[0].start;
    return noImage ? { warn: detail + '; ' + noImage + ' without an image' } : detail;
  });

  // --- Triggers ---
  // getProjectTriggers lists only the triggers of the user running this, so a trigger someone
  // else installed shows up as missing (a warning, not a failure).
  check('Triggers', function () {
    const triggers = ScriptApp.getProjectTriggers();
    const has = function (handler, sourceId) {
      return triggers.some(function (t) {
        return t.getHandlerFunction() === handler && (!sourceId || t.getTriggerSourceId() === sourceId);
      });
    };
    const missing = [];
    if (!has('handleEdit', ss.getId())) missing.push('handleEdit (run installEditTrigger)');
    if (!has('checkForNewAssets')) missing.push('checkForNewAssets (time-driven)');
    if (missing.length) return { warn: 'Not installed by you: ' + missing.join(', ') };
    return 'handleEdit and checkForNewAssets';
  });

  // --- Report ---
  const width = Math.max.apply(null, results.map(function (r) { return r.name.length; }));
  results.forEach(function (r) {
    console.log(r.status.padEnd(4) + '  ' + r.name.padEnd(width) + '  ' + r.detail);
  });
  const count = function (s) { return results.filter(function (r) { return r.status === s; }).length; };
  const summary = `${count('PASS')} passed, ${count('WARN')} warnings, ${count('FAIL')} failed`;
  console.log(summary);
  if (count('FAIL')) {
    throw new Error('Post-deploy tests failed: ' + results.filter(function (r) { return r.status === 'FAIL'; })
      .map(function (r) { return r.name; }).join(', '));
  }
  return summary;
}
