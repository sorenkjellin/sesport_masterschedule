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
  const checks = postDeployChecks_();
  const check = checks.check;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Settings ---
  check('Script Properties', function () {
    const required = ['DRIVE_FOLDER_ID', 'ELMO_URL', 'ELMO_KEY', 'ELMO_PROJECT', 'ELMO_CACHEAROO_URL',
                      'EPG_MAIL_ENDPOINT'];
    const result = scriptPropertiesCheck_(required, ['ELMO_FOLDER', 'EPG_MAIL_CC'], [TRASHED_KORSCHEMAN_KEY]);
    ['ELMO_URL', 'ELMO_CACHEAROO_URL'].forEach(function (k) {
      if (!/^https?:\/\//.test(requiredProperty_(k))) throw new Error(k + ' is not a URL');
    });
    return result;
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
    const expected = [[1, 'date'], [2, 'start'], [3, 'end'], [4, 'genre'], [5, 'league'], [10, 'live']];
    const wrong = expected.filter(function ([col, key]) {
      return COLUMNS[key].map(function (n) { return n.toLowerCase(); }).indexOf(headers[col - 1].toLowerCase()) < 0;
    }).map(function ([col, key]) {
      return `column ${columnLetter_(col)} should be "${COLUMNS[key][0]}", is "${headers[col - 1]}"`;
    });
    // Column S: the header may be anything the team calls it, but every filled cell must be a
    // körschema ID (a UUID, as newKalenderId_ writes), or the column has moved.
    const idLetter = columnLetter_(KALENDER_ID_COL);
    const idHeader = headers[KALENDER_ID_COL - 1];
    const idHeaders = ['', KALENDER_ID_HEADER.toLowerCase(), 'id'];
    if (idHeaders.indexOf(idHeader.toLowerCase()) < 0) {
      wrong.push(`column ${idLetter} should be "${KALENDER_ID_HEADER}", "ID" or empty, is "${idHeader}"`);
    }
    const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    const ids = sheet.getLastRow() > 1
      ? sheet.getRange(2, KALENDER_ID_COL, sheet.getLastRow() - 1, 1).getDisplayValues() : [];
    const badRow = ids.findIndex(function (r) { return r[0] !== '' && !uuid.test(String(r[0]).trim()); });
    if (badRow >= 0) {
      wrong.push(`column ${idLetter} row ${badRow + 2} is "${ids[badRow][0]}", not a körschema ID`);
    }
    if (wrong.length) throw new Error(wrong.join('; '));
    const idCount = ids.filter(function (r) { return r[0] !== ''; }).length;
    return `A, B, C, D, E, J as expected; ${idLetter} "${idHeader}" holds ${idCount} körschema IDs`;
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
    const linked = readProgramRows_(ss).filter(function (r) { return r.fileId; }).length;
    return `${rows} rows, ${linked} with a file linked in G; F is "${PROGRAM_LENGTH_HEADER}"`;
  });

  // --- Services and access ---
  check('Sheets advanced service', function () {
    Sheets.Spreadsheets.get(ss.getId(), { fields: 'spreadsheetId' });
    return 'enabled and authorized';
  });

  check('Drive: image folder', function () {
    return `"${DriveApp.getFolderById(config_().folderId).getName()}"`;
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
    const warnings = [];
    if (!has('handleEdit', ss.getId())) warnings.push('handleEdit is not installed by you (run installEditTrigger)');
    if (!has('checkProgramLengths')) {
      warnings.push('checkProgramLengths is not installed by you (run installProgramLengthTrigger)');
    }
    // The asset watcher moved to the Video Assets spreadsheet (see assetWatcher.js).
    if (has('checkForNewAssets')) {
      warnings.push('you still have a checkForNewAssets trigger here; the watcher moved to the Video Assets ' +
        'spreadsheet, and this one removes itself on its next run');
    }
    if (triggers.length >= 17) {
      warnings.push(`you have ${triggers.length} of the 20 triggers Google allows (archive old Kalender rows)`);
    }
    if (warnings.length) return { warn: warnings.join('; ') };
    return `handleEdit, checkProgramLengths; ${triggers.length} of 20 triggers`;
  });

  return checks.finish();
}
