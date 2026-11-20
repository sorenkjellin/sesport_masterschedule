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
                      'EPG_MAIL_ENDPOINT', 'WATCH_FOLDER_ID', 'WATCH_RECIPIENTS'];
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
    return `A, B, D, E, J as expected; ${idLetter} "${idHeader}" holds ${idCount} körschema IDs`;
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
    const watched = DriveApp.getFolderById(requiredProperty_('WATCH_FOLDER_ID'));
    const detail = `"${watched.getName()}", mails ${requiredProperty_('WATCH_RECIPIENTS')}`;
    if (watched.getId() === config_().folderId) return detail + ' (the same folder as DRIVE_FOLDER_ID)';
    return detail;
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

  check('DMC transfer: Cloud Run job', function () {
    if (!optionalProperty_('DMC_GCP_PROJECT', '')) {
      return { warn: 'DMC_GCP_PROJECT is not set, so Video → Skicka markerade till DMC cannot start' };
    }
    const video = ss.getSheetByName(VIDEO_SHEET_NAME);
    const want = VIDEO_HEADERS.slice(DMC_SEND_COL - 1, DMC_STATUS_COL);
    const have = video ? video.getRange(1, DMC_SEND_COL, 1, want.length).getDisplayValues()[0] : [];
    if (have.join('|') !== want.join('|')) {
      return { warn: `"${VIDEO_SHEET_NAME}" G-H are not "${want.join('", "')}"; run Video → Lista videor i Google Drive` };
    }
    const cfg = dmcJob_();
    const job = getCloudRunJob_(cfg);
    const container = ((job.template && job.template.template && job.template.template.containers) || [])[0] || {};
    const env = {};
    (container.env || []).forEach(function (e) { env[e.name] = e.value === undefined ? '(secret)' : e.value; });
    const missing = ['FTP_HOST', 'FTP_USER', 'FTP_PASSWORD', 'SPREADSHEET_ID'].filter(function (k) { return !env[k]; });
    if (missing.length) throw new Error(`job ${cfg.name} has no ${missing.join(', ')}`);
    if (env.SPREADSHEET_ID !== ss.getId()) {
      throw new Error(`job ${cfg.name} writes its status to spreadsheet ${env.SPREADSHEET_ID}, not this one`);
    }
    // latestCreatedExecution is {} until the job has run.
    const last = job.latestCreatedExecution || {};
    const lastRun = last.createTime
      ? `last run ${last.createTime}${last.completionStatus ? ' ' + last.completionStatus : ''}`
      : 'never run';
    return `${cfg.name} in ${cfg.project}/${cfg.region} sends to ${env.FTP_HOST}` +
      (env.FTP_DIR ? '/' + env.FTP_DIR : '') + '; ' + lastRun;
  });

  // The job's service account needs the spreadsheet, the Video folder and the FTP password; the
  // person running this needs permission to start the job.
  check('DMC transfer: permissions', function () {
    if (!optionalProperty_('DMC_GCP_PROJECT', '')) return { warn: 'skipped: DMC_GCP_PROJECT is not set' };
    const images = DriveApp.getFolderById(config_().folderId);
    const video = findSiblingFolder_(images, VIDEO_FOLDER_NAME);
    const access = dmcAccess_(video ? [{ id: video.getId(), name: video.getName(), role: 'reader' }] : []);
    const problems = access.problems.map(function (p) { return dmcProblemText_(p, 'en'); });
    const notes = [];
    let secretChecked = false;
    if (access.account) {
      // The secret named in the job's FTP_PASSWORD; its policy is readable only by Secret Manager admins.
      const container = ((access.job.template.template.containers) || [])[0] || {};
      const pw = (container.env || []).filter(function (e) { return e.name === 'FTP_PASSWORD'; })[0];
      const ref = pw && pw.valueSource && pw.valueSource.secretKeyRef;
      if (ref) {
        const secret = String(ref.secret).split('/').pop();
        try {
          const policy = getSecretIamPolicy_(dmcJob_().project, secret);
          const member = 'serviceAccount:' + access.account;
          const ok = (policy.bindings || []).some(function (b) {
            return b.role === 'roles/secretmanager.secretAccessor' && (b.members || []).indexOf(member) >= 0;
          });
          if (!ok) problems.push(`${access.account} may not read the secret ${secret} (roles/secretmanager.secretAccessor)`);
          secretChecked = true;
        } catch (e) {
          notes.push(`secret ${secret} not checked (no access to its policy)`);
        }
      }
    }
    if (problems.length) throw new Error(problems.join('; '));
    access.unchecked.forEach(function (name) { notes.push(`sharing of "${name}" not checked`); });
    const detail = `you may start the job; ${access.account} can edit this spreadsheet` +
      (video ? `, read "${video.getName()}"` : '') + (secretChecked ? ' and read the password' : '');
    return notes.length ? { warn: detail + '; ' + notes.join('; ') } : detail;
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
    // The watcher should run as WATCH_RUN_AS only; other accounts' triggers remove themselves.
    const runAs = optionalProperty_('WATCH_RUN_AS', '').trim();
    const me = Session.getEffectiveUser().getEmail();
    const watcherIsMine = !runAs || me.toLowerCase() === runAs.toLowerCase();
    const missing = [];
    const notes = [];
    if (!has('handleEdit', ss.getId())) missing.push('handleEdit (run installEditTrigger)');
    if (watcherIsMine && !has('checkForNewAssets')) missing.push('checkForNewAssets (time-driven)');
    if (!watcherIsMine && has('checkForNewAssets')) {
      notes.push(`you have a checkForNewAssets trigger, but the watcher runs as ${runAs}; it removes itself on its next run`);
    }
    if (!runAs) notes.push('WATCH_RUN_AS is not set, so every user\'s checkForNewAssets trigger runs the watcher');
    const warnings = (missing.length ? ['Not installed by you: ' + missing.join(', ')] : []).concat(notes);
    const watcher = runAs ? 'the watcher runs as ' + runAs : 'checkForNewAssets';
    if (warnings.length) return { warn: warnings.join('; ') };
    return 'handleEdit and ' + (watcherIsMine ? 'checkForNewAssets' : 'not checkForNewAssets (' + watcher + ')');
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
