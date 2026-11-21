/**
 * Post-deploy tests: run after each deploy (scripts/push.sh video-assets) and whenever something
 * looks wrong. Open this file in the Apps Script editor and press Run; runPostDeployTests is
 * preselected because it is the first function in the file. Results go to the execution log, and
 * the run is marked failed when a check fails.
 *
 * Every check only reads: nothing is written to the sheet, no job is started and no mail is sent.
 */
function runPostDeployTests() {
  const checks = postDeployChecks_();
  const check = checks.check;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  check('Script Properties', function () {
    return scriptPropertiesCheck_(
      ['VIDEO_FOLDER_ID', 'DMC_GCP_PROJECT', 'WATCH_FOLDER_ID', 'WATCH_RECIPIENTS'],
      ['DMC_REGION', 'DMC_JOB', 'WATCH_SUBJECT_PREFIX', 'WATCH_RUN_AS'],
      [DRIVE_WATCH_STATE_KEY]);
  });

  check('Video Assets sheet', function () {
    const sheet = ss.getSheetByName(VIDEO_SHEET_NAME);
    if (!sheet) return { warn: `"${VIDEO_SHEET_NAME}" is missing; run Video → Lista videor i Google Drive` };
    const have = sheet.getRange(1, 1, 1, VIDEO_HEADERS.length).getDisplayValues()[0];
    if (have.join('|') !== VIDEO_HEADERS.join('|')) {
      throw new Error(`the headers are "${have.join('", "')}", not "${VIDEO_HEADERS.join('", "')}"; ` +
        'run Video → Lista videor i Google Drive');
    }
    return `${Math.max(sheet.getLastRow() - 1, 0)} videos listed`;
  });

  check('Drive: Video folder', function () {
    return `"${DriveApp.getFolderById(requiredProperty_('VIDEO_FOLDER_ID')).getName()}"`;
  });

  check('Drive: watched asset folder', function () {
    const watched = DriveApp.getFolderById(requiredProperty_('WATCH_FOLDER_ID'));
    const detail = `"${watched.getName()}", mails ${requiredProperty_('WATCH_RECIPIENTS')}`;
    // The subfolders whose new files refresh an inbox tab (inbox.js).
    const missing = INBOX_LISTS.filter(function (l) { return !watched.getFoldersByName(l.folder).hasNext(); })
      .map(function (l) { return `"${l.folder}" (for ${l.sheet})`; });
    if (missing.length) return { warn: `${detail}; no subfolder ${missing.join(', ')}` };
    return `${detail}; lists ` + INBOX_LISTS.map(function (l) { return `"${l.folder}" in ${l.sheet}`; }).join(' and ');
  });

  check('DMC transfer: Cloud Run job', function () {
    const cfg = dmcJob_();
    const job = getCloudRunJob_(cfg);
    const container = ((job.template && job.template.template && job.template.template.containers) || [])[0] || {};
    const env = {};
    (container.env || []).forEach(function (e) { env[e.name] = e.value === undefined ? '(secret)' : e.value; });
    const missing = ['FTP_HOST', 'FTP_USER', 'FTP_PASSWORD', 'SPREADSHEET_ID'].filter(function (k) { return !env[k]; });
    if (missing.length) throw new Error(`job ${cfg.name} has no ${missing.join(', ')}`);
    if (env.SPREADSHEET_ID !== ss.getId()) {
      throw new Error(`job ${cfg.name} writes its status to spreadsheet ${env.SPREADSHEET_ID}, not this one ` +
        `(${ss.getId()}); set SPREADSHEET_ID in job.env and run scripts/deploy-dmc-transfer.sh`);
    }
    // latestCreatedExecution is {} until the job has run.
    const last = job.latestCreatedExecution || {};
    const lastRun = last.createTime
      ? `last run ${last.createTime}${last.completionStatus ? ' ' + last.completionStatus : ''}`
      : 'never run';
    return `${cfg.name} in ${cfg.project}/${cfg.region} sends to ${env.FTP_HOST}` +
      (env.FTP_DIR ? '/' + env.FTP_DIR : '') + '; ' + lastRun;
  });

  // The job's service account needs this spreadsheet, the Video folder and the FTP password; the
  // person running this needs permission to start the job.
  check('DMC transfer: permissions', function () {
    let video = null;
    try {
      video = DriveApp.getFolderById(requiredProperty_('VIDEO_FOLDER_ID'));
    } catch (e) { /* reported by "Drive: Video folder" */ }
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

  // getProjectTriggers lists only the triggers of the user running this, so a trigger someone
  // else installed shows up as missing (a warning, not a failure).
  check('Triggers', function () {
    const has = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getHandlerFunction() === 'checkForNewAssets';
    });
    // The watcher should run as WATCH_RUN_AS only; other accounts' triggers remove themselves.
    const runAs = optionalProperty_('WATCH_RUN_AS', '').trim();
    const me = Session.getEffectiveUser().getEmail();
    const watcherIsMine = !runAs || me.toLowerCase() === runAs.toLowerCase();
    const warnings = [];
    if (watcherIsMine && !has) warnings.push('checkForNewAssets is not installed by you (run installWatcherTrigger)');
    if (!watcherIsMine && has) {
      warnings.push(`you have a checkForNewAssets trigger, but the watcher runs as ${runAs}; it removes itself on its next run`);
    }
    if (!runAs) warnings.push('WATCH_RUN_AS is not set, so every user\'s checkForNewAssets trigger runs the watcher');
    if (warnings.length) return { warn: warnings.join('; ') };
    return watcherIsMine
      ? `checkForNewAssets (WATCH_RUN_AS: the watcher runs as you, ${runAs})`
      : `not checkForNewAssets (WATCH_RUN_AS: the watcher runs as ${runAs})`;
  });

  return checks.finish();
}
