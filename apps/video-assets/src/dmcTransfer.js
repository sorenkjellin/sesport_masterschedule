/**
 * Video Assets -> DMC's FTP server (FTPS, explicit TLS on port 21).
 *
 * Tick "Skicka till DMC" (column G) for the clips to send and choose Video → Skicka markerade till
 * DMC. Apps Script can neither speak FTP nor run for longer than 6 minutes, so the sending is done
 * by a Cloud Run job (services/dmc-transfer in the repo), which this starts with the ticked clips.
 * The job writes each clip's progress to "DMC-status" (column H):
 *   Köad <time> → Skickar … (with progress) → Skickad <time>, or Fel: <reason>.
 * Before queuing, it checks that the sender may start the job and that the spreadsheet and the
 * clips are shared with the job's service account (see dmcAccess_).
 * A file with the same name at DMC is overwritten. All clips go to the same folder there, so two
 * ticked clips with the same name are refused.
 *
 * Script Properties:
 *   DMC_GCP_PROJECT  ID of the Google Cloud project the job is in
 *   DMC_REGION       the job's region (optional, default europe-north1)
 *   DMC_JOB          the job's name (optional, default dmc-transfer)
 *
 * Setting up the job, its password and who may send: docs/dmc-transfer.md in the repo.
 */

// Columns of "Video Assets" (VIDEO_HEADERS in videos.js): G "Skicka till DMC", H "DMC-status".
const DMC_SEND_COL = 7;
const DMC_STATUS_COL = 8;

/** Menu: Video → Skicka markerade till DMC. Starts the transfer of the ticked clips. */
function sendCheckedToDmc() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VIDEO_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    throw new Error('Inga videor listade i "' + VIDEO_SHEET_NAME + '": välj Video → Lista videor i Google Drive först');
  }
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, DMC_STATUS_COL).getValues();
  const pick = pickDmcClips_(rows);
  if (pick.problems.length) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Skicka till DMC', pick.problems.join('\n\n') + '\n\nInget skickades.', ui.ButtonSet.OK);
    return;
  }
  if (!pick.clips.length) {
    toast_('Bocka i "Skicka till DMC" för klippen som ska skickas', 'DMC');
    return;
  }

  // Check the permissions first, so a missing share is reported here instead of in the job's log.
  const access = dmcAccess_(pick.clips.map(function (c) { return { id: c.id, name: c.name, role: 'reader' }; }));
  if (access.problems.length) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Skicka till DMC', access.problems.map(function (p) { return dmcProblemText_(p, 'sv'); })
      .join('\n\n') + '\n\nInget skickades.', ui.ButtonSet.OK);
    return;
  }

  const ids = pick.clips.map(function (c) { return c.id; });
  setDmcCells_(sheet, pick.clips, false, 'Köad ' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm'));
  try {
    runCloudRunJob_(dmcJob_(), {
      env: { FILE_IDS: ids.join(','), REQUESTED_BY: Session.getActiveUser().getEmail() },
      taskCount: ids.length, // one task per clip; the job's parallelism caps how many run at once
    });
  } catch (e) {
    setDmcCells_(sheet, pick.clips, true, 'Fel: överföringen startade inte: ' + e.message);
    throw e;
  }
  toast_(ids.length + ' klipp skickas till DMC. Följ dem i kolumnen DMC-status.', 'DMC');
}

/**
 * The ticked clips in rows (values of "Video Assets" A-H, from row 2), as { row, id, name }, and
 * why they cannot be sent: a clip already queued or being sent, or two clips with the same name
 * (they would overwrite each other at DMC).
 */
function pickDmcClips_(rows) {
  const clips = [];
  const problems = [];
  const byName = {};
  rows.forEach(function (r, i) {
    if (r[DMC_SEND_COL - 1] !== true || !r[0]) return;
    const clip = { row: i + 2, id: String(r[0]), name: String(r[1]) };
    const status = String(r[DMC_STATUS_COL - 1]);
    if (/^(Köad|Skickar)/.test(status)) {
      problems.push(`${clip.name} (rad ${clip.row}) är redan på väg: "${status}". ` +
        'Töm DMC-status om den har fastnat.');
    }
    (byName[clip.name] = byName[clip.name] || []).push(clip.row);
    clips.push(clip);
  });
  Object.keys(byName).forEach(function (name) {
    if (byName[name].length > 1) {
      problems.push(`${name} finns på rad ${byName[name].join(' och ')}: samma namn hos DMC, ` +
        'så de skulle skriva över varandra. Bocka i en i taget.');
    }
  });
  return { clips: clips, problems: problems };
}

function setDmcCells_(sheet, clips, ticked, status) {
  clips.forEach(function (c) {
    sheet.getRange(c.row, DMC_SEND_COL, 1, 2).setValues([[ticked, status]]);
  });
}

/** The Cloud Run job that sends clips to DMC (the `job` argument of the Cloud Run API in lib.js). */
function dmcJob_() {
  return {
    project: requiredProperty_('DMC_GCP_PROJECT'),
    region: optionalProperty_('DMC_REGION', 'europe-north1'),
    name: optionalProperty_('DMC_JOB', 'dmc-transfer'),
  };
}

// --- Permissions --------------------------------------------------------------
// The job runs as its own service account, which needs the spreadsheet (to write DMC-status) and
// the clips (to read them) shared with it; the person sending needs permission to start the job.
// These are checked before anything is queued, and by the post-deploy tests.

/**
 * What keeps the DMC job from working for the user running this, as a list of problems:
 *   { kind: 'job', detail }                 the job cannot be read (missing, or no permission)
 *   { kind: 'run' }                         the user may not start the job
 *   { kind: 'share', name, role, account }  an item is not shared with the job's service account
 * items: [{ id, name, role }] Drive items the job needs ('reader' or 'writer'); the spreadsheet
 * (writer) is always checked. Items whose sharing the user cannot see are listed in `unchecked`.
 */
function dmcAccess_(items) {
  const cfg = dmcJob_();
  const result = { account: '', problems: [], unchecked: [] };
  let job;
  try {
    job = getCloudRunJob_(cfg);
  } catch (e) {
    result.problems.push({ kind: 'job', detail: e.message });
    return result;
  }
  result.job = job;
  if (cloudRunJobPermissions_(cfg, ['run.jobs.runWithOverrides']).indexOf('run.jobs.runWithOverrides') < 0) {
    result.problems.push({ kind: 'run' });
  }
  const account = (job.template && job.template.template && job.template.template.serviceAccount) || '';
  result.account = account;
  if (!account) {
    result.unchecked.push('sharing (the job has no service account of its own)');
    return result;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [{ id: ss.getId(), name: ss.getName(), role: 'writer' }].concat(items || []).forEach(function (item) {
    let role;
    try {
      role = driveRoleOf_(item.id, account);
    } catch (e) {
      result.unchecked.push(item.name);
      return;
    }
    if (driveRoleRank_(role) < driveRoleRank_(item.role)) {
      result.problems.push({ kind: 'share', name: item.name, role: item.role, account: account });
    }
  });
  return result;
}

/** A problem from dmcAccess_ as a sentence, in Swedish ('sv', for the sheet) or English. */
function dmcProblemText_(p, lang) {
  const sv = lang === 'sv';
  const job = optionalProperty_('DMC_JOB', 'dmc-transfer');
  if (p.kind === 'job') {
    return sv ? 'Kan inte läsa överföringsjobbet ' + job + ': ' + p.detail
              : 'cannot read the job ' + job + ': ' + p.detail;
  }
  if (p.kind === 'run') {
    const who = Session.getActiveUser().getEmail() || (sv ? 'ditt konto' : 'this account');
    return sv ? `${who} får inte starta överföringen. Be en administratör ge dig rollen ` +
                `"Cloud Run Jobs Executor With Overrides" på jobbet ${job}.`
              : `${who} may not start the job; grant roles/run.jobsExecutorWithOverrides on ${job}`;
  }
  const editor = p.role === 'writer';
  return sv ? `"${p.name}" är inte delad med ${p.account} som ${editor ? 'redigerare' : 'läsare'}. ` +
              `Dela den så att jobbet kan ${editor ? 'skriva DMC-status' : 'läsa klippet'}.`
            : `"${p.name}" is not shared with ${p.account} as ${editor ? 'Editor' : 'Viewer'}`;
}
