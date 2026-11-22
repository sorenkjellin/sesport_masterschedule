// Recipients: who a clip in Video Assets can be sent to. Each recipient is a choice in the Action
// column (G) and a Cloud Run job that delivers the clip the way that recipient wants it. All share
// one status column, "Recipient status" (H), holding the clip's latest send, each status starting
// with the recipient's name:
//   Playout: Köad <time> → Playout: Skickar … (with progress) → Playout: Skickad <time>,
//   or Playout: Fel: <reason>.
// A clip on its way to a recipient is not sent anywhere until that is done. The job is started with
// FILE_IDS (the clip), SHEET_NAME, STATUS_HEADER and STATUS_PREFIX (where and how to write), and
// REQUESTED_BY (who chose it, for its log).
//
// Today there is one: Playout, delivered by DMC (dmcTransfer.js). A new recipient gets a file of its
// own with its description (like dmcRecipient_) and its job, and an entry in recipients_().
//
// A recipient description:
//   choice         the Action dropdown's text, e.g. "Send to Playout"
//   label          its name, in messages and before its statuses, e.g. "Playout"
//   formerHeaders  columns that held only its statuses before (unprefixed); they carry over
//   job            function returning the Cloud Run job, { project, region, name }
//   checks         for the post-deploy tests: requiredEnv (the job's settings that must be set),
//                  secretEnv (a setting holding a secret the job's account must be able to read),
//                  deployScript, and describe(env) (where it delivers, for the log)

// Video Assets' Action column (G), and the recipients' shared status column after it (H).
const ACTION_COL = 7;
const RECIPIENT_STATUS_HEADER = 'Recipient status';

/** Every recipient, in the order of their choices and status columns. */
function recipients_() {
  return [dmcRecipient_()];
}

/** The recipient whose Action choice is choice, or null. */
function recipientByChoice_(choice) {
  return recipients_().filter(function (r) { return r.choice === choice; })[0] || null;
}

/** Video Assets' headers: the columns from Drive, Action, and Recipient status. */
function videoAssetHeaders_() {
  return VIDEO_DRIVE_HEADERS.concat(['Action', RECIPIENT_STATUS_HEADER]);
}

/** Whether a status ("Playout: Skickar …", or an older one without a name) is a send on its way. */
function isOnItsWay_(status) {
  return /^([^:]+: )?(Köad|Skickar)\b/.test(String(status));
}

/**
 * Sends the clip on row of Video Assets to recipient. Returns a message saying what happened; when
 * it cannot be sent, the reason is also written to its status (unless it is already on its way,
 * when that status is kept). requestedBy: who chose it, for the job's log.
 */
function sendToRecipient_(recipient, sheet, row, requestedBy) {
  const width = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  const statusCol = headers.indexOf(RECIPIENT_STATUS_HEADER) + 1;
  if (!statusCol) throw new Error(`ingen kolumn "${RECIPIENT_STATUS_HEADER}": välj Video → Lista videor i Google Drive`);
  const prefix = recipient.label + ': ';
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  const r = rows[row - 2] || [];
  const id = String(r[0] || '');
  const name = String(r[1] || '');
  if (!id) return 'Ingen video på raden';
  const status = String(r[statusCol - 1]);
  const setStatus = function (text) { sheet.getRange(row, statusCol).setValue(prefix + text); };

  // One status per clip: a send on its way (to any recipient) has to finish first.
  if (isOnItsWay_(status)) {
    return `"${name}" är redan på väg: "${status}". Töm ${RECIPIENT_STATUS_HEADER} om den har fastnat.`;
  }
  // The recipient gets every clip in the same place, so one with the same name would overwrite it.
  const twin = rows.findIndex(function (o, i) {
    const other = String(o[statusCol - 1]);
    return i !== row - 2 && String(o[1]) === name && isOnItsWay_(other) &&
      (other.indexOf(prefix) === 0 || !/^[^:]+: /.test(other)); // to this recipient, or from before names
  });
  if (twin >= 0) {
    const why = `en annan fil med samma namn (rad ${twin + 2}) är på väg, och de skulle skriva över varandra hos ` +
      recipient.label.toLowerCase();
    setStatus('Fel: ' + why);
    return `"${name}" skickades inte: ${why}`;
  }

  // Check the permissions first, so a missing share is reported here instead of in the job's log.
  const access = jobAccess_(recipient.job(), [{ id: id, name: name, role: 'reader' }]);
  if (access.problems.length) {
    const why = access.problems.map(function (p) { return jobProblemText_(p, 'sv', RECIPIENT_STATUS_HEADER); }).join(' ');
    setStatus('Fel: ' + why);
    return `"${name}" skickades inte: ${why}`;
  }

  setStatus('Köad ' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm'));
  try {
    runCloudRunJob_(recipient.job(), {
      env: { FILE_IDS: id, SHEET_NAME: sheet.getName(), STATUS_HEADER: RECIPIENT_STATUS_HEADER, STATUS_PREFIX: prefix,
             REQUESTED_BY: requestedBy || '' },
      taskCount: 1,
    });
  } catch (e) {
    setStatus('Fel: överföringen startade inte: ' + e.message);
    return `"${name}" skickades inte: ${e.message}`;
  }
  return `"${name}" skickas till ${recipient.label.toLowerCase()}. Följ den i ${RECIPIENT_STATUS_HEADER}.`;
}

// --- Permissions --------------------------------------------------------------
// A recipient's job runs as its own service account, which needs the spreadsheet (to write its
// status) and the clips (to read them) shared with it; the account starting it (the edit trigger's
// owner) needs permission to start the job. These are checked before anything is queued, and by the
// post-deploy tests. The probe job (mediaProbe.js) is checked the same way.

/**
 * What keeps a Cloud Run job (cfg: { project, region, name }) that works on this spreadsheet from
 * working for the user running this, as a list of problems (each with jobName: cfg.name):
 *   { kind: 'job', detail }                 the job cannot be read (missing, or no permission)
 *   { kind: 'run' }                         the user may not start the job
 *   { kind: 'share', name, role, account }  an item is not shared with the job's service account
 * items: [{ id, name, role }] Drive items the job needs ('reader' or 'writer'); the spreadsheet
 * (writer) is always checked. Items whose sharing the user cannot see are listed in `unchecked`.
 */
function jobAccess_(cfg, items) {
  const result = { account: '', problems: [], unchecked: [] };
  let job;
  try {
    job = getCloudRunJob_(cfg);
  } catch (e) {
    result.problems.push({ kind: 'job', detail: e.message, jobName: cfg.name });
    return result;
  }
  result.job = job;
  if (cloudRunJobPermissions_(cfg, ['run.jobs.runWithOverrides']).indexOf('run.jobs.runWithOverrides') < 0) {
    result.problems.push({ kind: 'run', jobName: cfg.name });
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
      result.problems.push({ kind: 'share', name: item.name, role: item.role, account: account, jobName: cfg.name });
    }
  });
  return result;
}

/**
 * A problem from jobAccess_ as a sentence, in Swedish ('sv', for the sheet) or English.
 * statusHeader: the column the job writes (for the Swedish sharing advice).
 */
function jobProblemText_(p, lang, statusHeader) {
  const sv = lang === 'sv';
  const job = p.jobName;
  if (p.kind === 'job') {
    return sv ? 'Kan inte läsa överföringsjobbet ' + job + ': ' + p.detail
              : 'cannot read the job ' + job + ': ' + p.detail;
  }
  if (p.kind === 'run') {
    // The account the job is started as: the edit trigger's owner when sending from the Action column.
    const who = Session.getEffectiveUser().getEmail() || (sv ? 'ditt konto' : 'this account');
    return sv ? `${who} får inte starta överföringen. Be en administratör ge dig rollen ` +
                `"Cloud Run Jobs Executor With Overrides" på jobbet ${job}.`
              : `${who} may not start the job; grant roles/run.jobsExecutorWithOverrides on ${job}`;
  }
  const editor = p.role === 'writer';
  return sv ? `"${p.name}" är inte delad med ${p.account} som ${editor ? 'redigerare' : 'läsare'}. ` +
              `Dela den så att jobbet kan ${editor ? 'skriva ' + (statusHeader || 'sin status') : 'läsa klippet'}.`
            : `"${p.name}" is not shared with ${p.account} as ${editor ? 'Editor' : 'Viewer'} (for ${job})`;
}
