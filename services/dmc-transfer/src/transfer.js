// Sends one clip from Google Drive to DMC's FTP server over explicit TLS (FTPS on port 21), and
// writes how it is going to the clip's row in the sheet. Drive, the sheet and the clock come in as
// arguments, so the tests can run this against a fake FTP server.
'use strict';
const ftp = require('basic-ftp');

/**
 * The job's settings from its environment variables (set when the job is deployed; FILE_IDS when
 * the sheet starts it). Throws naming every one that is missing.
 */
function configFromEnv(env) {
  const missing = ['FTP_HOST', 'FTP_USER', 'FTP_PASSWORD', 'SPREADSHEET_ID', 'FILE_IDS'].filter((k) => !env[k]);
  if (missing.length) throw new Error('Not set: ' + missing.join(', '));
  const suffix = env.FTP_TEMP_SUFFIX === undefined ? '.part' : env.FTP_TEMP_SUFFIX;
  return {
    ftp: {
      host: env.FTP_HOST,
      port: Number(env.FTP_PORT || 21),
      user: env.FTP_USER,
      password: env.FTP_PASSWORD,
      dir: env.FTP_DIR || '',
      // Uploaded under name + suffix, renamed when complete, so DMC never picks up half a file.
      // "none" uploads under the final name directly.
      tempSuffix: suffix === 'none' ? '' : suffix,
    },
    spreadsheetId: env.SPREADSHEET_ID,
    sheetName: env.SHEET_NAME || 'Video Assets',
    // The Asset Manager says where to write (recipients.js); these are its defaults.
    statusHeader: env.STATUS_HEADER || 'Recipient status',
    // Put before every status, so the shared status column says who it is about ("Playout: ").
    statusPrefix: env.STATUS_PREFIX || '',
    timeZone: env.TIME_ZONE || 'Europe/Stockholm',
  };
}

/**
 * The Drive file IDs this task sends. The sheet starts one execution with a task per clip; task i
 * takes every clip whose position in FILE_IDS is i modulo the task count, so all clips are sent
 * whatever the task count is.
 */
function taskFileIds(env) {
  const ids = String(env.FILE_IDS || '').split(',').map((s) => s.trim()).filter(Boolean);
  const index = Number(env.CLOUD_RUN_TASK_INDEX || 0);
  const count = Number(env.CLOUD_RUN_TASK_COUNT || 1);
  return ids.filter((_, i) => i % count === index);
}

/** The name a Drive file gets at DMC: its Drive name, with "/" and "\" (folders to FTP) as "-". */
function ftpName(driveName) {
  return String(driveName).replace(/[/\\]/g, '-').trim();
}

function formatBytes(n) {
  return n >= 1e9 ? (n / 1e9).toFixed(1) + ' GB' : Math.round(n / 1e6) + ' MB';
}

/** "Skickar 1.2 GB av 4.8 GB (25 %)" */
function progressText(bytes, total) {
  if (!(total > 0)) return 'Skickar ' + formatBytes(bytes);
  return `Skickar ${formatBytes(bytes)} av ${formatBytes(total)} (${Math.floor((bytes * 100) / total)} %)`;
}

/** "2026-11-17 14:32" in timeZone. */
function stamp(date, timeZone) {
  return new Intl.DateTimeFormat('sv-SE', {
    timeZone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit',
  }).format(date);
}

/**
 * Sends the Drive file fileId to the FTP server, overwriting a file with the same name there.
 *
 * deps:
 *   drive            { metadata(id) -> { name, size }, download(id) -> Readable }
 *   status(id, text) writes the clip's status to the sheet
 *   ftp              { host, port, user, password, dir, tempSuffix } (see configFromEnv)
 *   timeZone, now, log, progressEveryMs (how often to write progress, default 30 s),
 *   tlsOptions       extra options for the TLS connection (the tests trust their own certificate)
 *
 * Status: "Skickar", then progress, then "Skickad <time>"; or "Fel: <reason>" and the error is
 * thrown. The upload is checked by comparing the file's size at DMC with its size in Drive.
 */
async function sendClip(fileId, deps) {
  const f = deps.ftp;
  const log = deps.log || console.log;
  const now = deps.now || (() => new Date());
  const every = deps.progressEveryMs === undefined ? 30000 : deps.progressEveryMs;
  const setStatus = async (text) => {
    try {
      await deps.status(fileId, text);
    } catch (e) {
      log(`Could not write the status of ${fileId} ("${text}"): ${e.message}`);
    }
  };

  await setStatus('Skickar');
  const client = new ftp.Client(60000);
  let partial = null; // a file at DMC to remove if something fails
  try {
    const meta = await deps.drive.metadata(fileId);
    const name = ftpName(meta.name);
    await client.access({
      host: f.host, port: f.port, user: f.user, password: f.password,
      secure: true, // explicit TLS: AUTH TLS before the password is sent; fails if the server refuses
      secureOptions: Object.assign({ servername: f.host }, deps.tlsOptions),
    });
    if (f.dir) await client.cd(f.dir);

    const uploadName = f.tempSuffix ? name + f.tempSuffix : name;
    let lastProgress = Date.now();
    client.trackProgress((info) => {
      if (Date.now() - lastProgress < every) return;
      lastProgress = Date.now();
      setStatus(progressText(info.bytes, meta.size));
    });
    partial = uploadName;
    log(`Sending ${fileId} as ${f.dir ? f.dir + '/' : ''}${name} (${formatBytes(meta.size || 0)})`);
    await client.uploadFrom(await deps.drive.download(fileId), uploadName);
    client.trackProgress();

    const size = await client.size(uploadName);
    if (Number.isFinite(meta.size) && size !== meta.size) {
      // Shown to users in the status column, which names the step (Playout), not the company.
      throw new Error(`The recipient has ${size} bytes of ${name}, Drive has ${meta.size}`);
    }
    if (uploadName !== name) {
      try {
        await client.rename(uploadName, name);
      } catch (e) {
        // Some servers do not rename onto an existing file: remove the old one, then rename.
        await client.remove(name, true);
        await client.rename(uploadName, name);
      }
    }
    partial = null;
    await setStatus('Skickad ' + stamp(now(), deps.timeZone || 'Europe/Stockholm'));
    return { name, size };
  } catch (e) {
    client.trackProgress();
    if (partial && !client.closed) await client.remove(partial, true).catch(() => {});
    await setStatus('Fel: ' + e.message);
    throw e;
  } finally {
    client.close();
  }
}

module.exports = { configFromEnv, taskFileIds, ftpName, progressText, stamp, sendClip };
