// The Cloud Run job's entry point. The Asset Manager starts an execution with FILE_IDS set to the
// videos in "Video Inbox" that have not been probed. Each video's results go to its row, and its
// "Probe status" says when it was probed, or why it could not be. The job fails if any video failed.
//
// Settings (environment variables): SPREADSHEET_ID (set when deploying), FILE_IDS (set per run),
// SHEET_NAME (default "Video Inbox"), TIME_ZONE (default Europe/Stockholm).
'use strict';
const { probeFile } = require('./probe');
const { googleApi, driveDownloadUrl } = require('./google');

const STATUS_HEADER = 'Probe status';

/** "2026-11-21 17:02" in timeZone. */
function stamp(date, timeZone) {
  return new Intl.DateTimeFormat('sv-SE', {
    timeZone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit',
  }).format(date);
}

/** The Drive file IDs this task probes: task i takes every ID at a position that is i modulo the count. */
function taskFileIds(env) {
  const ids = String(env.FILE_IDS || '').split(',').map((s) => s.trim()).filter(Boolean);
  const index = Number(env.CLOUD_RUN_TASK_INDEX || 0);
  const count = Number(env.CLOUD_RUN_TASK_COUNT || 1);
  return ids.filter((_, i) => i % count === index);
}

/**
 * Probes ids one after the other and writes each result; returns how many failed.
 * deps: api (googleApi()), sheet { spreadsheetId, sheetName }, timeZone, run, now.
 */
async function probeAll(ids, { api, sheet, timeZone = 'Europe/Stockholm', run, now = () => new Date() }) {
  const write = api.sheetRow(sheet);
  let failed = 0;
  for (const id of ids) {
    try {
      await write(id, { [STATUS_HEADER]: 'Probing' });
      const { values } = await probeFile({ url: driveDownloadUrl(id), token: await api.token(), run });
      await write(id, Object.assign({}, values, { [STATUS_HEADER]: 'Probed ' + stamp(now(), timeZone) }));
      console.log(`Probed ${id}: ${Object.values(values).join(' | ')}`);
    } catch (e) {
      failed++;
      console.error(`Failed to probe ${id}: ${e.message}`);
      try {
        await write(id, { [STATUS_HEADER]: 'Error: ' + e.message.slice(0, 300) });
      } catch (e2) {
        console.error(`Could not write the error for ${id}: ${e2.message}`);
      }
    }
  }
  return failed;
}

async function main(env) {
  const missing = ['SPREADSHEET_ID', 'FILE_IDS'].filter((k) => !env[k]);
  if (missing.length) throw new Error('Not set: ' + missing.join(', '));
  const ids = taskFileIds(env);
  console.log(`Task ${env.CLOUD_RUN_TASK_INDEX || 0}: ${ids.length} video(s), requested by ${env.REQUESTED_BY || 'unknown'}`);
  const failed = await probeAll(ids, {
    api: googleApi(),
    sheet: { spreadsheetId: env.SPREADSHEET_ID, sheetName: env.SHEET_NAME || 'Video Inbox' },
    timeZone: env.TIME_ZONE || 'Europe/Stockholm',
  });
  if (failed) process.exitCode = 1;
}

if (require.main === module) {
  main(process.env).catch((e) => {
    console.error(e.message);
    process.exitCode = 1;
  });
}

module.exports = { STATUS_HEADER, stamp, taskFileIds, probeAll, main };
