// The Cloud Run job's entry point. The sheet (Video → Skicka markerade till DMC) starts one
// execution with FILE_IDS set to the ticked clips and one task per clip; each task sends its clips
// one after the other and fails if any of them failed (Cloud Run then retries it, per the job's
// --max-retries). See docs/dmc-transfer.md for the settings.
'use strict';
const { configFromEnv, taskFileIds, sendClip } = require('./transfer');
const { googleApi } = require('./google');

async function main(env) {
  const cfg = configFromEnv(env);
  const ids = taskFileIds(env);
  const api = googleApi();
  const status = api.sheetStatus({
    spreadsheetId: cfg.spreadsheetId, sheetName: cfg.sheetName, header: cfg.statusHeader,
  });
  console.log(`Task ${env.CLOUD_RUN_TASK_INDEX || 0}: ${ids.length} clip(s), requested by ${env.REQUESTED_BY || 'unknown'}`);

  let failed = 0;
  for (const id of ids) {
    try {
      const sent = await sendClip(id, { drive: api.drive, status, ftp: cfg.ftp, timeZone: cfg.timeZone });
      console.log(`Sent ${id} as ${sent.name} (${sent.size} bytes)`);
    } catch (e) {
      failed++;
      console.error(`Failed to send ${id}: ${e.message}`);
    }
  }
  if (failed) process.exitCode = 1;
}

main(process.env).catch((e) => {
  console.error(e.message);
  process.exitCode = 1;
});
