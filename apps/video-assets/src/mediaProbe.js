// What the videos in "Video Inbox" are: duration, frame rate, size, codec, bitrate, scan, bit depth,
// alpha and audio.
// Apps Script cannot read video files, so a Cloud Run job does (services/media-probe in the repo,
// ffprobe). Whenever "Video Inbox" is rebuilt (by the asset watcher, or Inkorg → Uppdatera
// inkorgarna), the job is started for the videos whose "Probe status" is empty, and it writes each
// video's columns and "Probed <time>", or "Error: <reason>". The columns stay with their video when
// the tab is rebuilt. To probe a video again, clear its "Probe status".
//
// Script Properties: DMC_GCP_PROJECT and DMC_REGION (the job is in the same project as the DMC
// job), PROBE_JOB (optional, default media-probe).

// Columns E-O of "Video Inbox", after A-D (INBOX_HEADERS). The job finds them by these headers
// (RESULT_HEADERS in services/media-probe/src/probe.js, plus its status column).
const PROBE_HEADERS = ['Duration', 'Frame rate', 'Width', 'Height', 'Codec', 'Bitrate', 'Scan', 'Bit depth',
  'Alpha', 'Audio', 'Probe status'];

/** The Cloud Run job that probes videos (the `job` argument of the Cloud Run API in lib.js). */
function probeJob_() {
  return {
    project: requiredProperty_('DMC_GCP_PROJECT'),
    region: optionalProperty_('DMC_REGION', 'europe-north1'),
    name: optionalProperty_('PROBE_JOB', 'media-probe'),
  };
}

/** Whether ffprobe should look at a file of this type (Drive calls MXF application/mxf). */
function isProbeable_(mimeType) {
  return /^video\//.test(mimeType || '') || mimeType === 'application/mxf';
}

/**
 * Starts the probe job for the videos in files (rows of sheet, in order, as from listDriveFiles_)
 * whose "Probe status" (in statuses, same order) is empty, and marks them "Queued <time>".
 * Returns { started } (how many), or { started: 0, error } when the job could not be started; the
 * statuses are then left empty, so the next rebuild tries again.
 */
function startProbes_(sheet, files, statuses) {
  const statusCol = INBOX_HEADERS.length + PROBE_HEADERS.length; // A-D, then E-O: O
  const todo = [];
  files.forEach(function (f, i) {
    if (isProbeable_(f.mimeType) && !String(statuses[i] || '').trim()) todo.push(i);
  });
  if (!todo.length) return { started: 0 };
  try {
    runCloudRunJob_(probeJob_(), {
      env: { FILE_IDS: todo.map(function (i) { return files[i].id; }).join(','),
             REQUESTED_BY: Session.getEffectiveUser().getEmail() },
      taskCount: 1,
    });
  } catch (e) {
    console.log('The probe job did not start: ' + e.message);
    return { started: 0, error: e.message };
  }
  const queued = 'Queued ' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm');
  todo.forEach(function (i) { sheet.getRange(i + 2, statusCol).setValue(queued); });
  return { started: todo.length };
}
