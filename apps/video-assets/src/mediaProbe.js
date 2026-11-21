// What the videos in "Video Inbox" are: duration, frame rate, size, codec, bitrate (Mbit), scan,
// bit depth, alpha and audio.
// Apps Script cannot read video files, so a Cloud Run job does (services/media-probe in the repo,
// ffprobe). Whenever "Video Inbox" is rebuilt (by the asset watcher, or Inkorg → Uppdatera
// inkorgarna), the job is started for the videos whose "Probe status" is empty, and it writes each
// video's columns and "Probed <time>", or "Error: <reason>". The columns stay with their video when
// the tab is rebuilt. To probe a video again, clear its "Probe status".
//
// Script Properties: DMC_GCP_PROJECT and DMC_REGION (the job is in the same project as the DMC
// job), PROBE_JOB (optional, default media-probe).

// The status a file that is not a video gets instead of being probed.
const NOT_A_VIDEO = 'Not a video';

// Columns F-P of "Video Inbox", after A-E (INBOX_HEADERS). The job finds them by these headers
// (RESULT_HEADERS in services/media-probe/src/probe.js, plus its status column).
const PROBE_HEADERS = ['Duration', 'Frame rate', 'Width', 'Height', 'Codec', 'Mbit', 'Scan', 'Bit depth',
  'Alpha', 'Audio', 'Probe status'];

/** The Cloud Run job that probes videos (the `job` argument of the Cloud Run API in lib.js). */
function probeJob_() {
  return {
    project: requiredProperty_('DMC_GCP_PROJECT'),
    region: optionalProperty_('DMC_REGION', 'europe-north1'),
    name: optionalProperty_('PROBE_JOB', 'media-probe'),
  };
}

/** Whether ffprobe should look at a file of this type: any video (isVideoFile_ in videos.js). */
function isProbeable_(mimeType) {
  return isVideoFile_(mimeType);
}

/**
 * Starts the probe job for the videos in files (rows of sheet, in order, as from listDriveFiles_)
 * whose "Probe status" (in statuses, same order) is empty, and marks them "Queued <time>".
 * Returns { started } (how many), or { started: 0, error } when the job could not be started; the
 * statuses are then left empty, so the next rebuild tries again.
 */
function startProbes_(sheet, files, statuses) {
  const statusCol = INBOX_HEADERS.length + PROBE_HEADERS.length; // A-E, then F-P: P
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

// --- Results stored on the file (Drive properties) --------------------------------------------------
// A video's probe results go with it as Drive properties: the probe job stores them on the file
// whenever it probes it, the Action column writes them again with the approved profiles when it
// moves the video (actions.js), and Video Assets reads them (videos.js), so videos in the production
// folders need not be probed again. The keys are the probe job's
// (services/media-probe/src/properties.js; a test keeps them the same).

const PROBE_PROPERTY_KEYS = {
  Duration: 'sesport.duration', // whole seconds
  'Frame rate': 'sesport.frameRate',
  Width: 'sesport.width',
  Height: 'sesport.height',
  Codec: 'sesport.codec',
  Mbit: 'sesport.mbit',
  Scan: 'sesport.scan',
  'Bit depth': 'sesport.bitDepth',
  Alpha: 'sesport.alpha',
  Audio: 'sesport.audio',
};
const PROBED_PROPERTY = 'sesport.probed';
const APPROVED_PROPERTY = 'sesport.approvedFor';
// How many videos one run of "Analysera videor utan metadata" sends (the job's FILE_IDS has to fit
// in an environment variable); run it again for the rest.
const MAX_ASSET_PROBES = 500;

/**
 * The Drive properties for the video on row of Video Inbox: its probe results, when they were
 * probed, and the profiles it is approved for (from the note on its name). {} when not probed.
 */
function inboxRowProperties_(sheet, row) {
  const width = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function (h) { return String(h).trim(); });
  const cells = readCells_(sheet.getRange(row, 1, 1, width))[0];
  const status = cells[headers.indexOf('Probe status')];
  if (!status || !/^Probed/.test(status.text)) return {};
  const props = {};
  Object.keys(PROBE_PROPERTY_KEYS).forEach(function (h) {
    const cell = cells[headers.indexOf(h)];
    if (cell && cell.text !== '') props[PROBE_PROPERTY_KEYS[h]] = String(cell.value).slice(0, 100);
  });
  props[PROBED_PROPERTY] = status.text.replace(/^Probed\s*/, '');
  const note = sheet.getRange(row, headers.indexOf('Name') + 1).getNote();
  const approved = /^Approved for:/.test(note) ? note.split('\n').slice(1).filter(Boolean) : [];
  if (approved.length) props[APPROVED_PROPERTY] = approved.join(', ').slice(0, 100);
  return props;
}

/**
 * Menu: Video → Analysera videor utan metadata. Sends the videos in the Video folder that have no
 * stored results (MAX_ASSET_PROBES at a time) to the probe job, which stores them on the files.
 * The next Video → Lista videor shows them.
 */
function probeAssetsWithoutMetadata() {
  const folder = DriveApp.getFolderById(requiredProperty_('VIDEO_FOLDER_ID'));
  const props = driveFolderProperties_(folder);
  const missing = listDriveFiles_(folder, '').filter(function (f) {
    return isVideoFile_(f.mimeType) && !(props[f.id] || {})[PROBE_PROPERTY_KEYS.Duration];
  });
  if (!missing.length) { toast_('Alla videor har metadata', 'Video'); return { started: 0, left: 0 }; }
  const batch = missing.slice(0, MAX_ASSET_PROBES);
  runCloudRunJob_(probeJob_(), {
    env: { FILE_IDS: batch.map(function (f) { return f.id; }).join(','), TARGET: 'drive',
           REQUESTED_BY: Session.getEffectiveUser().getEmail() },
    taskCount: Math.min(4, Math.ceil(batch.length / 100)), // about 100 videos per task, 4 at a time
  });
  const left = missing.length - batch.length;
  toast_(`${batch.length} videor analyseras. Välj Video → Lista videor när de är klara` +
    (left ? `; ${left} till: kör det här igen sedan.` : '.'), 'Video', 10);
  return { started: batch.length, left: left };
}
