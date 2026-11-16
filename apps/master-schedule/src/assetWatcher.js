/**
 * Emails the files uploaded to a Drive folder (and its direct subfolders) since the last check.
 *
 * Script Properties (Project Settings -> Script Properties):
 *   WATCH_FOLDER_ID       ID of the Drive folder to watch (from its URL)
 *   WATCH_RECIPIENTS      who gets the email; several addresses separated by commas
 *   WATCH_SUBJECT_PREFIX  start of the email subject (optional,
 *                         default "[Drive Alert] New Asset Uploaded")
 */

// The Script Property holding the time of the watcher's last check.
const DRIVE_WATCH_STATE_KEY = 'LAST_CHECK_TIME';

/**
 * Emails the files uploaded to the watched folder (and its subfolders) since the last check.
 * Run this function on a time-driven trigger (e.g., every 5 to 15 minutes).
 */
function checkForNewAssets() {
  watchDriveFolder_({
    folderId: requiredProperty_('WATCH_FOLDER_ID'),
    recipients: requiredProperty_('WATCH_RECIPIENTS'),
    subjectPrefix: optionalProperty_('WATCH_SUBJECT_PREFIX', '[Drive Alert] New Asset Uploaded'),
    stateKey: DRIVE_WATCH_STATE_KEY,
  });
}

/**
 * Reset checkpoint if you want to force a re-check of the last hour.
 */
function resetCheckpoint() {
  PropertiesService.getScriptProperties().deleteProperty(DRIVE_WATCH_STATE_KEY);
  Logger.log('Checkpoint reset. The next run will scan the past hour.');
}
