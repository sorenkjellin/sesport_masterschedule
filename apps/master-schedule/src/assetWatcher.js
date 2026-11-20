/**
 * Emails the files uploaded to a Drive folder (and its direct subfolders) since the last check.
 *
 * Script Properties (Project Settings -> Script Properties):
 *   WATCH_FOLDER_ID       ID of the Drive folder to watch (from its URL)
 *   WATCH_RECIPIENTS      who gets the email; several addresses separated by commas
 *   WATCH_SUBJECT_PREFIX  start of the email subject (optional,
 *                         default "[Drive Alert] New Asset Uploaded")
 *   WATCH_RUN_AS          the one account whose trigger runs the watcher (optional; without it,
 *                         every user's trigger runs it). Anyone else's checkForNewAssets trigger
 *                         removes itself the first time it runs.
 */

// The Script Property holding the time of the watcher's last check.
const DRIVE_WATCH_STATE_KEY = 'LAST_CHECK_TIME';

/**
 * Emails the files uploaded to the watched folder (and its subfolders) since the last check.
 * Run this function on a time-driven trigger (e.g., every 5 to 15 minutes).
 */
function checkForNewAssets() {
  if (!isWatcherAccount_()) return;
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

/**
 * Whether this run's account (the trigger's owner) should run the watcher: true without
 * WATCH_RUN_AS, or when it is WATCH_RUN_AS. Otherwise deletes the account's own checkForNewAssets
 * triggers (a script can only delete the triggers of the account it runs as) and returns false.
 * An account whose address Google does not reveal runs the watcher as before.
 */
function isWatcherAccount_() {
  const runAs = optionalProperty_('WATCH_RUN_AS', '').trim();
  const me = Session.getEffectiveUser().getEmail();
  if (!runAs || !me || me.toLowerCase() === runAs.toLowerCase()) return true;
  deleteTriggers_('checkForNewAssets');
  console.log(`The watcher runs as ${runAs}, not ${me}: removed ${me}'s checkForNewAssets trigger`);
  return false;
}
