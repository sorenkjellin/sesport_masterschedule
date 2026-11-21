/**
 * Emails the files uploaded to a Drive folder (and its direct subfolders) since the last check.
 * First it refreshes the "Video Inbox" / "Image Inbox" tab when Inbox/Video or Inbox/Images got
 * new files (see inbox.js), and the mail links to the tabs it updated.
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
 * Lists the watched folder's subfolders that got new files in their inbox tabs, then emails the
 * files uploaded since the last check, with links to those tabs.
 * Run this function on a time-driven trigger (installWatcherTrigger).
 */
function checkForNewAssets() {
  if (!isWatcherAccount_()) return;
  watchDriveFolder_({
    folderId: requiredProperty_('WATCH_FOLDER_ID'),
    recipients: requiredProperty_('WATCH_RECIPIENTS'),
    subjectPrefix: optionalProperty_('WATCH_SUBJECT_PREFIX', '[Drive Alert] New Asset Uploaded'),
    stateKey: DRIVE_WATCH_STATE_KEY,
    intro: 'This is an automated message from the seSPORT Asset Manager',
    // First refresh the inbox tab of each subfolder that got new files (inbox.js); the mail then
    // links to the tabs that were updated, and says which could not be. The mail goes out either way.
    beforeMail: function (newFiles) {
      const folders = [];
      newFiles.forEach(function (f) { if (folders.indexOf(f.folderName) < 0) folders.push(f.folderName); });
      try {
        const results = listInboxes_(folders);
        results.forEach(function (r) { console.log(r.error ? `${r.sheet}: ${r.error}` : `${r.sheet}: ${r.count} files`); });
        return inboxMailNote_(results);
      } catch (e) {
        console.log('Inbox tabs not updated: ' + e.message);
        return 'The inbox tabs could not be updated: ' + e.message;
      }
    },
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

// How often the watcher runs. Apps Script allows 1, 5, 10, 15 or 30 minutes.
const WATCH_EVERY_MINUTES = 10;

/**
 * Run once from the editor, as the account in WATCH_RUN_AS: replaces your checkForNewAssets
 * triggers with one that runs every WATCH_EVERY_MINUTES minutes.
 */
function installWatcherTrigger() {
  deleteTriggers_('checkForNewAssets');
  ScriptApp.newTrigger('checkForNewAssets').timeBased().everyMinutes(WATCH_EVERY_MINUTES).create();
  Logger.log(`checkForNewAssets now runs every ${WATCH_EVERY_MINUTES} minutes as ${Session.getEffectiveUser().getEmail()}`);
}
