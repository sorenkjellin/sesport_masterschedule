/**
 * Configuration settings
 */
const DRIVE_WATCH_CONFIG = {
  // Replace with the ID of the parent Google Drive folder
  // (from: https://drive.google.com/drive/folders/YOUR_FOLDER_ID)
  PARENT_FOLDER_ID: '1T_ekwOZNoPo3LNAGV8xCbYCQpaZ46-5Z',

  // Recipient email address(es) - separate multiple with commas
  NOTIFICATION_EMAIL: 'soren@stinky.tv',

  // Subject line for the email notification
  EMAIL_SUBJECT_PREFIX: '[Drive Alert] New Asset Uploaded'
};

// The Script Property holding the time of the watcher's last check.
const DRIVE_WATCH_STATE_KEY = 'LAST_CHECK_TIME';

/**
 * Emails the files uploaded to the watched folder (and its subfolders) since the last check.
 * Run this function on a time-driven trigger (e.g., every 5 to 15 minutes).
 */
function checkForNewAssets() {
  watchDriveFolder_({
    folderId: DRIVE_WATCH_CONFIG.PARENT_FOLDER_ID,
    recipients: DRIVE_WATCH_CONFIG.NOTIFICATION_EMAIL,
    subjectPrefix: DRIVE_WATCH_CONFIG.EMAIL_SUBJECT_PREFIX,
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
