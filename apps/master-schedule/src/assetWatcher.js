/**
 * The asset watcher (mail about new files in the watched Drive folder) moved to the "Video Assets"
 * spreadsheet's script (apps/video-assets in the repo).
 *
 * This stays so that checkForNewAssets triggers still installed in this project do not fail with
 * "function not found": each one removes itself the next time it runs. (A script can only delete
 * the triggers of the account it runs as, so every account's trigger has to do this itself.)
 * Delete this file once the Triggers post-deploy check no longer mentions checkForNewAssets for
 * anyone.
 */
function checkForNewAssets() {
  deleteTriggers_('checkForNewAssets');
  console.log(`The asset watcher moved to the Video Assets spreadsheet: removed ` +
    `${Session.getEffectiveUser().getEmail()}'s checkForNewAssets trigger in the master schedule`);
}
