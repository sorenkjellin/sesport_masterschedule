// The Video Assets app: the script bound to the "Video Assets" spreadsheet. It lists the videos in
// the Drive "Video" folder, sends the ones ticked to DMC (dmcTransfer.js) and emails about new
// files in the watched asset folder (assetWatcher.js), listing what is in Inbox/Video and
// Inbox/Images in their own tabs (inbox.js). lib.js (shared by all seSport apps) is
// added at deploy.
//
// The Apps Script project uses the Google Cloud project that holds the DMC transfer job, whose
// sign-in is limited to sesport.tv accounts. That is why this lives apart from the master schedule,
// which people outside sesport.tv use too.
//
// Script Properties (Project Settings -> Script Properties):
//   VIDEO_FOLDER_ID    ID of the Drive "Video" folder (from its URL)
//   DMC_GCP_PROJECT, DMC_REGION, DMC_JOB                   see dmcTransfer.js
//   PROBE_JOB                                              see mediaProbe.js
//   WATCH_FOLDER_ID, WATCH_RECIPIENTS, WATCH_SUBJECT_PREFIX, WATCH_RUN_AS   see assetWatcher.js

const TIME_ZONE = 'Europe/Stockholm';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Video')
    .addItem('Lista videor i Google Drive', 'listDriveVideos')
    .addSeparator()
    .addItem('Skicka markerade till DMC', 'sendCheckedToDmc')
    .addToUi();

  SpreadsheetApp.getUi().createMenu('Inkorg')
    .addItem('Uppdatera inkorgarna', 'refreshInboxLists')
    .addToUi();
}
