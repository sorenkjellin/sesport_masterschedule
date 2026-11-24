// The master schedule app: menus, edit triggers and the settings shared across its files.
// The work itself is in the other files; lib.js (shared by all seSport apps) is added at deploy.

/**
 * seSport master schedule -> the "coming up" graphic in Elmo.
 *
 * The sheet holds the channel's schedule ("Template") and its images (a Drive folder). This script
 * uploads the images to an Elmo project and publishes the upcoming shows to Cachearoo (bucket
 * "sesport", key "comingup"), where the graphic (ComingUpWheel in the seSportPlayout project)
 * reads them. Nothing is published unless someone chooses Publish.
 *
 * Elmo menu:
 *   1. List images from Drive         lists the Drive folder's images (subfolders too) in the
 *                                     "Elmo images" sheet, with the path each gets in Elmo
 *   2. Upload changed images to Elmo  uploads the new and changed images; unchanged ones are skipped
 *   Publish coming up                 writes the next MAX_SHOWS (10) shows that have not started,
 *                                     soonest first, to Cachearoo
 *   Show coming up JSON               shows what Publish would write, without writing it
 *
 * A show's image: its "Episode image", else its "Series image", else its genre's image, the file
 * Genre/<Genre>.png in the Drive folder (images/Genre/Handboll.png in Elmo for the genre
 * Handboll). Image cells take a path in the Elmo project or a full URL.
 *
 * Warnings: an upcoming row that falls back to its genre's image while that image is not uploaded
 * gets a yellow "Genre image missing" in the "Warning" column (the graphic then shows its default
 * picture). The column follows each edit, Publish and upload.
 *
 * Script Properties (Project Settings -> Script Properties):
 *   DRIVE_FOLDER_ID    ID of the Drive folder with the images (from its URL)
 *   ELMO_URL           e.g. https://elmo.stinky.sesame-streams.com
 *   ELMO_KEY           write API key (the same one used for Cachearoo)
 *   ELMO_PROJECT       Elmo project to upload into
 *   ELMO_FOLDER        folder inside the project (optional, default "images")
 *   ELMO_CACHEAROO_URL where to write the schedule (ELMO_KEY is used for access)
 */

const TIME_ZONE = 'Europe/Stockholm';
const MASTER_SPREADSHEET_ID = '17kWFQnWMV-1AS1gw9DxL2JQgkRFD28gAOwCESiXAtFs';
const SCHEDULE_SHEET = 'Kalender'

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bilder')
    .addItem('Lista bilder i Google Drive', 'listDriveImages')
    .addItem('Ladda upp nya bilder till Elmo', 'uploadChangedImages')
    .addToUi();

  SpreadsheetApp.getUi().createMenu('Data')
    .addItem('Uppdatera data i coming-up-loopen', 'publishComingUp')
    .addItem("Maila EPG-data till Simply", "exportEPGAndEmail")
    .addSeparator()
    .addItem('Arkivera gamla rader', 'archiveOldCalendarRows')
    .addToUi();
}

// SIMPLE TRIGGER
// Keeps the Elmo schedule warnings up to date.
// This is the only function in the project named onEdit.
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === SCHEDULE_SHEET) {
    markMissingGenreImages();
  }
}

// INSTALLABLE EDIT TRIGGER
// Handles the Kalender, Playout and Program workflows that need authorization.
function handleEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  switch (sheet.getName()) {
    case 'Kalender':
      handleKalenderEdit_(e);
      break;
    case 'Playout':
      handlePlayoutEdit_(e);
      break;
    case PROGRAM_SHEET:
      handleProgramEdit_(e);
      break;
  }
}

function authorizeScript() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Test Sheets API authorization
  Sheets.Spreadsheets.get(ss.getId(), {
    fields: "spreadsheetId"
  });
  // Test Drive authorization
  DriveApp.getRootFolder().getName();
  // Test Gmail authorization required by exportEPGAndEmail.
  GmailApp.getAliases();
  Logger.log("Authorization successful");
}

function testAuthorization() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Test 1: Sheets API
  const result = Sheets.Spreadsheets.get(
    ss.getId(),{ fields: "spreadsheetId"});
  Logger.log("Sheets API: SUCCESS");
  Logger.log(result.spreadsheetId);
  // Test 2: Drive API
  const folder = DriveApp.getRootFolder();
  Logger.log("Drive API: SUCCESS");
  Logger.log(folder.getName());
}

function installEditTrigger() {
  // Replace any previous handleEdit triggers with one new installable edit trigger.
  deleteTriggers_('handleEdit');
  ensureEditTrigger_('handleEdit', SpreadsheetApp.getActiveSpreadsheet());
  Logger.log('Installable edit trigger created.');
}

function testTriggerAuthorization() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Program");
  try {
    const result = Sheets.Spreadsheets.get(
      ss.getId(),
      {
        ranges: ["Program!G2"],
        includeGridData: true,
        fields: "sheets(data(rowData(values(chipRuns))))"
      }
    );
    console.log("TRIGGER SHEETS API SUCCESS");
  } catch (err) {
    console.log("TRIGGER SHEETS API FAILED: " + err.message);
    throw err;
  }
}
