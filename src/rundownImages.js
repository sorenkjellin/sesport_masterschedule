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
 *   ELMO_VIDEO_FOLDER  folder inside the project for videos (optional, default "videos")
 *   ELMO_CACHEAROO_URL where to write the schedule (ELMO_KEY is used for access)
 */

const BUCKET = 'sesport';
const KEY = 'comingup';
const TIME_ZONE = 'Europe/Stockholm';
const MASTER_SPREADSHEET_ID = '17kWFQnWMV-1AS1gw9DxL2JQgkRFD28gAOwCESiXAtFs';
const SCHEDULE_SHEET = 'Kalender'
// at most this many upcoming shows are published, soonest first
const MAX_SHOWS = 10;
// a show's image is its episode image, else its series image, else its genre's: Genre/<genre>.png
// in the Drive folder, uploaded under ELMO_FOLDER ("images/Genre/Handboll.png" for the genre Handboll)
const GENRE_FOLDER = 'Genre';
const GENRE_EXT = '.png';
// the schedule column where a row is warned about a missing genre image, and the warning's fill
const WARNING_HEADER = 'Warning';
const WARNING_COLOR = '#fff2cc';

const SHEET_NAME = 'Elmo images';
const HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Elmo path',
                 'Elmo URL', 'Uploaded hash', 'Uploaded modified', 'Status'];
const COL = { id: 0, name: 1, modified: 2, size: 3, path: 4, url: 5, hash: 6, upMod: 7, status: 8 };
const MISSING = 'Missing in Drive';
// Apps Script stops a run at 6 minutes; leave room to write the last row.
// Rows not reached are picked up by the next run.
const MAX_RUNTIME_MS = 5 * 60 * 1000;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Bilder')
    .addItem('Lista bilder i Google Drive', 'listDriveImages')
    .addItem('Ladda upp nya bilder till Elmo', 'uploadChangedImages')
    .addToUi();

  SpreadsheetApp.getUi().createMenu('Video')
    .addItem('Lista videor i Google Drive', 'listDriveVideos')
    .addToUi();

  SpreadsheetApp.getUi().createMenu('Data')
    .addItem('Uppdatera data i coming-up-loopen', 'publishComingUp')
    .addItem("Maila EPG-data till Simply", "exportEPGAndEmail")
    .addSeparator()
    .addItem('Arkivera gamla rader', 'archiveOldCalendarRows')
    .addToUi();
}

// --- Step 1: Drive folder -> sheet ------------------------------------------

function listDriveImages() {
  const cfg = config_();
  const sheet = getSheet_();

  const existing = {};
  readTextRows_(sheet, HEADERS.length).forEach(function (r) { if (r[COL.id]) existing[r[COL.id]] = r; });

  const found = listDriveFiles_(DriveApp.getFolderById(cfg.folderId), 'image/');
  found.sort(function (a, b) { return a.relPath.localeCompare(b.relPath); });

  const seen = {};
  const rows = found.map(function (f) {
    seen[f.id] = true;
    const old = existing[f.id];
    return [
      f.id, f.name, f.modified, String(f.size), cfg.folder + '/' + f.relPath,
      old ? old[COL.url] : '',
      old ? old[COL.hash] : '',
      old ? old[COL.upMod] : '',
      old && old[COL.status] !== MISSING ? old[COL.status] : 'New',
    ];
  });

  // Keep rows whose Drive file is gone, flagged, so their Elmo URLs stay visible.
  Object.keys(existing).forEach(function (id) {
    if (seen[id]) return;
    const r = existing[id].slice();
    r[COL.status] = MISSING;
    rows.push(r);
  });

  writeTextRows_(sheet, rows, HEADERS.length);
  toast_('Listed ' + found.length + ' images from Drive');
}

// --- Videos: Drive folder -> "Video Assets" columns A-D ----------------------
// The "Video" folder is the sibling of the images folder (DRIVE_FOLDER_ID). Lists every video in
// it and its subfolders with the same details as "Elmo images" columns A-D.
// Like "Elmo images", the list is rebuilt on every run, sorted by folder path and name.
// Only columns A-F are cleared and written.
const VIDEO_SHEET_NAME = 'Video Assets';
const VIDEO_FOLDER_NAME = 'Video';
// Drive file ID, Name, Drive modified, Size, Elmo path, Duration
const VIDEO_HEADERS = HEADERS.slice(0, 5).concat(['Duration']);

function listDriveVideos() {
  const cfg = config_();
  const imagesFolder = DriveApp.getFolderById(cfg.folderId);
  const videoFolder = findSiblingFolder_(imagesFolder, VIDEO_FOLDER_NAME);
  if (!videoFolder) {
    throw new Error('Hittade ingen mapp "' + VIDEO_FOLDER_NAME + '" bredvid "' + imagesFolder.getName() + '"');
  }

  const sheet = getOrCreateSheet_(VIDEO_SHEET_NAME, VIDEO_HEADERS);
  const found = listDriveFiles_(videoFolder, 'video/');
  found.sort(function (a, b) { return a.relPath.localeCompare(b.relPath); });
  // The folder inside the Elmo project: ELMO_VIDEO_FOLDER, default "videos".
  const elmoFolder = trimSlashes_(optionalProperty_('ELMO_VIDEO_FOLDER', 'videos'));
  // Duration as H:MM:SS:FF (25 fps); '' for a video Drive has not processed or cannot read.
  const durations = driveVideoDurationsMs_(found.map(function (f) { return f.id; }));
  const rows = found.map(function (f) {
    const ms = durations[f.id];
    return [f.id, f.name, f.modified, String(f.size), elmoFolder + '/' + f.relPath,
            ms == null ? '' : formatTimecode_(ms, 25)];
  });

  // Plain text, as in "Elmo images", so the timestamps stay as written.
  writeTextRows_(sheet, rows, VIDEO_HEADERS.length);
  toast_('Listed ' + found.length + ' videos from Drive');
}

// --- Step 2: sheet -> Elmo, only what changed ------------------------------
function uploadChangedImages() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { toast_('Another upload is already running'); return; }
  const started = Date.now();

  try {
    const cfg = config_();
    const sheet = getSheet_();
    const rows = readTextRows_(sheet, HEADERS.length);
    const elmo = getElmoFiles_(cfg);   // path -> { hash, url }
    ensureElmoProject_(cfg);

    const pathCount = {};
    rows.forEach(function (r) {
      if (r[COL.id] && r[COL.status] !== MISSING) pathCount[r[COL.path]] = (pathCount[r[COL.path]] || 0) + 1;
    });

    const stats = { uploaded: 0, unchanged: 0, failed: 0, remaining: 0 };

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r[COL.id] || r[COL.status] === MISSING) continue;
      if (Date.now() - started > MAX_RUNTIME_MS) { stats.remaining++; continue; }

      const path = r[COL.path];
      const before = r.join('\u0000');
      try {
        if (pathCount[path] > 1) {
          r[COL.status] = 'Duplicate Elmo path, rename in Drive';
          stats.failed++;
        } else {
          const file = DriveApp.getFileById(r[COL.id]);
          const modified = file.getLastUpdated().toISOString();
          const inElmo = elmo[path];

          if (r[COL.hash] && r[COL.upMod] === modified && inElmo && inElmo.hash === r[COL.hash]) {
            // Fast path: untouched in Drive since our last upload, and Elmo
            // still has that version. No download needed.
            r[COL.url] = inElmo.url;
            r[COL.status] = 'Up to date';
            stats.unchanged++;
          } else {
            const blob = file.getBlob();
            const hash = sha256Hex_(blob.getBytes());
            if (inElmo && inElmo.hash === hash) {
              r[COL.url] = inElmo.url;
              r[COL.status] = 'Up to date';
              stats.unchanged++;
            } else {
              r[COL.url] = uploadToElmo_(cfg, blob, path);
              r[COL.status] = 'Uploaded ' + new Date().toISOString();
              elmo[path] = { hash: hash, url: r[COL.url] };
              stats.uploaded++;
            }
            r[COL.hash] = hash;
          }
          r[COL.modified] = modified;
          r[COL.upMod] = modified;
        }
      } catch (e) {
        r[COL.status] = 'Error: ' + e.message;
        stats.failed++;
      }

      // Write each changed row as we go so progress survives a timeout.
      if (r.join('\u0000') !== before) writeTextRow_(sheet, i, r);
    }

    markMissingGenreImages();
    toast_('Uploaded ' + stats.uploaded + ', unchanged ' + stats.unchanged +
           ', failed ' + stats.failed +
           (stats.remaining ? ', ' + stats.remaining + ' left for next run' : ''));
  } finally {
    lock.releaseLock();
  }
}

// --- Config -----------------------------------------------------------------

/** The Elmo settings from Script Properties; also the `elmo` argument of the Elmo API in lib.js. */
function config_() {
  return {
    folderId: requiredProperty_('DRIVE_FOLDER_ID'),
    elmoUrl: requiredProperty_('ELMO_URL').replace(/\/+$/, ''),
    key: requiredProperty_('ELMO_KEY'),
    project: requiredProperty_('ELMO_PROJECT'),
    folder: trimSlashes_(optionalProperty_('ELMO_FOLDER', 'images')),
  };
}

function getSheet_() {
  return getOrCreateSheet_(SHEET_NAME, HEADERS);
}

// Writes the schedule to Cachearoo.
function publishComingUp() {
  const base = requiredProperty_('ELMO_CACHEAROO_URL');
  const apiKey = requiredProperty_('ELMO_KEY');
  const project = requiredProperty_('ELMO_PROJECT');
  const value = buildComingUp(readSchedule(), uploadedImages(), project, new Date());
  markMissingGenreImages();
  cachearooWrite_(base, apiKey, BUCKET, KEY, value);
  toast_(`${value.shows.length} shows published`);
}

// The JSON that would be written, without writing it.
function showComingUpJson() {
  const project = optionalProperty_('ELMO_PROJECT', '');
  showTextDialog_('Coming up',
    JSON.stringify(buildComingUp(readSchedule(), uploadedImages(), project, new Date()), null, 2));
}

// The schedule columns the graphic uses. A header matches ignoring case and surrounding spaces;
// the first of a column's names that the sheet has is used.
const COLUMNS = {
  date: ['Datum'],
  start: ['Börjar'],
  end: ['Slutar'],
  genre: ['Genre'],
  league: ['Ligans namn'],
  series: ['Programtitel', 'Progamtitel'],
  episode: ['Avsnitt'],
  live: ['Live/Rerun'],
  description: ['Avsnittsbeskrivning'],
  seriesImage: ['Programbild'],
  episodeImage: ['Avsnittsbild'],
};

// The schedule sheet: its rows (as displayed), its lowercased headers and each column's index
// (-1 when the sheet does not have it).
function readSchedule() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SCHEDULE_SHEET);
  if (!sheet) throw new Error(`There is no sheet named "${SCHEDULE_SHEET}"`);
  const [headers, ...rows] = sheet.getDataRange().getDisplayValues();
  const names = headers.map((h) => String(h).trim().toLowerCase());
  const cols = findColumns_(headers, COLUMNS);
  ['date', 'start'].forEach((key) => {
    if (cols[key] < 0) throw new Error(`The sheet "${SCHEDULE_SHEET}" has no column "${COLUMNS[key][0]}"`);
  });
  return { sheet, names, rows, cols };
}

function field(schedule, row, key) {
  const i = schedule.cols[key];
  return i >= 0 ? String(row[i] ?? '').trim() : '';
}

// The shows not yet started, soonest first, at most MAX_SHOWS.
function buildComingUp(schedule, uploaded, project, now) {
  const shows = [];
  for (const row of schedule.rows) {
    const date = field(schedule, row, 'date');
    const startTime = field(schedule, row, 'start');
    const start = parseLocal(date, startTime);
    if (!start || start <= now) continue;
    const endTime = field(schedule, row, 'end');
    let end = endTime ? parseLocal(date, endTime) : null;
    // a show that ends after midnight
    if (end && end <= start) end = new Date(end.getTime() + 24 * 3600 * 1000);

    const title = field(schedule, row, 'episode') || field(schedule, row, 'series');
    shows.push({
      id: `${date} ${startTime} ${title}`,
      start: start.toISOString(),
      end: end ? end.toISOString() : null,
      sport: field(schedule, row, 'genre'),
      league: field(schedule, row, 'league'),
      title,
      description: field(schedule, row, 'description'),
      image: elmoFileUrl_(pickImage(schedule, row, uploaded).image, project),
      live: field(schedule, row, 'live').toLowerCase() === 'live',
    });
  }
  shows.sort((a, b) => (a.start < b.start ? -1 : a.start > b.start ? 1 : 0));
  return { shows: shows.slice(0, MAX_SHOWS) };
}

// "2026-11-07" or "7.11.2026", and "17.50" or "17:50", in Stockholm time
function parseLocal(date, time) {
  return parseLocalDateTime_(date, time, TIME_ZONE);
}

// The images uploaded to Elmo, by the image sheet: byPath[elmo path] is true; byName[file name]
// (lowercased, with and without its extension) is its Elmo path, or null when uploaded images in
// different folders share the name.
function uploadedImages() {
  const byPath = {};
  const byName = {};
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (sheet) {
    readTextRows_(sheet, HEADERS.length).forEach(function (r) {
      if (!/^(Uploaded|Up to date)/.test(r[COL.status])) return;
      const path = r[COL.path];
      byPath[path] = true;
      [r[COL.name], path.split('/').pop()].forEach(function (name) {
        const full = String(name).trim().toLowerCase();
        [full, full.replace(/\.[^.]+$/, '')].forEach(function (key) {
          if (key) byName[key] = key in byName && byName[key] !== path ? null : path;
        });
      });
    });
  }
  return { byPath, byName };
}

// An image cell as a URL or an uploaded image's Elmo path; '' when it is neither.
function resolveImage(value, uploaded) {
  if (/^https?:\/\//i.test(value)) return value;
  const path = value.replace(/^\/+/, '');
  if (uploaded.byPath[path]) return path;
  return path.indexOf('/') < 0 ? uploaded.byName[path.toLowerCase()] || '' : '';
}

// A row's image (an Elmo path or a URL, '' for none): its episode image, else its series image,
// else its genre's, the first that is in Elmo; and a warning for each of them that is not.
function pickImage(schedule, row, uploaded) {
  const warnings = [];
  let image = '';
  [['Avsnittsbild', 'episodeImage'], ['Programbild', 'seriesImage']].forEach(function ([label, key]) {
    const value = field(schedule, row, key);
    if (!value) return;
    const found = resolveImage(value, uploaded);
    if (!found) warnings.push(`${label} not in Elmo: ${value}`);
    else if (!image) image = found;
  });
  const genre = genreImage(field(schedule, row, 'genre'));
  if (!image && genre) {
    if (uploaded.byPath[genre]) image = genre;
    else warnings.push(`Genre image missing: ${genre}`);
  }
  return { image, warnings };
}

// Writes a warning beside each upcoming schedule row whose image cells are not in Elmo, or that
// falls back to a genre image that is not uploaded (by the image sheet); clears the others.
function markMissingGenreImages() {
  const schedule = readSchedule();
  const sheet = schedule.sheet;
  let col = schedule.names.indexOf(WARNING_HEADER.toLowerCase());
  if (col < 0) {
    col = schedule.names.length;
    sheet.getRange(1, col + 1).setValue(WARNING_HEADER).setFontWeight('bold');
  }
  if (!schedule.rows.length) return;
  const uploaded = uploadedImages();
  const now = new Date();
  const warnings = schedule.rows.map(function (row) {
    const start = parseLocal(field(schedule, row, 'date'), field(schedule, row, 'start'));
    return start && start > now ? pickImage(schedule, row, uploaded).warnings.join('; ') : '';
  });
  sheet.getRange(2, col + 1, warnings.length, 1)
    .setValues(warnings.map((w) => [w]))
    .setBackgrounds(warnings.map((w) => [w ? WARNING_COLOR : null]));
}

// The genre's image path in the project, as the upload names it ("/" in a name becomes "-")
function genreImage(genre) {
  if (!genre) return '';
  const folder = trimSlashes_(optionalProperty_('ELMO_FOLDER', 'images'));
  return `${folder}/${GENRE_FOLDER}/${genre.replace(/\//g, '-')}${GENRE_EXT}`;
}

/**
 * Helper function to configure validation for a specific row.
 */
function updateDropdownForRow(sheet, row) {
  const genre = sheet.getRange(row, 4).getValue().toString().trim();
  const league = sheet.getRange(row, 5).getValue().toString().trim();

  const hemmaCell = sheet.getRange(row, 6);
  const bortaCell = sheet.getRange(row, 7);

  // If either genre or league is missing, clear validations
  if (!genre || !league) {
    hemmaCell.clearDataValidations();
    bortaCell.clearDataValidations();
    return;
  }

  const targetHeader = genre + ': ' + league;
  const ss = sheet.getParent();
  const ligorSheet = ss.getSheetByName('Ligor&Lag');
  if (!ligorSheet) return;

  const lastCol = ligorSheet.getLastColumn();
  const headers = ligorSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const matchIndex = headers.findIndex(h => h.toString().trim() === targetHeader);

  if (matchIndex === -1) {
    hemmaCell.clearDataValidations();
    bortaCell.clearDataValidations();
    return;
  }

  const colNum = matchIndex + 1;
  const colValues = ligorSheet.getRange(2, colNum, ligorSheet.getLastRow() - 1, 1).getValues();

  // Filter out blank rows
  const teams = colValues
    .map(r => r[0])
    .filter(t => t !== null && t !== undefined && t.toString().trim() !== '');

  if (teams.length === 0) {
    hemmaCell.clearDataValidations();
    bortaCell.clearDataValidations();
    return;
  }

  // Build chip dropdown validation rule
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(teams, true)
    .setAllowInvalid(false)
    .build();

  hemmaCell.setDataValidation(rule);
  bortaCell.setDataValidation(rule);
}

/**
 * Exports the "EPG" tab values to an XLSX file and emails it.
 * Script Properties (Project Settings -> Script Properties):
 *   EPG_MAIL_ENDPOINT  email address where to send expeorted data from the EPG tab
*/
function exportEPGAndEmail() {
  const recipient = requiredProperty_('EPG_MAIL_ENDPOINT');
  const cc = optionalProperty_('EPG_MAIL_CC', '');
  const sheetName = 'EPG';
  const fileName = 'seSport_EPG.xlsx';
  const DROPPED_COLUMN = 15; // column P (0-based) is not sent

  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${sheetName}" was not found!`);
    return;
  }

  // Display values, so dates and times are sent as shown and no formulas are carried over.
  const values = sourceSheet.getDataRange().getDisplayValues()
    .map((row) => row.filter((_, i) => i !== DROPPED_COLUMN));
  const blob = valuesToXlsx_(values, sheetName, fileName);

  const subject = `seSport EPG Export for Simply (${new Date().toLocaleDateString()})`;
  const body = `Hi,\n\nPlease find attached the updated seSport programming schedule. \n\nBest Regards \nteam seSport`;
  sendMail_(recipient, subject, body, {
    attachments: [blob],
    cc: cc,
    name: 'Master Schedule Exporter'
  });

  // Notify user
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Export sent successfully to ${recipient} (cc ${cc})`,
    'Done',
    5
  );
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
// Handles the Kalender and Playout workflows that need authorization.
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
  }
}

function ensureKorSchemaEditTrigger_(spreadsheet) {
  ensureEditTrigger_('syncKorSchemaValueToKalender_', spreadsheet);
}

function syncKorSchemaValueToKalender_(e) {
  if (!e || !e.range) return;

  const scheduleSheet = e.range.getSheet();
  const value = scheduleSheet.getRange('F2').getValue();
  if (value == scheduleSheet.getRange('J1').getValue()) return;
  scheduleSheet.getRange('J1').setValue(value);

  const id = String(scheduleSheet.getRange('J2').getValue()).trim();
  if (!id) return;

  const masterSpreadsheet = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const kalenderSheet = masterSpreadsheet.getSheetByName('Kalender');
  if (!kalenderSheet) {
    throw new Error('Sheet "Kalender" was not found in the master spreadsheet.');
  }

  const kalenderRow = findKalenderRowById_(kalenderSheet, id);
  if (!kalenderRow) {
    Logger.log('No Kalender row has the körschema ID ' + id + ' in column S.');
    return;
  }
  kalenderSheet.getRange(kalenderRow, 3).setValue(value);
}

// Kalender column S holds an ID per row that has a körschema; the körschema holds the same ID in J2.
// The ID, not the row number, links the two, so rows can be moved or sorted in Kalender.
const KALENDER_ID_COL = 19; // Column S
const KALENDER_ID_HEADER = 'Körschema-ID';

function newKalenderId_(sheet, row) {
  const header = sheet.getRange(1, KALENDER_ID_COL);
  if (header.getValue() === '') header.setValue(KALENDER_ID_HEADER);
  const id = Utilities.getUuid();
  sheet.getRange(row, KALENDER_ID_COL).setValue(id);
  return id;
}

// The Kalender row (1-based) whose column S holds the ID, or 0.
function findKalenderRowById_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const ids = sheet.getRange(2, KALENDER_ID_COL, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) return i + 2;
  }
  return 0;
}

// Kalender edit handling
function handleKalenderEdit_(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  // Start time edited (also when pasted over several rows): update each row's körschema.
  if (col <= 2 && e.range.getLastColumn() >= 2) {
    for (let r = row; r <= e.range.getLastRow(); r++) {
      syncKalenderStartTimeToKorSchema_(sheet, r);
    }
    return;
  }

  if (col === 4) {
    handleKalenderGenreEdit_(sheet, row);
    return;
  }

  if (col === 5) {
    handleKalenderLeagueEdit_(sheet, row);
    return;
  }

  if (col === 10) {
    handleKalenderLiveEdit_(sheet, row);
    return;
  }

  if (col === 11) handleKalenderScheduleCheckboxEdit_(e, sheet, row);
}

function handleKalenderGenreEdit_(sheet, row) {
  // Genre edited: update home/away team dropdowns and League dropdown.
  updateDropdownForRow(sheet, row);
  // Genre edited: update the League dropdown in column E.
  updateLeagueDropdown_(sheet, row);
}

function handleKalenderLeagueEdit_(sheet, row) {
  // League edited: update home/away team dropdowns.
  updateDropdownForRow(sheet, row);
}

function handleKalenderLiveEdit_(sheet, row) {
  const liveCell = sheet.getRange(row, 10);
  const checkboxCell = sheet.getRange(row, 11);

  if (String(liveCell.getDisplayValue()).trim() === 'Live') {
    const checkboxCriteria = SpreadsheetApp.newDataValidation()
      .requireCheckbox()
      .setAllowInvalid(false)
      .build();

    // Apply the checkbox to the adjacent cell and default it to un-checked (false)
    checkboxCell.setDataValidation(checkboxCriteria);
    checkboxCell.setValue(false);
  } else {
    checkboxCell.clearDataValidation();
    checkboxCell.clearContent();
  }
}

function handleKalenderScheduleCheckboxEdit_(e, sheet, row) {
  const checkboxCell = e.range;
  if (e.range.getValue() !== true) return;

  if (!kalenderRowHasScheduleDateAndTime_(sheet, row)) {
    handleKalenderMissingScheduleDetails_(checkboxCell);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Vänta medan körschemat skapas...',
    'Skapar körschema',
    20
  );
  const outputCell = sheet.getRange(row, 12); // Column L

  // Reset the checkbox immediately.
  checkboxCell.setValue(false);

  // Clear any previous result in column L.
  outputCell.clearContent();
  outputCell.clearFormat();
  outputCell.clearNote();

  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lock.waitLock(30000);
    lockAcquired = true;
    createKalenderScheduleFile_(e, sheet, row, outputCell);
  } catch (err) {
    handleKalenderScheduleCreationError_(err, outputCell);
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// Copies the start time in Kalender column B to E2 of the row's körschema.
// The körschema file is found through the link in column L, and the tab through J2,
// which holds the row's ID from column S (a file can hold several körscheman for one date).
function syncKalenderStartTimeToKorSchema_(sheet, row) {
  const formula = sheet.getRange(row, 12).getFormula(); // Column L
  const fileId = formula ? extractDriveId(formula) : null;
  if (!fileId) return; // No körschema created for this row.
  const id = String(sheet.getRange(row, KALENDER_ID_COL).getValue()).trim();
  if (!id) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Raden saknar körschema-ID i kolumn S',
      'Starttiden synkades inte',
      8
    );
    return;
  }

  let scheduleSpreadsheet;
  try {
    scheduleSpreadsheet = SpreadsheetApp.openById(fileId);
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Kunde inte öppna körschemat: ' + err.message,
      'Starttiden synkades inte',
      8
    );
    return;
  }

  const scheduleSheet = scheduleSpreadsheet.getSheets().find(function (s) {
    return String(s.getRange('J2').getValue()).trim() === id;
  });
  if (!scheduleSheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Hittade ingen flik med radens körschema-ID i J2 i "' + scheduleSpreadsheet.getName() + '"',
      'Starttiden synkades inte',
      8
    );
    return;
  }

  scheduleSheet.getRange('E2').setValue(sheet.getRange(row, 2).getValue());
  SpreadsheetApp.flush();
  // F2 in the körschema may depend on E2; send it back to Kalender column C as on creation.
  syncKorSchemaValueToKalender_({ range: scheduleSheet.getRange('F2') });
}

function kalenderRowHasScheduleDateAndTime_(sheet, row) {
  return sheet.getRange(row, 1).getValue() != '' &&
    sheet.getRange(row, 2).getValue() != '';
}

function handleKalenderMissingScheduleDetails_(checkboxCell) {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Fyll i datum och starttid innan du försöker igen',
    'Kan inte skapa körschema!',
    5
  );
  checkboxCell.setValue(false);
}

function createKalenderScheduleFile_(e, sheet, row, outputCell) {
  const genre = String(sheet.getRange(row, 4).getDisplayValue()).trim();
  const league = String(sheet.getRange(row, 5).getDisplayValue()).trim();
  const newFileName = String(sheet.getRange(row, 1).getDisplayValue());

  if (!genre || !league) {
    throw new Error('Genre eller liga saknas på raden.');
  }

  const searchKey = genre + ': ' + league;

  const ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const programSheet = ss.getSheetByName('Program');

  if (!programSheet) {
    throw new Error('Fliken "Program" kunde inte hittas.');
  }

  const lastRow = programSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error('Det finns inga rader att söka i på fliken "Program".');
  }

  // Search Program column C.
  const values = programSheet
    .getRange(2, 3, lastRow - 1, 1)
    .getDisplayValues();

  let matchRow = -1;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === searchKey) {
      matchRow = i + 2;
      break;
    }
  }

  if (matchRow === -1) {
    throw new Error(
      'Ingen matchning hittades i Program kolumn C för: ' + searchKey
    );
  }

  // The smart chip is in column G on the matching Program row.
  const sourceCell = programSheet.getRange(matchRow, 7);
  const schedule = createOrAppendKalenderSchedule_(sourceCell, newFileName);

  // Create and insert a native smart file chip instead of Rich Text
  outputCell.setFormula(`=HYPERLINK("${schedule.file.getUrl()}"; "Körschema")`);
  outputCell.setVerticalAlignment("middle");

/*
  outputCell.setNote(
    'Kopierad från Program!G' + matchRow +
    '\nOriginal: ' + sourceCell.getDisplayValue()
  );
*/
  SpreadsheetApp.getActiveSpreadsheet().toast(
    schedule.addedTab
      ? 'Körschemat har lagts till som en ny flik i befintlig fil'
      : 'Filen ligger i Produktion/Körscheman',
    'Körschemat skapat',
    5
  );

  const episodeName = String(sheet.getRange(row, 8).getDisplayValue()).trim() + ' | '
    + String(sheet.getRange(row, 9).getDisplayValue()).trim();
  schedule.sheet.setName(getUniqueKalenderTabName_(
    schedule.sheet.getParent(), schedule.sheet, episodeName
  ));
  schedule.sheet.getRange("G1").setValue(episodeName);
  schedule.sheet.getRange("E1").setValue(sheet.getRange(row, 1).getDisplayValue());
  schedule.sheet.getRange("E2").setValue(sheet.getRange(row, 2).getValue());
  // A new ID each time, so a körschema made again for the row doesn't share it with the old one.
  schedule.sheet.getRange("J2").setValue(newKalenderId_(sheet, row));
  SpreadsheetApp.flush();
  ensureKorSchemaEditTrigger_(schedule.sheet.getParent());
  syncKorSchemaValueToKalender_({ range: schedule.sheet.getRange('F2') });
}

function getKalenderTabName_(episodeName) {
  const tabName = episodeName.replace(/[\\/?*\[\]:]/g, ' ').trim().slice(0, 100);
  return tabName || 'Körschema';
}

function getUniqueKalenderTabName_(spreadsheet, sheet, episodeName) {
  const baseName = getKalenderTabName_(episodeName);
  const existingNames = {};
  spreadsheet.getSheets().forEach(function (existingSheet) {
    if (existingSheet.getSheetId() !== sheet.getSheetId()) {
      existingNames[existingSheet.getName()] = true;
    }
  });

  if (!existingNames[baseName]) return baseName;

  for (let suffixNumber = 2; ; suffixNumber++) {
    const suffix = ' (' + suffixNumber + ')';
    const candidate = baseName.slice(0, 100 - suffix.length) + suffix;
    if (!existingNames[candidate]) return candidate;
  }
}

function createOrAppendKalenderSchedule_(sourceCell, fileName) {
  const sourceFile = getFileChipFile_(sourceCell);
  const destinationFolder = getParentFolderForFile_(sourceFile);
  const sameNameFiles = destinationFolder.getFilesByName(fileName);
  let existingFile = null;

  while (sameNameFiles.hasNext()) {
    const candidate = sameNameFiles.next();
    if (existingFile) {
      throw new Error('Flera filer med namnet "' + fileName + '" finns redan.');
    }
    existingFile = candidate;
  }

  if (existingFile) {
    if (existingFile.getMimeType() !== MimeType.GOOGLE_SHEETS) {
      throw new Error('En fil med namnet "' + fileName + '" finns redan, men den är inte Google Sheets.');
    }

    const sourceSpreadsheet = SpreadsheetApp.openById(sourceFile.getId());
    const targetSpreadsheet = SpreadsheetApp.openById(existingFile.getId());
    const sourceSheet = sourceSpreadsheet.getSheets()[0];
    if (!sourceSheet) throw new Error('Mallfilen saknar ett kalkylblad.');

    return {
      file: existingFile,
      sheet: sourceSheet.copyTo(targetSpreadsheet),
      addedTab: true,
    };
  }

  const copiedFile = sourceFile.makeCopy(fileName, destinationFolder);
  const copiedSpreadsheet = SpreadsheetApp.openById(copiedFile.getId());
  return {
    file: copiedFile,
    sheet: copiedSpreadsheet.getSheets()[0],
    addedTab: false,
  };
}

function handleKalenderScheduleCreationError_(err, outputCell) {
  outputCell.setValue('Fel: ' + err.message);
  outputCell.setNote(err.stack || err.message);

  console.error(err);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Kunde inte skapa körschemat: ' + err.message,
    'Fel',
    8
  );
}


// Update the League dropdown in Kalender column E.
function updateLeagueDropdown_(sheet, row) {
  const selectedSport = String(
    sheet.getRange(row, 4).getDisplayValue()
  ).trim();

  const targetCell = sheet.getRange(row, 5);

  if (!selectedSport) {
    targetCell.clearDataValidations();
    targetCell.clearContent();
    return;
  }

  const ss = sheet.getParent();
  const programSheet = ss.getSheetByName('Program');

  if (!programSheet || programSheet.getLastRow() < 2) {
    targetCell.clearDataValidations();
    return;
  }

  const settingsData = programSheet
    .getRange(2, 1, programSheet.getLastRow() - 1, 2)
    .getValues();

  const matchingLeagues = settingsData
    .filter(function (r) {
      return r[0] &&
        String(r[0]).trim().toLowerCase() === selectedSport.toLowerCase();
    })
    .map(function (r) {
      return r[1];
    })
    .filter(Boolean);

  if (matchingLeagues.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(matchingLeagues, true)
      .setAllowInvalid(false)
      .build();

    targetCell.setDataValidation(rule);
  } else {
    targetCell.clearDataValidations();
  }
}

// Playout edit handling.
function handlePlayoutEdit_(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Content Type in column E controls the dropdown in column F.
  if (row < 2 || col !== 5) return;

  const targetCell = sheet.getRange(row, 6);
  const namedRangeName = String(
    sheet.getRange(row, 5).getDisplayValue()
  ).trim();

  if (!namedRangeName) {
    targetCell.clearDataValidation();
    targetCell.clearContent();
    return;
  }

  const namedRange = e.source.getRangeByName(namedRangeName);

  if (namedRange) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(namedRange)
      .setAllowInvalid(false)
      .build();

    targetCell.setDataValidation(rule);
  } else {
    targetCell.clearDataValidation();
  }
}

/**
 * Copies the Drive file represented by a smart chip.
 *
 * Destination:
 * One folder level above the file's current containing folder.
 *
 * Returns the copied Drive File object.
 */
function copyFileChipToParent_(cell) {
  const originalFile = getFileChipFile_(cell);
  const destinationFolder = getParentFolderForFile_(originalFile);
  const copyName = 'Copy of ' + originalFile.getName();
  return originalFile.makeCopy(copyName, destinationFolder);
}


/**
 * Manual version of the copier.
 *
 * This can open the copied file because it is intended to be
 * called directly by a user, for example from a menu.
 */
function copyAndOpenFileChip(cell) {
  const range = resolveCellRange_(cell);
  try {
    const copiedFile = copyFileChipToParent_(range);
    if (!copiedFile) return null;
    openUrlInBrowser(copiedFile.getUrl(), copiedFile.getName());
    return copiedFile;
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'Kunde inte kopiera filen:\n\n' + err.message
    );
    return null;
  }
}

function createRunningOrderFromTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  if (activeSheet.getName() !== 'Program') {
    SpreadsheetApp.getUi().alert(
      'Markera först en rad på fliken Program.'
    );
    return;
  }
  const activeRow = activeSheet.getActiveCell().getRow();
  if (activeRow < 2) {
    SpreadsheetApp.getUi().alert(
      'Markera en datarad, inte rubrikraden.'
    );
    return;
  }
  const targetCell = activeSheet.getRange(activeRow, 6);
  copyAndOpenFileChip(targetCell);
}

/**
 * Configuration settings
 */
const CONFIG = {
  SOURCE_SHEET: 'Kalender',
  DEST_SHEET: 'KalenderArkiv',
  SETTINGS_SHEET: 'Inställningar',
  DATE_CELL: 'B3',            // Cell containing the cutoff date (set to 'A3' if your date is in A3)
  DELETE_ORIGINAL_ROWS: true // Set to true if you want to remove rows from Kalender after copying
};

/**
 * Copies static values of all rows from 'Kalender' where the date in Column A is older than the date in Inställningar.
 * Ensures KalenderArkiv receives only raw values (no formulas or scripts).
 */
function archiveOldCalendarRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const destSheet = ss.getSheetByName(CONFIG.DEST_SHEET);
  const settingsSheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);

  if (!sourceSheet || !destSheet || !settingsSheet) {
    SpreadsheetApp.getUi().alert('Ett eller flera arkivblad saknas (kontrollera Kalender, KalenderArkiv och Inställningar).');
    return;
  }

  // 1. Get cutoff date from Inställningar
  const rawDate = settingsSheet.getRange(CONFIG.DATE_CELL).getValue();
  const cutoffDate = new Date(rawDate);
  if (isNaN(cutoffDate.getTime())) {
    SpreadsheetApp.getUi().alert(`Kunde inte läsa av ett giltigt datum från Inställningar (${CONFIG.DATE_CELL}).`);
    return;
  }

  // Normalize cutoff date to start of day (midnight) for accurate date-only comparison
  cutoffDate.setHours(0, 0, 0, 0);

  // 2. Read Kalender static values (skipping header row 1)
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Ingen data hittades i Kalender.');
    return;
  }

  // getValues() retrieves the computed, raw values (stripping all underlying formulas)
  const sourceValues = sourceSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const rowsToCopy = [];
  const rowsToDelete = [];

  for (let i = 0; i < sourceValues.length; i++) {
    const rowDateVal = sourceValues[i][0]; // Column A
    if (rowDateVal) {
      const rowDate = new Date(rowDateVal);
      if (!isNaN(rowDate.getTime())) {
        rowDate.setHours(0, 0, 0, 0);
        if (rowDate < cutoffDate) {
          // Push a shallow clone of the row values
          rowsToCopy.push(sourceValues[i].slice());
          rowsToDelete.push(i + 2); // 1-based sheet row index
        }
      }
    }
  }

  if (rowsToCopy.length === 0) {
    SpreadsheetApp.getUi().alert('Inga rader var äldre än ' + Utilities.formatDate(cutoffDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.');
    return;
  }

  // 3. Find first empty row in KalenderArkiv by inspecting Column A (Datum)
  // This avoids placing data at row 1012 if checkboxes/empty rows exist below row 1
  const destColA = destSheet.getRange('A:A').getValues();
  let firstEmptyRow = destColA.length + 1;
  for (let r = 1; r < destColA.length; r++) { // Skip header row 0 (A1)
    if (destColA[r][0] === '' || destColA[r][0] === null) {
      firstEmptyRow = r + 1;
      break;
    }
  }

  const numRows = rowsToCopy.length;
  const numCols = rowsToCopy[0].length;
  const destTargetRange = destSheet.getRange(firstEmptyRow, 1, numRows, numCols);

  // Clear any existing data validations/checkboxes in the destination target range
  destTargetRange.clearDataValidations();

  // Write static values only
  destTargetRange.setValues(rowsToCopy);

  // 4. Optionally delete copied rows from Kalender (from bottom up to preserve row indices)
  if (CONFIG.DELETE_ORIGINAL_ROWS) {
    for (let r = rowsToDelete.length - 1; r >= 0; r--) {
      sourceSheet.deleteRow(rowsToDelete[r]);
    }
  }

  SpreadsheetApp.getUi().alert(
    `Klar! ${rowsToCopy.length} rad(er) kopierades som rena värden till '${CONFIG.DEST_SHEET}' (börjar på rad ${firstEmptyRow})` +
    (CONFIG.DELETE_ORIGINAL_ROWS ? ' och togs bort från Kalender.' : '.')
  );
}


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

/**
 * Checks the monitored parent folder and its subfolders for newly uploaded files.
 * Run this function on a time-driven trigger (e.g., every 5 to 15 minutes).
 */
function checkForNewAssets() {
  const props = PropertiesService.getScriptProperties();
  const lastCheckIso = props.getProperty('LAST_CHECK_TIME');
  const now = new Date();

  // If first run, look back 1 hour (adjust if needed)
  const lastCheckTime = lastCheckIso
    ? new Date(lastCheckIso)
    : new Date(now.getTime() - 60 * 60 * 1000);

  const parentFolder = DriveApp.getFolderById(DRIVE_WATCH_CONFIG.PARENT_FOLDER_ID);

  // Array to hold all newly discovered files across folders
  const newFiles = [];

  // 1. Scan files directly inside the parent folder
  scanFolderForNewFiles(parentFolder, 'Root / Parent Folder', lastCheckTime, newFiles);

  // 2. Scan all subfolders inside the parent folder
  const subfolders = parentFolder.getFolders();
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    scanFolderForNewFiles(subfolder, subfolder.getName(), lastCheckTime, newFiles);
  }

  // If new files were detected, send the formatted email notification
  if (newFiles.length > 0) {
    sendNotificationEmail(parentFolder.getName(), newFiles);
  }

  // Update the checkpoint timestamp to the current run time
  props.setProperty('LAST_CHECK_TIME', now.toISOString());
}

/**
 * Helper function to inspect files within a specific folder.
 */
function scanFolderForNewFiles(folder, folderLabel, lastCheckTime, resultsArray) {
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const createdTime = file.getDateCreated();

    // Check if the file was created after the previous checkpoint
    if (createdTime.getTime() > lastCheckTime.getTime()) {
      resultsArray.push({
        folderName: folderLabel,
        folderUrl: folder.getUrl(),
        name: file.getName(),
        url: file.getUrl(),
        sizeMb: (file.getSize() / (1024 * 1024)).toFixed(2),
        type: file.getMimeType(),
        created: Utilities.formatDate(createdTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      });
    }
  }
}

/**
 * Sends a structured notification email grouped by subfolder.
 */
function sendNotificationEmail(parentFolderName, newFiles) {
  const count = newFiles.length;

  // Group assets by subfolder name
  const groupedByFolder = {};
  newFiles.forEach(file => {
    if (!groupedByFolder[file.folderName]) {
      groupedByFolder[file.folderName] = [];
    }
    groupedByFolder[file.folderName].push(file);
  });

  const updatedFoldersList = Object.keys(groupedByFolder).join(', ');
  const subject = `${DRIVE_WATCH_CONFIG.EMAIL_SUBJECT_PREFIX}: ${count} new asset(s) in [${updatedFoldersList}]`;

  let body = `Hello,\n\n${count} new asset(s) were uploaded across the monitored folders in "${parentFolderName}".\n\n`;
  body += `==============================================\n\n`;

  // Iterate through each updated folder
  for (const [folderName, files] of Object.entries(groupedByFolder)) {
    body += `📁 SUBFOLDER: ${folderName} (${files.length} new asset${files.length > 1 ? 's' : ''})\n`;
    body += `Folder link: ${files[0].folderUrl}\n`;
    body += `----------------------------------------------\n`;

    files.forEach((file, idx) => {
      body += `  ${idx + 1}. ${file.name}\n`;
      body += `     • Size: ${file.sizeMb} MB\n`;
      body += `     • Type: ${file.type}\n`;
      body += `     • Uploaded: ${file.created}\n`;
      body += `     • File link: ${file.url}\n\n`;
    });
  }

  body += `==============================================\n`;
  body += `Best regards,\nAutomated Asset Watcher`;

  sendMail_(DRIVE_WATCH_CONFIG.NOTIFICATION_EMAIL, subject, body);

  Logger.log(`Notification sent for ${count} file(s) across: ${updatedFoldersList}`);
}

/**
 * Reset checkpoint if you want to force a re-check of the last hour.
 */
function resetCheckpoint() {
  PropertiesService.getScriptProperties().deleteProperty('LAST_CHECK_TIME');
  Logger.log('Checkpoint reset. The next run will scan the past hour.');
}
