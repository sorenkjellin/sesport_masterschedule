// Körscheman: creating one from a Kalender row (from the template in Program column G), the ID
// that links the two, and syncing start time (Kalender -> körschema) and F2 (körschema -> Kalender).
// Each körschema spreadsheet gets an installable edit trigger from this project.

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
  const destinationFolder = getGrandparentFolder_(sourceFile);
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
  const destinationFolder = getGrandparentFolder_(originalFile);
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
    openUrlInBrowser_(copiedFile.getUrl(), copiedFile.getName(), 'Opening File Copy');
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
