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
