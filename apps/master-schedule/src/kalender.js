// Kalender edits (routed from handleEdit): league and team dropdowns and the Live checkbox.
// The körschema checkbox and start-time sync are in korschema.js.

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
