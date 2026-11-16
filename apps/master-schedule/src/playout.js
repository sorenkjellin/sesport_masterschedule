// Playout edits (routed from handleEdit): Content Type in column E sets the dropdown in column F.

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
