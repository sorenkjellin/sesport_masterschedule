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
  const COLUMNS_SENT = 16; // A to P

  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert(`Sheet "${sheetName}" was not found!`);
    return;
  }

  // Display values, so dates and times are sent as shown and no formulas are carried over.
  const values = sourceSheet.getRange(1, 1, Math.max(sourceSheet.getLastRow(), 1), COLUMNS_SENT).getDisplayValues();
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
