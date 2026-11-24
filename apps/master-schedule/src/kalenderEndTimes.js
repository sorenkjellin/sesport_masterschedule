// Kalender column C (Slutar): typed by hand, except on rows with a körschema linked in column L. There
// it is the end time (F2) of the körschema's tab, the one with the row's ID from column S in J2.
// An edit in the körschema updates it at once (syncKorSchemaValueToKalender_); checkProgramLengths
// (program.js) does too, every minute, and handleEdit puts it back when someone types over it.
// On the other rows, editing Börjar, Genre or Ligans namn fills it in from Program's Längd
// (copyProgramLengths_): a copy, which can be typed over.

const KALENDER_END_COL = 3; // C
const KALENDER_LINK_COL = 12; // L
const KALENDER_FIRST_READ_COL = 2; // B
const KALENDER_START_COL = 2; // B
const KALENDER_GENRE_COL = 4; // D
const KALENDER_LEAGUE_COL = 5; // E
const KALENDER_END_NOTE = 'Sluttid styrs av körschemats längd';
const KALENDER_FILES = { cacheKey: 'KALENDER_END_TIMES', read: readKorSchemaEndTimes_ };

// Kalender edits of C or L (from handleKalenderEdit_).
function handleKalenderEndEdit_(e) {
  const endEdited = coversColumn_(e.range, KALENDER_END_COL);
  const linkEdited = coversColumn_(e.range, KALENDER_LINK_COL);
  if (!endEdited && !linkEdited) return;
  const firstRow = e.range.getRow();
  const lastRow = firstRow + e.range.getNumRows() - 1;
  if (lastRow < 2) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; // The next checkProgramLengths corrects the rows.
  try {
    const restored = updateKalenderEndTimes_({ firstRow: firstRow, lastRow: lastRow });
    if (restored && endEdited && !linkEdited && !coversColumn_(e.range, 2)) {
      toast_('Sluttiden styrs av körschemats längd', 'Sluttiden återställdes', 6);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Kalender edits of Börjar, Genre or Ligans namn (from handleKalenderEdit_): on each edited row with
 * all three and no körschema, Slutar becomes Börjar + the Längd of the Program row with that genre and
 * league. With no such row, Slutar is emptied, so that it shows; when the row has no length, it
 * is left alone. An end time pasted along with them is kept.
 */
function copyProgramLengths_(e) {
  const edited = [KALENDER_START_COL, KALENDER_GENRE_COL, KALENDER_LEAGUE_COL].some(function (col) {
    return coversColumn_(e.range, col);
  });
  if (!edited) return;
  const firstRow = e.range.getRow();
  const lastRow = firstRow + e.range.getNumRows() - 1;
  const endEdited = coversColumn_(e.range, KALENDER_END_COL);

  SpreadsheetApp.flush(); // The reads below go past SpreadsheetApp.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rows = readKalenderRows_(ss).filter(function (r) {
    return r.row >= firstRow && r.row <= lastRow && !r.fileId && r.start !== null && r.genre && r.league &&
      !(endEdited && r.hasEnd);
  });
  if (!rows.length) return;
  const lengths = readProgramLengths_(ss);
  const sheet = ss.getSheetByName(SCHEDULE_SHEET);
  rows.forEach(function (r) {
    const length = lengths[programKey_(r.genre, r.league)];
    const cell = sheet.getRange(r.row, KALENDER_END_COL);
    if (length === undefined) cell.clearContent(); // Not in Program.
    if (!length) return;
    // In whole seconds, and as a time of day also past midnight, as it would be typed.
    cell.setValue(Math.round((r.start + length) * 86400) % 86400 / 86400);
    if (!r.isTime) cell.setNumberFormat(kalenderTimeFormat_(ss));
  });
}

/**
 * Program's lengths (days) by programKey_, null for those without one; of rows with the same genre and
 * league, the first with a length counts.
 */
function readProgramLengths_(ss) {
  const answer = Sheets.Spreadsheets.Values.get(ss.getId(), PROGRAM_SHEET + '!A1:F', {
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
  });
  const rows = answer.values || [];
  const header = String((rows[0] || [])[PROGRAM_LENGTH_COL - 1] || '').trim();
  if (header !== PROGRAM_LENGTH_HEADER) {
    throw new Error('Program column F should be "' + PROGRAM_LENGTH_HEADER + '", is "' + header + '"');
  }
  const lengths = {};
  rows.slice(1).forEach(function (r) {
    const key = programKey_(r[0], r[1]);
    const length = timeOfCell_(r[PROGRAM_LENGTH_COL - 1]);
    // Not a length: none, or a plain number (90?).
    if (!lengths[key]) lengths[key] = length > 0 && length < 1 ? length : null;
  });
  return lengths;
}

function programKey_(genre, league) {
  return [genre, league].map(function (v) { return String(v == null ? '' : v).trim().toLowerCase(); }).join('\n');
}

/** A time (days) from a cell's value: a number, or text such as "18.00" or "18:00"; otherwise null. */
function timeOfCell_(value) {
  if (typeof value === 'number') return value;
  const m = /^\s*(\d{1,2})[:.](\d{2})(?:[:.](\d{2}))?\s*$/.exec(String(value == null ? '' : value));
  return m ? (Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] || 0)) / 86400 : null;
}

function kalenderTimeFormat_(ss) {
  return 'hh"' + timeSeparator_(ss.getSpreadsheetLocale()) + '"mm';
}

/**
 * Updates the rows of Kalender from options.firstRow to options.lastRow, reading only files not read
 * before; or, with options.changes, all rows (see refreshLinkedFiles_). Returns how many end times
 * were written.
 */
function updateKalenderEndTimes_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCHEDULE_SHEET);
  const readRows = function () {
    return readKalenderRows_(ss).filter(function (r) {
      return options.changes || (r.row >= options.firstRow && r.row <= options.lastRow);
    });
  };

  let rows = readRows();
  const fileIds = rows.map(function (r) { return r.fileId; });
  if (refreshLinkedFiles_(KALENDER_FILES, fileIds, options.changes)) rows = readRows(); // Rows may have moved.
  const files = readLinkedCache_(KALENDER_FILES).files;

  let written = 0;
  rows.forEach(function (r) {
    const file = r.fileId ? files[r.fileId] : null;
    if (r.fileId && (!file || file.stale)) return; // The next run reads it.
    const cell = sheet.getRange(r.row, KALENDER_END_COL);
    const end = file ? kalenderEndTime_(file, r.id) : {};
    if (end.days != null && !sameDuration_(r.value, end.days)) {
      // The row may have moved just now, and a wrong end time on an unlinked row would stay there.
      if (String(sheet.getRange(r.row, KALENDER_ID_COL).getDisplayValue()).trim() !== r.id) return;
      cell.setValue(end.days);
      written++;
    }
    if (end.days != null && !r.isTime) cell.setNumberFormat(kalenderTimeFormat_(ss));
    markLinkedCell_(cell, r.note, !!file, KALENDER_END_NOTE, 'sluttiden', end.error);
  });
  return written;
}

/** The row's end time, as { days, error }: its tab's, from the file's entry. */
function kalenderEndTime_(file, id) {
  const tab = id ? (file.tabs || {})[id] : null;
  if (tab) return { days: tab.days, error: file.error || tab.error };
  return { error: file.error || (id ? 'ingen flik har radens ID i J2' : 'raden saknar ID i kolumn S') };
}

/**
 * Every row below the header as { row, start (days, or null), genre, league, value, hasEnd, isTime, note
 * (the last four of C), fileId (linked in L, or null), id (S) }. Throws when column C is not Slutar.
 */
function readKalenderRows_(ss) {
  const response = Sheets.Spreadsheets.get(ss.getId(), {
    ranges: [SCHEDULE_SHEET + '!B1:S'],
    includeGridData: true,
    fields: 'sheets(data(rowData(values(effectiveValue,formattedValue,note,hyperlink,chipRuns,' +
      'userEnteredValue(formulaValue),effectiveFormat(numberFormat(type))))))',
  });
  const data = ((response.sheets || [])[0] || {}).data || [];
  const rowData = (data[0] || {}).rowData || [];
  const at = function (rowValues, col) { return (rowValues.values || [])[col - KALENDER_FIRST_READ_COL] || {}; };
  const header = (at(rowData[0] || {}, KALENDER_END_COL).formattedValue || '').trim();
  if (COLUMNS.end.indexOf(header) === -1) {
    throw new Error('Kalender column C should be "' + COLUMNS.end[0] + '", is "' + header + '"');
  }
  return rowData.slice(1).map(function (rowValues, i) {
    const end = at(rowValues, KALENDER_END_COL);
    const link = at(rowValues, KALENDER_LINK_COL);
    const type = ((end.effectiveFormat || {}).numberFormat || {}).type;
    return {
      row: i + 2,
      start: startTime_(at(rowValues, KALENDER_START_COL)),
      genre: (at(rowValues, KALENDER_GENRE_COL).formattedValue || '').trim(),
      league: (at(rowValues, KALENDER_LEAGUE_COL).formattedValue || '').trim(),
      hasEnd: !!(end.formattedValue || '').trim(),
      value: (end.effectiveValue || {}).numberValue,
      isTime: type === 'TIME' || type === 'DATE_TIME',
      note: end.note || '',
      fileId: linkedFileId_(link) || extractDriveId((link.userEnteredValue || {}).formulaValue || ''),
      id: (at(rowValues, KALENDER_ID_COL).formattedValue || '').trim(),
    };
  });
}

function startTime_(cell) {
  const value = cell.effectiveValue || {};
  return timeOfCell_(value.numberValue !== undefined ? value.numberValue : value.stringValue);
}

/** A körschema file's end times: { tabs: { <the ID in J2>: { days } or { days (the last known), error } } }. */
function readKorSchemaEndTimes_(fileId, previous) {
  const titles = (Sheets.Spreadsheets.get(fileId, { fields: 'sheets(properties(title))' }).sheets || [])
    .map(function (s) { return s.properties.title; });
  const answer = Sheets.Spreadsheets.Values.batchGet(fileId, {
    ranges: titles.map(function (t) { return "'" + t.replace(/'/g, "''") + "'!F2:J2"; }),
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
  });
  const tabs = {};
  (answer.valueRanges || []).forEach(function (range) {
    const values = (range.values || [])[0] || [];
    const id = String(values[4] == null ? '' : values[4]).trim();
    if (!id || tabs[id]) return; // As when syncing the start time, the first tab with the ID counts.
    if (typeof values[0] === 'number') {
      tabs[id] = { days: values[0] };
    } else {
      const last = ((previous || {}).tabs || {})[id] || {};
      tabs[id] = { days: last.days, error: values[0] === '' ? 'F2 är tom' : 'F2 är inte en tid' };
    }
  });
  return { tabs: tabs };
}
