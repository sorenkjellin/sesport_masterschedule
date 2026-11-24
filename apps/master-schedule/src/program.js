// Program column F (Längd): typed by hand, except on rows with a körschema linked in column G. There
// it is the körschema's length, F2 - E2 (end minus start) on its first visible tab.
//
// checkProgramLengths runs every minute (installProgramLengthTrigger) with the access of whoever
// installed it, who must be able to open the linked files (linkedFiles.js says when each is read). It
// also keeps Kalender's end times up to date (kalenderEndTimes.js). One timer instead of an edit trigger in each
// file, because Google allows 20 triggers per user and script.
// handleEdit puts the length back when someone types over it, and fills it in when a link is added.

const PROGRAM_SHEET = 'Program';
const PROGRAM_LENGTH_COL = 6; // F
const PROGRAM_LENGTH_HEADER = 'Längd';
const PROGRAM_NOTE = 'Längden styrs av grundkörschemat';
const PROGRAM_FILES = {
  cacheKey: 'PROGRAM_LENGTHS_2',
  read: function (id) { return { days: readKorSchemaLength_(id) }; },
};
const TRIGGERS_CHECKED_KEY = 'KORSCHEMA_TRIGGERS_CHECKED';

/** The time trigger's function: brings Program's and Kalender's governed cells up to date. */
function checkProgramLengths() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // Busy (a körschema being created?): the next run catches up.
  try {
    // One failing (a column moved?) must not stop the others.
    const changes = {};
    const errors = [];
    [
      function () { updateProgramLengths_({ changes: changes }); },
      function () { updateKalenderEndTimes_({ changes: changes }); },
      removeUnusedKorSchemaTriggersHourly_,
    ].forEach(function (step) {
      try {
        step();
      } catch (e) {
        console.error(e.stack || e.message);
        errors.push(e.message);
      }
    });
    if (errors.length) throw new Error(errors.join('; '));
  } finally {
    lock.releaseLock();
  }
}

// Here because this runs as the owner of the körschema triggers, whoever archives rows.
function removeUnusedKorSchemaTriggersHourly_() {
  const cache = CacheService.getScriptCache();
  if (cache.get(TRIGGERS_CHECKED_KEY)) return;
  cache.put(TRIGGERS_CHECKED_KEY, 'yes', 3600);
  removeUnusedKorSchemaTriggers_();
  restoreKorSchemaTriggers_();
}

function installProgramLengthTrigger() {
  deleteTriggers_('checkProgramLengths');
  ScriptApp.newTrigger('checkProgramLengths').timeBased().everyMinutes(1).create();
  Logger.log('checkProgramLengths now runs every minute.');
}

// Program edits (routed from handleEdit): the length (F) or the link (G) of some rows.
function handleProgramEdit_(e) {
  const lengthEdited = coversColumn_(e.range, PROGRAM_LENGTH_COL);
  const linkEdited = coversColumn_(e.range, PROGRAM_LENGTH_COL + 1);
  if (!lengthEdited && !linkEdited) return;
  const firstRow = e.range.getRow();
  const lastRow = firstRow + e.range.getNumRows() - 1;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; // The next checkProgramLengths corrects the rows.
  try {
    const restored = updateProgramLengths_({ firstRow: firstRow, lastRow: lastRow });
    if (restored && lengthEdited && !linkEdited) {
      toast_('Längden styrs av körschemat i kolumn G', 'Längden återställdes', 6);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates the rows of Program from options.firstRow to options.lastRow, reading only files not read
 * before; or, with options.changes, all rows, also reading the files changed since the last check
 * (see refreshLinkedFiles_). Returns how many lengths were written.
 */
function updateProgramLengths_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PROGRAM_SHEET);
  if (!sheet) throw new Error('Fliken "Program" kunde inte hittas.');
  const readRows = function () {
    return readProgramRows_(ss).filter(function (r) {
      return options.changes || (r.row >= options.firstRow && r.row <= options.lastRow);
    });
  };

  let rows = readRows();
  const fileIds = rows.map(function (r) { return r.fileId; });
  if (refreshLinkedFiles_(PROGRAM_FILES, fileIds, options.changes)) rows = readRows(); // Rows may have moved.
  const files = readLinkedCache_(PROGRAM_FILES).files;

  const format = programLengthFormat_(ss.getSpreadsheetLocale());
  let written = 0;
  rows.forEach(function (r) {
    const file = r.fileId ? files[r.fileId] : null;
    const cell = sheet.getRange(r.row, PROGRAM_LENGTH_COL);
    // The whole column, but a plain number typed by hand (90?) is not a time: left as it is.
    const isPlainNumber = r.value !== undefined && !r.isTime && !(file && file.days != null);
    if (!isPlainNumber && !sameFormat_(r.format, format)) cell.setNumberFormat(format);
    if (r.fileId && !file) return; // Linked just now: the next run reads it.
    if (file && file.days != null && !sameDuration_(r.value, file.days)) {
      cell.setValue(file.days);
      written++;
    }
    markLinkedCell_(cell, r.note, !!file, PROGRAM_NOTE, 'längden', file && file.error);
  });
  return written;
}

/**
 * Every row below the header as { row, value, format, isTime (column F's number format), note, fileId (of
 * the file linked in G, as a chip or a link, or null) }. Throws when column F is not Längd.
 */
function readProgramRows_(ss) {
  const response = Sheets.Spreadsheets.get(ss.getId(), {
    ranges: [PROGRAM_SHEET + '!F1:G'],
    includeGridData: true,
    fields: 'sheets(data(rowData(values(effectiveValue,formattedValue,note,hyperlink,chipRuns,' +
      'effectiveFormat(numberFormat)))))',
  });
  const data = ((response.sheets || [])[0] || {}).data || [];
  const rowData = (data[0] || {}).rowData || [];
  const header = ((((rowData[0] || {}).values || [])[0] || {}).formattedValue || '').trim();
  if (header !== PROGRAM_LENGTH_HEADER) {
    throw new Error('Program column F should be "' + PROGRAM_LENGTH_HEADER + '", is "' + header + '"');
  }
  return rowData.slice(1).map(function (rowValues, i) {
    const length = (rowValues.values || [])[0] || {};
    const link = (rowValues.values || [])[1] || {};
    const value = length.effectiveValue || {};
    const format = (length.effectiveFormat || {}).numberFormat || {};
    return {
      row: i + 2,
      value: value.numberValue,
      format: format.pattern || '',
      isTime: format.type === 'TIME',
      note: length.note || '',
      fileId: linkedFileId_(link),
    };
  });
}

/** A körschema's length in days, from E2:F2 of its first visible tab; throws when there is none. */
function readKorSchemaLength_(fileId) {
  const answer = Sheets.Spreadsheets.Values.get(fileId, 'E2:F2', {
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
  });
  const values = (answer.values || [])[0] || [];
  return korSchemaLength_(values[0], values[1]);
}

/** end - start (Sheets times, in days) in whole seconds; an empty start is 0:00, and past midnight works. */
function korSchemaLength_(start, end) {
  if (end === '' || end == null) throw new Error('F2 är tom');
  if (typeof end !== 'number') throw new Error('F2 är inte en tid');
  if (start === '' || start == null) start = 0;
  if (typeof start !== 'number') throw new Error('E2 är inte en tid');
  // Compared in whole seconds, so that rounding never turns an equal start and end into 24 hours.
  let seconds = Math.round((end - start) * 86400);
  if (seconds < 0) seconds += 86400;
  if (seconds < 0) throw new Error('F2 är före E2');
  return (seconds || 0) / 86400; // || 0: rounding can give -0
}

/**
 * Column F's number format: hours without a leading zero and minutes, with the time separator of the
 * spreadsheet's locale (sv_SE: 1.30). Sheets does not say what it is, and JavaScript's own locale
 * data has ":" for Swedish, so the languages that write times with "." are listed here.
 */
function programLengthFormat_(locale) {
  return '[h]"' + timeSeparator_(locale) + '"mm';
}

function timeSeparator_(locale) {
  const language = String(locale || '').split(/[_-]/)[0].toLowerCase();
  return ['sv', 'da', 'fi', 'id'].indexOf(language) !== -1 ? '.' : ':';
}

/** Whether two number formats are the same, however their literals are quoted (Sheets may rewrite it). */
function sameFormat_(a, b) {
  return a.replace(/["\\]/g, '') === b.replace(/["\\]/g, '');
}
