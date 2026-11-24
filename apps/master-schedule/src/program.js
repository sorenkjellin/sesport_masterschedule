// Program column F (Längd): typed by hand, except on rows with a körschema linked in column G. There
// it is the körschema's length, F2 - E2 (end minus start) on its first visible tab.
//
// checkProgramLengths runs every minute (installProgramLengthTrigger) with the access of whoever
// installed it, who must be able to open the linked files. It reads a file when it is new here or has
// changed, and all of them once an hour. One timer instead of an edit trigger in each file, because
// Google allows 20 triggers per user and script.
// handleEdit puts the length back when someone types over it, and fills it in when a link is added.

const PROGRAM_SHEET = 'Program';
const PROGRAM_LENGTH_COL = 6; // F
const PROGRAM_LENGTH_HEADER = 'Längd';
// The lines this code writes in column F's notes start like this; other lines are left alone.
const PROGRAM_NOTE_PREFIX = 'Körschema: ';
const PROGRAM_LINKED_COLOR = '#666666';
const PROGRAM_CACHE_KEY = 'PROGRAM_LENGTHS';
const TRIGGERS_CHECKED_KEY = 'KORSCHEMA_TRIGGERS_CHECKED';
const PROGRAM_FULL_READ_MS = 60 * 60 * 1000;
// Drive can be late reporting a change, so each check also looks back over the last few minutes.
const PROGRAM_LOOK_BACK_MS = 5 * 60 * 1000;

/** The time trigger's function: brings column F up to date with the linked körscheman. */
function checkProgramLengths() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // Busy (a körschema being created?): the next run catches up.
  try {
    updateProgramLengths_({});
    removeUnusedKorSchemaTriggersHourly_();
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
}

function installProgramLengthTrigger() {
  deleteTriggers_('checkProgramLengths');
  ScriptApp.newTrigger('checkProgramLengths').timeBased().everyMinutes(1).create();
  Logger.log('checkProgramLengths now runs every minute.');
}

// Program edits (routed from handleEdit): the length (F) or the link (G) of some rows.
function handleProgramEdit_(e) {
  const first = e.range.getColumn();
  const last = first + e.range.getNumColumns() - 1;
  if (last < PROGRAM_LENGTH_COL || first > PROGRAM_LENGTH_COL + 1) return;
  const firstRow = e.range.getRow();
  const lastRow = firstRow + e.range.getNumRows() - 1;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return; // The next checkProgramLengths corrects the rows.
  try {
    const restored = updateProgramLengths_({ firstRow: firstRow, lastRow: lastRow });
    if (restored && last === PROGRAM_LENGTH_COL) {
      toast_('Längden styrs av körschemat i kolumn G', 'Längden återställdes', 6);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Updates the rows of Program from options.firstRow to options.lastRow, reading only files not read
 * before; without them all rows, also reading the files changed since the last check. Returns how
 * many lengths were written.
 */
function updateProgramLengths_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PROGRAM_SHEET);
  if (!sheet) throw new Error('Fliken "Program" kunde inte hittas.');
  const allRows = !options.firstRow;
  const readRows = function () {
    return readProgramRows_(ss).filter(function (r) {
      return allRows || (r.row >= options.firstRow && r.row <= options.lastRow);
    });
  };

  let rows = readRows();
  if (refreshProgramFiles_(rows, allRows)) rows = readRows(); // Rows may have moved meanwhile.
  const files = readProgramCache_().files;

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
    const note = programNote_(r.note, file);
    if (note !== r.note) {
      const wasLinked = hasProgramNote_(r.note);
      cell.setNotes([[note]]);
      if (wasLinked !== !!file) cell.setFontColor(file ? PROGRAM_LINKED_COLOR : null);
    }
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

/** The Drive ID of the file a cell (as the Sheets API gives it) links to, or null. */
function linkedFileId_(cell) {
  const uris = (cell.chipRuns || []).map(function (run) {
    return ((run.chip || {}).richLinkProperties || {}).uri;
  }).concat(cell.hyperlink).filter(Boolean);
  for (const uri of uris) {
    const id = extractDriveId(uri);
    if (id) return id;
  }
  return null;
}

/**
 * Reads the rows' files that need it into the cache: those not there, and with lookForChanges those
 * that failed or have changed, or once an hour all of them. Returns how many were read.
 */
function refreshProgramFiles_(rows, lookForChanges) {
  const cache = readProgramCache_();
  const now = Date.now();
  const ids = Array.from(new Set(rows.map(function (r) { return r.fileId; }).filter(Boolean)));
  const readAll = lookForChanges && !(now - cache.fullAt < PROGRAM_FULL_READ_MS);
  const changed = lookForChanges && !readAll && ids.length
    ? spreadsheetsChangedSince_(new Date(cache.checkedAt - PROGRAM_LOOK_BACK_MS)) : new Set();

  const toRead = ids.filter(function (id) {
    const known = cache.files[id];
    return readAll || !known || changed.has(id) || (lookForChanges && known.error);
  });
  const files = {};
  // All rows were read: forget files no longer linked. Otherwise the others are still needed.
  Object.keys(cache.files).forEach(function (id) {
    if (!lookForChanges || ids.indexOf(id) !== -1) files[id] = cache.files[id];
  });
  toRead.forEach(function (id) {
    try {
      files[id] = { days: readKorSchemaLength_(id) };
    } catch (e) {
      console.log('Could not read the length of körschema ' + id + ': ' + e.message);
      // The last known length is kept.
      files[id] = { days: (files[id] || {}).days, error: String(e.message).slice(0, 200) };
    }
  });

  cache.files = files;
  if (lookForChanges) cache.checkedAt = now;
  if (readAll) cache.fullAt = now;
  // The longest CacheService keeps it; without it, everything is simply read again.
  CacheService.getScriptCache().put(PROGRAM_CACHE_KEY, JSON.stringify(cache), 21600);
  return toRead.length;
}

/** { files: { <file ID>: { days, error } }, checkedAt, fullAt (ms) }; the times are absent at first. */
function readProgramCache_() {
  const text = CacheService.getScriptCache().get(PROGRAM_CACHE_KEY);
  return text ? JSON.parse(text) : { files: {} };
}

/** The IDs of the spreadsheets, in any drive the user can see, modified after date. */
function spreadsheetsChangedSince_(date) {
  const ids = new Set();
  let pageToken = '';
  do {
    const query = {
      q: "mimeType = 'application/vnd.google-apps.spreadsheet' and modifiedTime > '" + date.toISOString() + "'",
      corpora: 'allDrives', includeItemsFromAllDrives: 'true', supportsAllDrives: 'true',
      pageSize: '1000', fields: 'nextPageToken,files(id)',
    };
    if (pageToken) query.pageToken = pageToken;
    const url = 'https://www.googleapis.com/drive/v3/files?' + Object.keys(query).map(function (k) {
      return k + '=' + encodeURIComponent(query[k]);
    }).join('&');
    const page = googleApiRequest_(url, { method: 'get' }, 'Listing changed spreadsheets');
    (page.files || []).forEach(function (f) { ids.add(f.id); });
    pageToken = page.nextPageToken;
  } while (pageToken);
  return ids;
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
  const language = String(locale || '').split(/[_-]/)[0].toLowerCase();
  return '[h]"' + (['sv', 'da', 'fi', 'id'].indexOf(language) !== -1 ? '.' : ':') + '"mm';
}

/** Whether two number formats are the same, however their literals are quoted (Sheets may rewrite it). */
function sameFormat_(a, b) {
  return a.replace(/["\\]/g, '') === b.replace(/["\\]/g, '');
}

function sameDuration_(a, b) {
  return typeof a === 'number' && Math.abs(a - b) < 0.5 / 86400;
}

function hasProgramNote_(note) {
  return note.split('\n').some(function (line) { return line.indexOf(PROGRAM_NOTE_PREFIX) === 0; });
}

/** note with this code's line replaced by the one for file (from the cache), or removed without one. */
function programNote_(note, file) {
  const lines = note.split('\n').filter(function (line) { return line.indexOf(PROGRAM_NOTE_PREFIX) !== 0; });
  if (file) {
    lines.push(PROGRAM_NOTE_PREFIX + (file.error
      ? 'längden kunde inte hämtas (' + file.error.replace(/\s+/g, ' ') + '). Det senaste värdet står kvar.'
      : 'längden hämtas från filen i kolumn G.'));
  }
  return lines.join('\n').trim();
}
