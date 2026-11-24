// Cells whose value comes from a linked körschema: Program's Längd (program.js) and Kalender's Slutar
// (kalenderEndTimes.js). What they share: which files need reading, the cache of what was read, and
// the look of a governed cell (grey, with a note).

const LINKED_COLOR = '#666666';
// The notes' lines written here are the cell's own note, or start with this (a problem); others are kept.
const LINKED_NOTE_PREFIX = 'Körschema: ';
// Each file is read again at least this often, in case a change was missed.
const LINKED_REREAD_MS = 60 * 60 * 1000;
// Drive can be late reporting a change, so each check also looks back over the last few minutes.
const LINKED_LOOK_BACK_MS = 5 * 60 * 1000;
// Files read per run and sheet: Google allows 60 Sheets API reads a minute and user.
const LINKED_MAX_READS = 15;

/**
 * Reads the files (ids) that need it into the cache. kind: { cacheKey, read(id, previous) }, where
 * read returns an entry (any object) or throws; previous is the file's last entry. changes: null to
 * read only files never read or marked stale; an object (shared by a run's calls, so Drive is asked
 * once) to also read those changed, those not read for an hour and those with an error.
 * Returns the number of files read. At most LINKED_MAX_READS; the next run takes the rest.
 */
function refreshLinkedFiles_(kind, fileIds, changes) {
  const cache = readLinkedCache_(kind);
  const now = Date.now();
  const ids = Array.from(new Set(fileIds.filter(Boolean)));
  const changed = changes && ids.length
    ? spreadsheetsChangedSince_((cache.checkedAt || now - LINKED_REREAD_MS) - LINKED_LOOK_BACK_MS, changes) : new Set();

  // The most urgent first: 0 new or changed, 1 not read for an hour, 2 failed last time.
  const urgency = function (id) {
    const known = cache.files[id];
    if (!known || known.stale || changed.has(id)) return 0;
    if (!changes) return -1;
    if (known.error) return 2;
    return now - known.at < LINKED_REREAD_MS ? -1 : 1;
  };
  const toRead = ids.map(function (id) { return { id: id, urgency: urgency(id) }; })
    .filter(function (f) { return f.urgency >= 0; })
    .sort(function (a, b) { return a.urgency - b.urgency; })
    .slice(0, LINKED_MAX_READS);

  const files = {};
  // All rows were given: forget files no longer linked. Otherwise the others are still needed.
  Object.keys(cache.files).forEach(function (id) {
    if (!changes || ids.indexOf(id) !== -1) files[id] = cache.files[id];
  });
  toRead.forEach(function (f) {
    const previous = files[f.id];
    try {
      files[f.id] = Object.assign({}, kind.read(f.id, previous), { at: now });
    } catch (e) {
      console.log('Could not read körschema ' + f.id + ': ' + e.message);
      // Too many reads: nothing is wrong with the file, so it only waits.
      if (/quota/i.test(e.message) && previous) return;
      // The last known values are kept.
      files[f.id] = Object.assign({}, previous, { at: now, stale: false, error: String(e.message).slice(0, 200) });
    }
  });

  cache.files = files;
  if (changes) cache.checkedAt = now;
  writeLinkedCache_(kind, cache);
  return toRead.length;
}

/** { files: { <file ID>: entry with at (ms), and stale or error }, checkedAt (ms, absent at first) }. */
function readLinkedCache_(kind) {
  const text = CacheService.getScriptCache().get(kind.cacheKey);
  return text ? JSON.parse(text) : { files: {} };
}

function writeLinkedCache_(kind, cache) {
  // The longest CacheService keeps it; without it, everything is simply read again.
  CacheService.getScriptCache().put(kind.cacheKey, JSON.stringify(cache), 21600);
}

/** Marks a file as changed by this script, so that its old values are not used. */
function markLinkedFileStale_(kind, fileId) {
  const cache = readLinkedCache_(kind);
  if (!cache.files[fileId]) return;
  cache.files[fileId].stale = true;
  writeLinkedCache_(kind, cache);
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
 * The IDs of the spreadsheets, in any drive the user can see, modified after since (ms). changes
 * keeps the answer for the run's next call.
 */
function spreadsheetsChangedSince_(since, changes) {
  if (changes.ids && changes.since <= since) return changes.ids;
  // A minute more than asked for, so that it also covers the next call.
  changes.since = since - 60000;
  changes.ids = listSpreadsheetsChangedSince_(new Date(changes.since));
  return changes.ids;
}

function listSpreadsheetsChangedSince_(date) {
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

function sameDuration_(a, b) {
  return typeof a === 'number' && Math.abs(a - b) < 0.5 / 86400;
}

function hasLinkedNote_(note, ownNote) {
  return note.split('\n').some(function (line) { return isLinkedNoteLine_(line, ownNote); });
}

function isLinkedNoteLine_(line, ownNote) {
  return line === ownNote || line.indexOf(LINKED_NOTE_PREFIX) === 0;
}

/**
 * The look of a cell: grey with ownNote while linked (with error, a line saying that what could not
 * be read); back to normal when not. note is its present note. Only written when it changes.
 */
function markLinkedCell_(cell, note, linked, ownNote, what, error) {
  const lines = note.split('\n').filter(function (line) { return !isLinkedNoteLine_(line, ownNote); });
  if (linked) {
    lines.push(error
      ? LINKED_NOTE_PREFIX + what + ' kunde inte hämtas (' + error.replace(/\s+/g, ' ') +
        '). Det senaste värdet står kvar.'
      : ownNote);
  }
  const wanted = lines.join('\n').trim();
  if (wanted === note) return;
  const wasLinked = hasLinkedNote_(note, ownNote);
  cell.setNotes([[wanted]]);
  if (wasLinked !== linked) cell.setFontColor(linked ? LINKED_COLOR : null);
}

/** Whether a range covers the column. */
function coversColumn_(range, col) {
  return range.getColumn() <= col && range.getColumn() + range.getNumColumns() - 1 >= col;
}
