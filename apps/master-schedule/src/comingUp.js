// The "coming up" feed: reads Kalender, picks each show's image and publishes the upcoming shows
// to Cachearoo for the graphic; also the missing-image warnings in Kalender.

const BUCKET = 'sesport';
const KEY = 'comingup';

// at most this many upcoming shows are published, soonest first
const MAX_SHOWS = 10;
// a show's image is its episode image, else its series image, else its genre's: Genre/<genre>.png
// in the Drive folder, uploaded under ELMO_FOLDER ("images/Genre/Handboll.png" for the genre Handboll)
const GENRE_FOLDER = 'Genre';
const GENRE_EXT = '.png';
// the schedule column where a row is warned about a missing genre image, and the warning's fill
const WARNING_HEADER = 'Warning';
const WARNING_COLOR = '#fff2cc';

// Writes the schedule to Cachearoo.
function publishComingUp() {
  const base = requiredProperty_('ELMO_CACHEAROO_URL');
  const apiKey = requiredProperty_('ELMO_KEY');
  const project = requiredProperty_('ELMO_PROJECT');
  const value = buildComingUp(readSchedule(), uploadedImages(), project, new Date());
  markMissingGenreImages();
  cachearooWrite_(base, apiKey, BUCKET, KEY, value);
  toast_(`${value.shows.length} shows published`, 'Elmo');
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
