// Data -> Förhandsgranska EPG: the EPG tab as the third parties will show it, a card per show by
// day, with warnings for what looks wrong. Only reads; its Skicka button runs exportEPGAndEmail
// (epgExport.js). The window is epgPreviewDialog.html.

const EPG_SHEET = 'EPG';
// A header matches ignoring case and surrounding spaces.
const EPG_COLUMNS = {
  date: ['Date'],
  start: ['Start time'],
  end: ['End time'],
  title: ['Title'],
  league: ['League'],
  episode: ['Episode title'],
  live: ['Live Rerun', 'Live/Rerun'],
  category: ['Desc.'], // "Sport": a category, not a description
  episodeDescription: ['Episode description'],
  seriesImage: ['Series Image'],
  episodeImage: ['Episode image'],
  genre: ['Genre'],
  seriesPortrait: ['Series Img 2:3'],
  episodePortrait: ['Episode Img 2:3'],
  cast: ['Cast/Credits'],
  role: ['Cast/credits role'],
};
const EPG_LONG_SHOW_HOURS = 8;

function previewEpg() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EPG_SHEET);
  if (!sheet) throw new Error('Fliken ' + EPG_SHEET + ' saknas');
  const template = HtmlService.createTemplateFromFile('epgPreviewDialog');
  // "<" escaped, so that no text in the sheet can end the page's script.
  template.data = JSON.stringify(buildEpgPreview_(sheet.getDataRange().getDisplayValues()))
    .replace(/</g, '\\u003c');
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setWidth(1100).setHeight(760), 'Förhandsgranska EPG');
}

/** Called by the window's Skicka button. Returns the recipient. */
function sendEpgFromPreview() {
  exportEPGAndEmail();
  return requiredProperty_('EPG_MAIL_ENDPOINT');
}

/**
 * The EPG tab's rows (as displayed, header first) as { days: [{ date ('yyyy-MM-dd' or the cell's
 * text), shows }], missing (headers not found), count, warnings (number of shows with any) }. A
 * show: { start, end, endFromNext, length, genre, category, title, subtitle, series, description,
 * image, portrait, live, cast, role, warnings }.
 *
 * A missing end time is an error. The show is then shown until the next one starts, as the
 * receiver (Simply) does.
 */
function buildEpgPreview_(values) {
  const headers = values[0] || [];
  const cols = findColumns_(headers, EPG_COLUMNS);
  const missing = ['date', 'start', 'end', 'title', 'episodeImage', 'seriesImage']
    .filter(function (k) { return cols[k] < 0; }).map(function (k) { return EPG_COLUMNS[k][0]; });
  const get = function (row, key) { return cols[key] < 0 ? '' : String(row[cols[key]] || '').trim(); };

  const shows = values.slice(1).filter(function (row) {
    return ['date', 'start', 'title', 'episode'].some(function (k) { return get(row, k); });
  }).map(function (row) {
    const warnings = [];
    const start = parseLocalDateTime_(get(row, 'date'), get(row, 'start'), TIME_ZONE);
    let end = start && get(row, 'end') ? parseLocalDateTime_(get(row, 'date'), get(row, 'end'), TIME_ZONE) : null;
    if (!start) warnings.push('Datum eller starttid kan inte läsas');
    else if (!get(row, 'end')) warnings.push('Sluttid saknas');
    else if (!end) warnings.push('Sluttiden kan inte läsas');
    // A show that ends after midnight.
    if (start && end && end <= start) end = new Date(end.getTime() + 24 * 3600 * 1000);

    const image = get(row, 'episodeImage') || get(row, 'seriesImage');
    if (!image) warnings.push('Bild saknas');
    else if (!/^https:\/\//i.test(image)) warnings.push('Bilden är ingen webbadress: ' + image);
    const portrait = get(row, 'episodePortrait') || get(row, 'seriesPortrait');
    if (portrait && !/^https:\/\//i.test(portrait)) warnings.push('2:3-bilden är ingen webbadress: ' + portrait);
    const description = get(row, 'episodeDescription');
    if (!description) warnings.push('Beskrivning saknas');

    const episode = get(row, 'episode');
    return {
      at: start ? start.getTime() : null, until: end ? end.getTime() : null,
      date: start ? Utilities.formatDate(start, TIME_ZONE, 'yyyy-MM-dd') : get(row, 'date'),
      start: get(row, 'start'), end: get(row, 'end'), endFromNext: false, length: '',
      genre: get(row, 'genre'), category: get(row, 'category'),
      title: episode || get(row, 'title'), subtitle: get(row, 'league'), series: episode ? get(row, 'title') : '',
      description: description, image: /^https:\/\//i.test(image) ? image : '',
      portrait: /^https:\/\//i.test(portrait) ? portrait : '', live: get(row, 'live'),
      cast: get(row, 'cast'), role: get(row, 'role'), warnings: warnings,
    };
  });

  // In time order (rows whose time cannot be read last, as they are), and overlaps marked.
  shows.sort(function (a, b) {
    if (a.at === null || b.at === null) return (a.at === null) - (b.at === null);
    return a.at - b.at;
  });
  const timed = shows.filter(function (s) { return s.at !== null; });
  timed.forEach(function (s, i) {
    const next = timed[i + 1];
    if (s.until === null && next) {
      s.until = next.at;
      s.end = Utilities.formatDate(new Date(next.at), TIME_ZONE, 'HH:mm');
      s.endFromNext = true;
    }
    if (s.until === null) return;
    const minutes = Math.round((s.until - s.at) / 60000);
    s.length = epgLength_(minutes);
    if (s.endFromNext) {
      s.warnings[s.warnings.indexOf('Sluttid saknas')] =
        'Sluttid saknas. Simply kommer att sätta sluttiden till starttiden för nästa program, vilket gör att ' +
        'programmet ser ut att vara ' + s.length + ' långt.';
    } else if (minutes > EPG_LONG_SHOW_HOURS * 60) {
      s.warnings.push('Längre än ' + EPG_LONG_SHOW_HOURS + ' timmar: stämmer sluttiden?');
    }
    if (next && next.at < s.until) {
      next.warnings.push('Börjar innan ' + s.title + ' har slutat (' + s.end + ')');
    }
  });

  const days = [];
  shows.forEach(function (s) {
    const day = days[days.length - 1];
    if (day && day.date === s.date) day.shows.push(s); else days.push({ date: s.date, shows: [s] });
    delete s.at;
    delete s.until;
  });
  return { days: days, missing: missing, count: shows.length,
           warnings: shows.filter(function (s) { return s.warnings.length; }).length };
}

/** 125 -> "2 h 5 min", 45 -> "45 min", 120 -> "2 h". */
function epgLength_(minutes) {
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  return [h ? h + ' h' : '', m || !h ? m + ' min' : ''].filter(Boolean).join(' ');
}
