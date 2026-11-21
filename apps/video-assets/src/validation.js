// Video validation: which of the profiles in "Video Profiles" each probed video in "Video Inbox"
// meets.
//
// "Video Profiles" has a row per profile: its name in column A, then rules under headers that name
// Video Inbox columns. "Min <column>" and "Max <column>" are limits (Max duration, Min MBit, Max
// Mbit); other headers must match a Video Inbox header (Frame rate, Codec, Audio …). Upper and
// lower case do not matter. A rule is a value, or values separated by "/" ("8/10",
// "none / stereo (2 ch)"); in text, * stands for any text ("ProRes 422*" allows ProRes 422, 422 HQ,
// 422 LT …); an empty cell allows anything. Profiles can be added at any time.
//
// In "Video Inbox", for each probed video:
//   approved for one or more profiles: a note on its name, "Approved for:" and the profiles, one
//   per line;
//   approved for none: the row is greyed out, and the cells that make it fail are light red, with
//   a note saying why:
//     a rule all profiles share ("global"):      "Frame rate should be 50"
//     a codec no profile allows:                 "DNxHD is not an allowed codec."
//     what the profiles with its codec require:  "Max duration is 0:01:30 when using ProRes 422"
//   When each property is allowed by some profile with the codec, but no profile allows them all,
//   the properties the nearest profile (fewest failures) does not allow are marked, naming it.
//
// Runs when Video Inbox is rebuilt, when a profile is edited (onEdit), on every run of the asset
// watcher (which picks up new probe results within its 10 minutes), and from Inkorg → Validera videor.

const PROFILES_SHEET = 'Video Profiles';
const VALIDATION_COLORS = { greyBackground: '#efefef', greyText: '#666666', red: '#f4cccc' };

/** Menu: Inkorg → Validera videor. */
function validateVideoInbox() {
  const r = validateVideoInbox_();
  toast_(r.skipped ? r.skipped : `${r.approved} godkända, ${r.failed} underkända, ${r.unprobed} ej analyserade`, 'Validering');
}

/**
 * Validates every probed video in Video Inbox against Video Profiles and marks the results.
 * Returns { approved, failed, unprobed } counts, or { skipped: why } when there is nothing to do.
 */
function validateVideoInbox_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inbox = ss.getSheetByName('Video Inbox');
  const profileSheet = ss.getSheetByName(PROFILES_SHEET);
  if (!inbox || inbox.getLastRow() < 2) return { skipped: 'Video Inbox är tom' };
  if (!profileSheet) return { skipped: `Ingen flik "${PROFILES_SHEET}"` };

  const width = inbox.getLastColumn();
  const n = inbox.getLastRow() - 1;
  const headers = inbox.getRange(1, 1, 1, width).getDisplayValues()[0].map(function (h) { return String(h).trim(); });
  const profiles = readProfiles_(profileSheet, headers);
  const cells = readCells_(inbox.getRange(2, 1, n, width));
  const nameCol = headers.indexOf('Name');
  const statusCol = headers.indexOf('Probe status');

  const backgrounds = [];
  const fontColors = [];
  const notes = [];
  const counts = { approved: 0, failed: 0, unprobed: 0 };
  cells.forEach(function (row) {
    const bg = row.map(function () { return null; });
    const fc = row.map(function () { return null; });
    const nt = row.map(function () { return ''; });
    backgrounds.push(bg); fontColors.push(fc); notes.push(nt);
    if (!/^Probed/.test(row[statusCol].text) || !profiles.list.length) { counts.unprobed++; return; }
    const result = validateAsset_(row, profiles);
    if (result.approved.length) {
      counts.approved++;
      if (nameCol >= 0) nt[nameCol] = 'Approved for:\n' + result.approved.join('\n');
      return;
    }
    counts.failed++;
    bg.fill(VALIDATION_COLORS.greyBackground);
    fc.fill(VALIDATION_COLORS.greyText);
    result.problems.forEach(function (p) {
      bg[p.col] = VALIDATION_COLORS.red;
      nt[p.col] = nt[p.col] ? nt[p.col] + '\n' + p.note : p.note;
    });
  });
  const range = inbox.getRange(2, 1, n, width);
  range.setBackgrounds(backgrounds).setFontColors(fontColors).setNotes(notes);
  return counts;
}

/**
 * Each cell of a range as { value, text }: value is a number for a number cell (a duration as its
 * seconds, read from the shown h:mm:ss, since Apps Script reads duration cells as dates), else the
 * shown text; text is the shown text.
 */
function readCells_(range) {
  const values = range.getValues();
  const shown = range.getDisplayValues();
  return shown.map(function (line, r) {
    return line.map(function (text, c) {
      return { value: cellValue_(values[r][c], text), text: String(text).trim() };
    });
  });
}

function cellValue_(value, text) {
  const days = parseDuration_(text);
  if (days !== '') return Math.round(days * 86400);
  if (typeof value === 'number') return value;
  const number = parseNumber_(text);
  return number === null ? String(text).trim() : number;
}

/** "29,97" / "1 920" / "50" as a number; null when the text is not one. */
function parseNumber_(text) {
  const t = String(text).replace(/[\s  ]/g, '').replace(',', '.');
  return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : null;
}

/**
 * The profiles of the Video Profiles sheet: { columns: [{ header, col (in Video Inbox), bound
 * ('min', 'max' or null) }], list: [{ name, rules: [[allowed values] or null per column] }],
 * unknown: [headers that name no Video Inbox column] }.
 */
function readProfiles_(sheet, inboxHeaders) {
  const lower = inboxHeaders.map(function (h) { return h.toLowerCase(); });
  const width = sheet.getLastColumn();
  const n = sheet.getLastRow() - 1;
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(function (h) { return String(h).trim(); });
  const columns = [];
  const unknown = [];
  headers.forEach(function (h, i) {
    if (i === 0 || !h) return; // column A: the profile's name
    const m = /^(min|max)\s+(.+)$/i.exec(h);
    const target = (m ? m[2] : h).toLowerCase();
    const col = lower.indexOf(target);
    if (col < 0) { unknown.push(h); return; }
    columns.push({ header: h, col: col, at: i, bound: m ? m[1].toLowerCase() : null });
  });
  const list = [];
  if (n > 0) {
    readCells_(sheet.getRange(2, 1, n, width)).forEach(function (row) {
      if (!row[0].text) return;
      list.push({
        name: row[0].text,
        rules: columns.map(function (c) {
          const cell = row[c.at];
          if (!cell.text) return null;
          // "8/10" -> [8, 10]; a single cell value (a number, a duration) as it is.
          const parts = cell.text.split('/').map(function (p) { return p.trim(); }).filter(Boolean);
          return parts.length > 1 ? parts.map(function (p) { return cellValue_('', p); }) : [cell.value];
        }),
      });
    });
  }
  return { columns: columns, list: list, unknown: unknown };
}

/** Whether value meets a rule (allowed values, or null for anything) of a column with bound. */
function meetsRule_(value, rule, bound) {
  if (!rule) return true;
  if (value === '' || value == null) return false;
  if (bound) {
    if (typeof value !== 'number' || typeof rule[0] !== 'number') return false;
    return bound === 'min' ? value >= rule[0] : value <= rule[0];
  }
  return rule.some(function (allowed) {
    if (typeof allowed === 'number' && typeof value === 'number') return Math.abs(allowed - value) < 1e-9;
    const text = String(allowed).toLowerCase();
    if (text.indexOf('*') >= 0) return wildcardRegExp_(text).test(String(value).toLowerCase());
    return text === String(value).toLowerCase();
  });
}

/** "prores 422*" as a pattern: * stands for any text (none included); the rest is literal. */
function wildcardRegExp_(text) {
  return new RegExp('^' + text.split('*').map(function (part) {
    return part.replace(/[.+?^${}()|[\]\\]/g, '\\$&');
  }).join('.*') + '$');
}

/** A rule as written in a note: durations as h:mm:ss, lists with " / ". */
function ruleText_(rule, header) {
  return rule.map(function (v) {
    if (typeof v === 'number' && /duration/i.test(header)) {
      const pad = function (x) { return String(x).padStart(2, '0'); };
      return `${Math.floor(v / 3600)}:${pad(Math.floor(v / 60) % 60)}:${pad(v % 60)}`;
    }
    return String(v);
  }).join(' / ');
}

/**
 * One video (a row from readCells_) against the profiles: { approved: [profile names],
 * problems: [{ col (in Video Inbox), note }] } (problems only when no profile approves it).
 */
function validateAsset_(row, profiles) {
  const cols = profiles.columns;
  const value = function (i) { return row[cols[i].col].value; };
  const failsIn = function (p) {
    return cols.map(function (c, i) { return i; }).filter(function (i) { return !meetsRule_(value(i), p.rules[i], cols[i].bound); });
  };
  const approved = profiles.list.filter(function (p) { return !failsIn(p).length; }).map(function (p) { return p.name; });
  if (approved.length) return { approved: approved, problems: [] };

  const problems = [];
  const same = function (a, b) { return JSON.stringify(a) === JSON.stringify(b); };
  // Global rules: every profile has the same one.
  const global = cols.map(function (c, i) {
    const first = profiles.list[0].rules[i];
    return !!first && profiles.list.every(function (p) { return same(p.rules[i], first); });
  });
  global.forEach(function (isGlobal, i) {
    if (isGlobal && !meetsRule_(value(i), profiles.list[0].rules[i], cols[i].bound)) {
      problems.push({ col: cols[i].col, note: `${cols[i].header} should be ${ruleText_(profiles.list[0].rules[i], cols[i].header)}` });
    }
  });

  // The codec decides which profiles are compared.
  const codecAt = cols.findIndex(function (c) { return !c.bound && /^codec$/i.test(c.header); });
  if (codecAt < 0 || global[codecAt]) return { approved: [], problems: problems };
  const codec = value(codecAt);
  const withCodec = profiles.list.filter(function (p) { return meetsRule_(codec, p.rules[codecAt], null); });
  if (!withCodec.length) {
    problems.push({ col: cols[codecAt].col, note: `${codec} is not an allowed codec.` });
    return { approved: [], problems: problems };
  }

  // A property no profile with this codec allows: what they allow, as one rule (the widest limit,
  // or every allowed value).
  cols.forEach(function (c, i) {
    if (global[i] || i === codecAt) return;
    const rules = withCodec.map(function (p) { return p.rules[i]; });
    if (rules.some(function (r) { return !r; })) return; // one of them allows anything
    if (rules.some(function (r) { return meetsRule_(value(i), r, c.bound); })) return;
    let allowed;
    if (c.bound) {
      const limits = rules.map(function (r) { return r[0]; });
      allowed = [c.bound === 'min' ? Math.min.apply(null, limits) : Math.max.apply(null, limits)];
    } else {
      allowed = [];
      rules.forEach(function (r) { r.forEach(function (v) { if (!allowed.some(function (a) { return same(a, v); })) allowed.push(v); }); });
    }
    problems.push({ col: c.col, note: `${c.header} is ${ruleText_(allowed, c.header)} when using ${codec}` });
  });
  if (problems.length) return { approved: [], problems: problems };

  // Every property is allowed by some profile with the codec, but none allows them all: mark what
  // the nearest one (fewest failures) does not allow.
  const nearest = withCodec.map(function (p) { return { p: p, fails: failsIn(p) }; })
    .sort(function (a, b) { return a.fails.length - b.fails.length; })[0];
  nearest.fails.forEach(function (i) {
    problems.push({ col: cols[i].col,
      note: `${cols[i].header} is ${ruleText_(nearest.p.rules[i], cols[i].header)} when using ${codec} (${nearest.p.name})` });
  });
  return { approved: [], problems: problems };
}
