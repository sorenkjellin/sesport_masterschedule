/**
 * Generic building blocks shared by all seSport Apps Script apps: Script Properties, HTTP,
 * Cachearoo, Elmo's file API, mail, Drive (including a folder watcher), spreadsheets, dates, UI
 * and triggers.
 *
 * scripts/push.sh copies this file into each app when it is deployed, so it runs as part of the
 * app: Script Properties, locks and triggers are the app's own.
 *
 * Nothing here knows about a particular app's sheets, columns or settings; everything a function
 * needs comes in as an argument. Keep it that way so this file rarely has to change. A change here
 * reaches an app the next time that app is pushed.
 *
 * Only functions are defined here (no top-level constants), so it does not matter in which
 * order Apps Script loads the project's files. Names end in "_" so they stay out of the editor's
 * Run menu and cannot be called from menus or the browser.
 */

// --- Script Properties -------------------------------------------------------

/** A Script Property's value; throws when it is not set. */
function requiredProperty_(name) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) throw new Error(`The script property ${name} is not set`);
  return value;
}

/** A Script Property's value, or fallback when it is not set. */
function optionalProperty_(name, fallback) {
  return PropertiesService.getScriptProperties().getProperty(name) || fallback;
}

/** "/images/" -> "images" */
function trimSlashes_(path) {
  return String(path).replace(/^\/+|\/+$/g, '');
}

// --- HTTP --------------------------------------------------------------------

/** Throws unless the UrlFetchApp response is 2xx. `what` names the request in the error. */
function checkResponse_(res, what) {
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(what + ' failed (' + code + '): ' + res.getContentText().slice(0, 300));
  }
  return res;
}

// --- Cachearoo ---------------------------------------------------------------

/**
 * The URL of a value in Cachearoo: <base>/_data/<bucket>/<key>.
 * A base that already points at a value (contains /_data/) is used as it is.
 */
function cachearooUrl_(base, bucket, key) {
  if (/\/_data\//.test(base)) return base;
  return `${base.replace(/\/+$/, '')}/_data/${encodeURIComponent(bucket)}/${encodeURIComponent(key)}`;
}

/** Writes value, as JSON, to bucket/key in the Cachearoo at base. */
function cachearooWrite_(base, apiKey, bucket, key, value) {
  checkResponse_(UrlFetchApp.fetch(cachearooUrl_(base, bucket, key), {
    method: 'put',
    contentType: 'application/json',
    headers: { 'API-Key': apiKey, 'x-api-key': apiKey },
    payload: JSON.stringify(value),
    muteHttpExceptions: true,
  }), 'Cachearoo write');
}

// --- Elmo file API -----------------------------------------------------------
// `elmo` is { elmoUrl, key, project }: the Elmo origin (no trailing slash), a write API key,
// and the project the files live in.

/** Every file in the project: path -> { hash, url }. The CDN URL ends in the SHA-256. */
function getElmoFiles_(elmo) {
  const res = UrlFetchApp.fetch(
    elmo.elmoUrl + '/api/files/projects/' + encodeURIComponent(elmo.project) + '/?recurse=1',
    { headers: { 'x-api-key': elmo.key }, muteHttpExceptions: true });
  if (res.getResponseCode() === 404) return {};
  checkResponse_(res, 'list project');

  const files = {};
  JSON.parse(res.getContentText()).list.forEach(function (item) {
    if (!item.folder && item.url) files[item.name] = { hash: item.url.split('/').pop(), url: item.url };
  });
  return files;
}

/** Registers the project so it shows in Elmo's project list. No-op if it exists. */
function ensureElmoProject_(elmo) {
  checkResponse_(UrlFetchApp.fetch(
    elmo.elmoUrl + '/api/files/projects/' + encodeURIComponent(elmo.project),
    { method: 'put', headers: { 'x-api-key': elmo.key }, muteHttpExceptions: true }), 'create project');
}

/** Uploads blob to path in the project: upload session -> PUT to S3 -> complete. Returns the CDN URL. */
function uploadToElmo_(elmo, blob, path) {
  const contentType = blob.getContentType() || 'application/octet-stream';
  const meta = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': elmo.key },
    payload: JSON.stringify({ project: elmo.project, path: path, contentType: contentType }),
    muteHttpExceptions: true,
  };
  const sessionRes = UrlFetchApp.fetch(elmo.elmoUrl + '/api/files/upload-session', meta);
  checkResponse_(sessionRes, 'open upload session');
  const session = JSON.parse(sessionRes.getContentText());

  // Content-Type is part of the presigned signature, so it must match. No API key to S3.
  checkResponse_(UrlFetchApp.fetch(session.url, {
    method: 'put', contentType: contentType, payload: blob.getBytes(), muteHttpExceptions: true,
  }), 'upload to S3');

  const doneRes = UrlFetchApp.fetch(
    elmo.elmoUrl + '/api/files/upload-session/' + session.id + '/complete', meta);
  checkResponse_(doneRes, 'complete upload');
  return JSON.parse(doneRes.getContentText()).url;
}

/**
 * A path in the project ("images/Genre/Handboll.png") as Elmo's file API serves it, relative to
 * the origin Elmo runs on. A full URL is returned as it is; '' for no path.
 */
function elmoFileUrl_(path, project) {
  if (!path) return '';
  if (/^https?:\/\//i.test(path)) return path;
  const segments = path.replace(/^\/+/, '').split('/').map(encodeURIComponent);
  return `/api/files/projects/${encodeURIComponent(project)}/${segments.join('/')}`;
}

// --- Mail --------------------------------------------------------------------

/**
 * Sends a plain-text email from the account running the script.
 * options (all optional): cc, bcc, name (sender name), replyTo, attachments (blobs), htmlBody.
 * Empty options are left out, so an unset cc is simply not sent.
 */
function sendMail_(to, subject, body, options) {
  if (!to) throw new Error(`No recipient for the email "${subject}"`);
  const opts = {};
  Object.keys(options || {}).forEach(function (k) {
    if (options[k] != null && options[k] !== '') opts[k] = options[k];
  });
  GmailApp.sendEmail(to, subject, body, opts);
}

// --- Spreadsheets ------------------------------------------------------------

/** The named sheet (created if missing) with headers written in bold in row 1, which is frozen. */
function getOrCreateSheet_(name, headers, spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** The rows below the header, as displayed, in the first `width` columns. */
function readTextRows_(sheet, width) {
  const n = sheet.getLastRow() - 1;
  if (n <= 0) return [];
  return sheet.getRange(2, 1, n, width).getDisplayValues();
}

// Plain-text format so timestamps and hashes are kept exactly as written
// (Sheets would otherwise turn ISO dates into date values).

/** Replaces the rows below the header, in the first `width` columns only, with rows as plain text. */
function writeTextRows_(sheet, rows, width) {
  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, width).clearContent();
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, width).setNumberFormat('@').setValues(rows);
  }
}

/** Writes one row as plain text; index 0 is the first row below the header. */
function writeTextRow_(sheet, index, row) {
  sheet.getRange(index + 2, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}

/**
 * Each column's index in headers (-1 when missing). columns maps a key to the header names it
 * may have, e.g. { series: ['Programtitel', 'Progamtitel'] }; the first that is present wins.
 * Matching ignores case and surrounding spaces.
 */
function findColumns_(headers, columns) {
  const names = headers.map(function (h) { return String(h).trim().toLowerCase(); });
  const cols = {};
  Object.keys(columns).forEach(function (key) {
    const name = columns[key].map(function (n) { return n.toLowerCase(); })
      .find(function (n) { return names.indexOf(n) >= 0; });
    cols[key] = name ? names.indexOf(name) : -1;
  });
  return cols;
}

/**
 * values (a 2D array) as an .xlsx blob named fileName, with one sheet named sheetName.
 * Goes through a temporary spreadsheet, which is trashed afterwards.
 */
function valuesToXlsx_(values, sheetName, fileName) {
  if (!values.length || !values[0].length) throw new Error(`Nothing to export to ${fileName}`);
  const temp = SpreadsheetApp.create('Temp export ' + fileName);
  try {
    const sheet = temp.getSheets()[0];
    sheet.setName(sheetName);
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    sheet.autoResizeColumns(1, values[0].length);
    SpreadsheetApp.flush();

    const res = UrlFetchApp.fetch(
      `https://docs.google.com/spreadsheets/d/${temp.getId()}/export?format=xlsx`,
      { headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` }, muteHttpExceptions: true });
    checkResponse_(res, 'export ' + fileName);
    return res.getBlob().setName(fileName);
  } finally {
    DriveApp.getFileById(temp.getId()).setTrashed(true);
  }
}

/** Resolves either a Range object or an A1 notation string (on the active sheet) to a Range. */
function resolveCellRange_(cell) {
  if (typeof cell === 'string') {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange(cell);
  }
  if (cell && typeof cell.getA1Notation === 'function') return cell;
  throw new Error('Ange en giltig cellreferens eller ett Range-objekt.');
}

// --- Smart chips -------------------------------------------------------------

/** The URI of the first file smart chip in range, or null. Uses the Sheets advanced service. */
function getSmartChipFileUrl(range) {
  const sheet = range.getSheet();
  const fullA1Range = "'" + sheet.getName().replace(/'/g, "''") + "'!" + range.getA1Notation();
  const response = Sheets.Spreadsheets.get(sheet.getParent().getId(), {
    ranges: [fullA1Range],
    includeGridData: true,
    fields: 'sheets(data(rowData(values(chipRuns))))',
  });

  for (const s of response.sheets || []) {
    for (const block of s.data || []) {
      for (const row of block.rowData || []) {
        for (const cell of row.values || []) {
          for (const run of cell.chipRuns || []) {
            const chip = run.chip;
            if (chip && chip.richLinkProperties && chip.richLinkProperties.uri) {
              return chip.richLinkProperties.uri;
            }
          }
        }
      }
    }
  }
  Logger.log('No file smart chip found in ' + fullA1Range);
  return null;
}

/** The Drive file behind the file smart chip in cell (a Range or A1 string); throws if there is none. */
function getFileChipFile_(cell) {
  const range = resolveCellRange_(cell);
  const fileUrl = getSmartChipFileUrl(range);
  if (!fileUrl) throw new Error('Ingen fil-smart chip hittades i ' + range.getA1Notation());
  const fileId = extractDriveId(fileUrl);
  if (!fileId) throw new Error('Kunde inte hitta ett giltigt Drive-fil-ID i länken.');
  return DriveApp.getFileById(fileId);
}

// --- Drive -------------------------------------------------------------------

/** The Drive ID in a Google Docs / Drive URL (".../d/<id>/..." or "...?id=<id>"), or null. */
function extractDriveId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * Every file in folder and its subfolders whose MIME type starts with mimePrefix
 * ('image/', 'video/', '' for all), as { id, name, relPath, modified (ISO), size }.
 * relPath is the path below folder; a "/" inside a file or folder name becomes "-".
 */
function listDriveFiles_(folder, mimePrefix, prefix, out) {
  prefix = prefix || '';
  out = out || [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().indexOf(mimePrefix || '') !== 0) continue;
    out.push({
      id: f.getId(),
      name: f.getName(),
      relPath: prefix + f.getName().replace(/\//g, '-'),
      modified: f.getLastUpdated().toISOString(),
      size: f.getSize(),
    });
  }
  const subs = folder.getFolders();
  while (subs.hasNext()) {
    const sub = subs.next();
    listDriveFiles_(sub, mimePrefix, prefix + sub.getName().replace(/\//g, '-') + '/', out);
  }
  return out;
}

/** A folder named name next to folder (in one of its parents), or null. */
function findSiblingFolder_(folder, name) {
  const parents = folder.getParents();
  while (parents.hasNext()) {
    const siblings = parents.next().getFoldersByName(name);
    if (siblings.hasNext()) return siblings.next();
  }
  return null;
}

/**
 * The folder one level above the folder that holds file (its "grandparent").
 * Falls back to My Drive's root when there is no such folder.
 */
function getGrandparentFolder_(file) {
  const parents = file.getParents();
  if (parents.hasNext()) {
    const grandParents = parents.next().getParents();
    if (grandParents.hasNext()) return grandParents.next();
  }
  return DriveApp.getRootFolder();
}

/**
 * Each video's duration in milliseconds by Drive file ID; null when Drive has not processed the
 * video yet or cannot read its format. Asks the Drive REST API (DriveApp has no duration) in
 * batches of 50.
 */
function driveVideoDurationsMs_(ids) {
  const token = ScriptApp.getOAuthToken();
  const out = {};
  for (let i = 0; i < ids.length; i += 50) {
    const batch = ids.slice(i, i + 50);
    const responses = UrlFetchApp.fetchAll(batch.map(function (id) {
      return {
        url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) +
          '?fields=videoMediaMetadata(durationMillis)&supportsAllDrives=true',
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true,
      };
    }));
    responses.forEach(function (res, j) {
      let ms = null;
      if (res.getResponseCode() === 200) {
        const meta = JSON.parse(res.getContentText()).videoMediaMetadata;
        if (meta && meta.durationMillis != null) ms = Number(meta.durationMillis);
      } else {
        Logger.log('Duration of ' + batch[j] + ' failed (' + res.getResponseCode() + ')');
      }
      out[batch[j]] = ms;
    });
  }
  return out;
}

// --- Drive folder watcher ----------------------------------------------------

/**
 * Emails a list of the files created in a Drive folder (and its direct subfolders) since the last
 * check. Run it from a time-driven trigger. Options:
 *   folderId       the folder to watch
 *   recipients     email address(es), comma-separated
 *   subjectPrefix  the start of the subject (optional)
 *   stateKey       the Script Property that holds the time of the last check (optional; give each
 *                  watcher in an app its own)
 *   firstLookbackMs  how far back the first check looks (optional, default one hour)
 * Returns the new files found.
 */
function watchDriveFolder_(options) {
  const props = PropertiesService.getScriptProperties();
  const stateKey = options.stateKey || 'DRIVE_WATCH_LAST_CHECK_' + options.folderId;
  const now = new Date();
  const last = props.getProperty(stateKey);
  const since = last ? new Date(last) : new Date(now.getTime() - (options.firstLookbackMs || 3600000));

  const folder = DriveApp.getFolderById(options.folderId);
  const newFiles = findNewDriveFiles_(folder, since);
  if (newFiles.length) {
    sendMail_(options.recipients, newFilesSubject_(options.subjectPrefix, newFiles),
      newFilesEmailBody_(folder.getName(), newFiles));
    Logger.log(`Notification sent for ${newFiles.length} file(s) in ${folder.getName()}`);
  }
  props.setProperty(stateKey, now.toISOString());
  return newFiles;
}

/**
 * The files created after since in folder and its direct subfolders (not deeper), as
 * { folderName, folderUrl, name, url, sizeMb, type, created }. Files directly in folder get the
 * folder name "Root / Parent Folder".
 */
function findNewDriveFiles_(folder, since) {
  const out = [];
  const scan = function (f, label) {
    const files = f.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const created = file.getDateCreated();
      if (created.getTime() <= since.getTime()) continue;
      out.push({
        folderName: label,
        folderUrl: f.getUrl(),
        name: file.getName(),
        url: file.getUrl(),
        sizeMb: (file.getSize() / (1024 * 1024)).toFixed(2),
        type: file.getMimeType(),
        created: Utilities.formatDate(created, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      });
    }
  };
  scan(folder, 'Root / Parent Folder');
  const subfolders = folder.getFolders();
  while (subfolders.hasNext()) {
    const sub = subfolders.next();
    scan(sub, sub.getName());
  }
  return out;
}

function groupByFolder_(files) {
  const grouped = {};
  files.forEach(function (file) {
    (grouped[file.folderName] = grouped[file.folderName] || []).push(file);
  });
  return grouped;
}

function newFilesSubject_(prefix, files) {
  const folders = Object.keys(groupByFolder_(files)).join(', ');
  return `${prefix || 'New files'}: ${files.length} new asset(s) in [${folders}]`;
}

/** A plain-text email body listing files (from findNewDriveFiles_), grouped by subfolder. */
function newFilesEmailBody_(parentFolderName, files) {
  let body = `Hello,\n\n${files.length} new asset(s) were uploaded across the monitored folders in "${parentFolderName}".\n\n`;
  body += `==============================================\n\n`;
  for (const [folderName, group] of Object.entries(groupByFolder_(files))) {
    body += `📁 SUBFOLDER: ${folderName} (${group.length} new asset${group.length > 1 ? 's' : ''})\n`;
    body += `Folder link: ${group[0].folderUrl}\n`;
    body += `----------------------------------------------\n`;
    group.forEach(function (file, idx) {
      body += `  ${idx + 1}. ${file.name}\n`;
      body += `     • Size: ${file.sizeMb} MB\n`;
      body += `     • Type: ${file.type}\n`;
      body += `     • Uploaded: ${file.created}\n`;
      body += `     • File link: ${file.url}\n\n`;
    });
  }
  body += `==============================================\n`;
  body += `Best regards,\nAutomated Asset Watcher`;
  return body;
}

// --- Hashing, time and dates -------------------------------------------------

/** The SHA-256 of a byte array, as lowercase hex. */
function sha256Hex_(bytes) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); })
    .join('');
}

/** Milliseconds as a timecode H:MM:SS:FF at fps frames per second (default 25); partial frames dropped. */
function formatTimecode_(ms, fps) {
  fps = fps || 25;
  const pad = function (n) { return String(n).padStart(2, '0'); };
  const h = Math.floor(ms / 3600000);
  const m = Math.floor(ms / 60000) % 60;
  const s = Math.floor(ms / 1000) % 60;
  const frames = Math.floor((ms % 1000) * fps / 1000);
  return h + ':' + pad(m) + ':' + pad(s) + ':' + pad(frames);
}

/**
 * A date ("2026-11-07" or "7.11.2026") and a time ("17.50" or "17:50") in timeZone as a Date;
 * null when either cannot be read.
 */
function parseLocalDateTime_(date, time, timeZone) {
  const iso = date.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  const dmy = date.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  const t = time.match(/^(\d{1,2})[:.](\d{2})/);
  if (!(iso || dmy) || !t) return null;
  const [y, m, d] = iso ? [iso[1], iso[2], iso[3]] : [dmy[3], dmy[2], dmy[1]];
  const text = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')} ${t[1].padStart(2, '0')}:${t[2]}`;
  return Utilities.parseDate(text, timeZone, 'yyyy-MM-dd HH:mm');
}

// --- UI ----------------------------------------------------------------------

/** Logs msg and shows it as a toast; the toast is skipped where there is no UI (triggers). */
function toast_(msg, title, seconds) {
  Logger.log(msg);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, title || '', seconds || 5);
  } catch (e) { /* no UI in triggers */ }
}

/** text with &, <, >, " and ' escaped for HTML. */
function escapeHtml_(text) {
  return String(text).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/** Shows text, preformatted, in a modal dialog. */
function showTextDialog_(title, text, width, height) {
  const html = HtmlService.createHtmlOutput(`<pre>${escapeHtml_(text)}</pre>`)
    .setWidth(width || 700).setHeight(height || 500);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}

/**
 * Opens url in a new browser tab through a small dialog (with a link in case pop-ups are blocked).
 * title names what is opened; dialogTitle is the dialog's own title.
 */
function openUrlInBrowser_(url, title, dialogTitle) {
  const html = HtmlService.createHtmlOutput(
    `<!DOCTYPE html>
     <html>
       <head>
         <base target="_blank">
         <script>
           window.open(${JSON.stringify(url)}, '_blank');
           setTimeout(() => google.script.host.close(), 1500);
         </script>
         <style>
           body { font-family: Roboto, Arial, sans-serif; padding: 15px; font-size: 13px; }
           a { color: #1a73e8; text-decoration: none; font-weight: bold; }
         </style>
       </head>
       <body>
         <p>Opening <strong>${escapeHtml_(title)}</strong>...</p>
         <p>If the window didn't open automatically, <a href="${escapeHtml_(url)}">click here to open</a>.</p>
       </body>
     </html>`
  ).setWidth(350).setHeight(120);
  SpreadsheetApp.getUi().showModalDialog(html, dialogTitle || 'Opening');
}

// --- Triggers ----------------------------------------------------------------

/** Deletes every trigger that runs handler (a function name), in any spreadsheet. */
function deleteTriggers_(handler) {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === handler) ScriptApp.deleteTrigger(trigger);
  });
}

/** Adds an installable onEdit trigger running handler on spreadsheet, unless there already is one. */
function ensureEditTrigger_(handler, spreadsheet) {
  const id = spreadsheet.getId();
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === handler && trigger.getTriggerSourceId() === id;
  });
  if (!exists) ScriptApp.newTrigger(handler).forSpreadsheet(spreadsheet).onEdit().create();
}
