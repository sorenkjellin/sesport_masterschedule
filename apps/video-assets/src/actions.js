// What to do with a file in an inbox tab: the "Action" column (D) of Video Inbox and Image Inbox.
//
// Its dropdown offers the subfolders of the production folder (Video: VIDEO_FOLDER_ID, Images:
// IMAGES_FOLDER_ID), read from Drive each time, so a new subfolder shows up without code changes,
// and "Delete":
//   Video Inbox: an approved video gets the Video subfolders; a video that failed validation, could
//                not be probed, or a file that is not a video gets "Delete" only; a video still
//                being probed gets no dropdown yet (validation.js sets these).
//   Image Inbox: images are not validated, so every image gets the Images subfolders.
// Choosing a subfolder moves the file there, a video with its probe results stored on it as Drive
// properties (mediaProbe.js); "Delete" moves it to the trash (restorable for 30
// days in a shared drive). The tab is then rebuilt, so its row goes; a moved video also refreshes
// Video Assets. When another file in the target folder has the same name:
//   the same content (Drive's MD5 checksum): a re-upload of a file already there, so that one is
//     kept as it is (its Drive ID, links and history) and the inbox copy goes to the trash;
//   other content: the moved file gets a -001, -002 … suffix before its extension
//     (clip.mov -> clip-001.mov), since Video Assets and the recipients cannot tell two files with the same
//     name apart.
//
// Moving and trashing need full Drive access, which a simple trigger does not have: handleEdit is
// an installable edit trigger (installEditTrigger), and acts with the access of whoever installed it.

const ACTION_HEADER = 'Action';
const DELETE_CHOICE = 'Delete';

/** The subfolder names of an inbox list's production folder, sorted. */
function productionFolderNames_(list) {
  const names = [];
  const it = DriveApp.getFolderById(requiredProperty_(list.production)).getFolders();
  while (it.hasNext()) names.push(it.next().getName());
  return names.sort(function (a, b) { return a.localeCompare(b); });
}

/**
 * Sets the Action dropdown of each row below the header: choices[i] is row i + 2's list ([] for
 * none), and clears it on the rows below, down to oldLastRow.
 */
function setActionDropdowns_(sheet, choices, oldLastRow) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const col = headers.indexOf(ACTION_HEADER) + 1;
  if (!col) return;
  if (choices.length) {
    sheet.getRange(2, col, choices.length, 1).setDataValidations(choices.map(function (list) {
      return [list.length ? SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build() : null];
    }));
  }
  if (oldLastRow > choices.length + 1) {
    sheet.getRange(choices.length + 2, col, oldLastRow - choices.length - 1, 1).clearDataValidations();
  }
}

/**
 * Another file in folder with file's name and the same content (Drive's MD5 checksum), or null.
 * Files without a checksum (Google Docs and the like) are never identical.
 */
function identicalFileIn_(folder, file) {
  const md5 = driveMd5_(file.getId());
  if (!md5) return null;
  const it = folder.getFilesByName(file.getName());
  while (it.hasNext()) {
    const other = it.next();
    if (other.getId() !== file.getId() && driveMd5_(other.getId()) === md5) return other;
  }
  return null;
}

/** A Drive file's MD5 checksum, from the Drive advanced service ('' when it has none). */
function driveMd5_(id) {
  return Drive.Files.get(id, { fields: 'md5Checksum', supportsAllDrives: true }).md5Checksum || '';
}

/**
 * The name file can have in folder: its own, unless another file there (another Drive ID) has it;
 * then the first of name-001.ext, name-002.ext … that no other file there has.
 */
function freeFileName_(folder, file) {
  const takenByOther = function (name) {
    const it = folder.getFilesByName(name);
    while (it.hasNext()) if (it.next().getId() !== file.getId()) return true;
    return false;
  };
  const name = file.getName();
  if (!takenByOther(name)) return name;
  const dot = name.lastIndexOf('.');
  const base = dot > 0 ? name.slice(0, dot) : name;
  const ext = dot > 0 ? name.slice(dot) : '';
  for (let n = 1; n < 1000; n++) {
    const candidate = `${base}-${String(n).padStart(3, '0')}${ext}`;
    if (!takenByOther(candidate)) return candidate;
  }
  throw new Error(`"${name}" finns redan 1000 gånger i ${folder.getName()}`);
}

/** Run once from the editor: installs handleEdit for this spreadsheet (replacing any earlier one). */
function installEditTrigger() {
  deleteTriggers_('handleEdit');
  ensureEditTrigger_('handleEdit', SpreadsheetApp.getActiveSpreadsheet());
  Logger.log(`handleEdit installed; it acts as ${Session.getEffectiveUser().getEmail()}`);
}

// INSTALLABLE EDIT TRIGGER: a choice in an Action column (the inbox tabs, Video Assets), or an
// edited profile (validation.js).
function handleEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() === PROFILES_SHEET) {
    validateVideoInbox_();
    return;
  }
  if (sheet.getName() === VIDEO_SHEET_NAME) {
    if (e.range.getRow() < 2 || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
    if (e.range.getColumn() !== ACTION_COL) return;
    const choice = String(e.value || '').trim();
    if (choice) runVideoAssetAction_(sheet, e.range.getRow(), choice, editorOf_(e));
    return;
  }
  const list = INBOX_LISTS.filter(function (l) { return l.sheet === sheet.getName(); })[0];
  if (!list || e.range.getRow() < 2 || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  if (e.range.getColumn() !== headers.indexOf(ACTION_HEADER) + 1) return;
  const choice = String(e.value || '').trim();
  if (!choice) return;
  // Moving takes a few seconds: say so straight away.
  const name = sheet.getRange(e.range.getRow(), headers.indexOf('Name') + 1).getDisplayValue();
  toast_(choice === DELETE_CHOICE ? `Deleting "${name}"…` : `Moving "${name}" to ${choice}…`, list.sheet, 5);
  runInboxAction_(list, sheet, e.range.getRow(), choice);
}

/** Who made an edit: the event's user (installable triggers in the domain), else the script's users. */
function editorOf_(e) {
  const fromEvent = e && e.user && e.user.getEmail ? e.user.getEmail() : '';
  return fromEvent || Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
}

/**
 * A choice in Video Assets' Action column: send to a recipient (recipients.js), or Delete (to the
 * trash, then Video Assets is listed again). The choice is cleared either way.
 */
function runVideoAssetAction_(sheet, row, choice, requestedBy) {
  const name = sheet.getRange(row, 2).getDisplayValue();
  const recipient = recipientByChoice_(choice);
  toast_(recipient ? `Sending "${name}" to ${recipient.label.toLowerCase()}…` : `Deleting "${name}"…`, VIDEO_SHEET_NAME, 5);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sheet.getRange(row, ACTION_COL).setValue('');
    let done;
    if (recipient) {
      done = sendToRecipient_(recipient, sheet, row, requestedBy);
    } else if (choice === DELETE_CHOICE) {
      DriveApp.getFileById(sheet.getRange(row, 1).getDisplayValue()).setTrashed(true);
      done = `"${name}" ligger nu i papperskorgen`;
      listDriveVideos();
    } else {
      done = `"${choice}" är inget val här`;
    }
    console.log(done);
    toast_(done, VIDEO_SHEET_NAME, 10);
  } catch (e) {
    console.log(`${VIDEO_SHEET_NAME} row ${row} (${name}): ${choice} failed: ${e.message}`);
    toast_(`${name}: ${e.message}`, VIDEO_SHEET_NAME, 10);
  } finally {
    lock.releaseLock();
  }
}

/** Moves (or trashes) the file on row of an inbox tab as choice says, then rebuilds the tab. */
function runInboxAction_(list, sheet, row, choice) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const nameCol = headers.indexOf('Name') + 1;
  const id = sheet.getRange(row, 1).getDisplayValue();
  const name = sheet.getRange(row, nameCol).getDisplayValue();
  try {
    const file = DriveApp.getFileById(id);
    let done;
    if (choice === DELETE_CHOICE) {
      file.setTrashed(true);
      done = `"${name}" ligger nu i papperskorgen`;
    } else {
      // Only an approved video may be moved, whatever was typed or pasted.
      if (list.probe && !/^Approved for/.test(sheet.getRange(row, nameCol).getNote())) {
        throw new Error(`"${name}" är inte godkänd för någon profil`);
      }
      const found = DriveApp.getFolderById(requiredProperty_(list.production)).getFoldersByName(choice);
      if (!found.hasNext()) throw new Error(`det finns ingen mapp "${choice}"`);
      const target = found.next();
      const twin = identicalFileIn_(target, file);
      if (twin) {
        // A re-upload: what is in production stays; the inbox copy goes.
        file.setTrashed(true);
        done = `"${name}" fanns redan i ${choice} med samma innehåll: inkorgens kopia ligger nu i papperskorgen`;
      } else {
        const newName = freeFileName_(target, file);
        // The probe results go with the video (mediaProbe.js), so Video Assets need not probe it.
        if (list.probe) {
          const props = inboxRowProperties_(sheet, row);
          if (Object.keys(props).length) setDriveProperties_(file.getId(), props);
        }
        if (newName !== file.getName()) file.setName(newName);
        file.moveTo(target);
        done = newName === name ? `"${name}" är flyttad till ${choice}` : `"${name}" är flyttad till ${choice} som "${newName}"`;
      }
    }
    console.log(done);
    listInboxes_([list.folder]);
    if (list.probe && choice !== DELETE_CHOICE) {
      try { listDriveVideos(); } catch (e) { console.log('Video Assets not refreshed: ' + e.message); }
    }
    toast_(done, list.sheet);
  } catch (e) {
    sheet.getRange(row, headers.indexOf(ACTION_HEADER) + 1).setValue('');
    console.log(`${list.sheet} row ${row} (${name}): ${choice} failed: ${e.message}`);
    toast_(`${name}: ${e.message}`, list.sheet, 10);
  } finally {
    lock.releaseLock();
  }
}
