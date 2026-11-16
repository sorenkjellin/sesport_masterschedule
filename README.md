# seSport master schedule – Apps Script

The Apps Script project bound to the seSport master schedule spreadsheet: the Elmo image and video
lists, the "coming up" feed to Cachearoo, körschema handling in Kalender, the EPG export, archiving
and the Drive asset watcher.

## Layout

- `src/` is what is deployed to Apps Script (clasp's `rootDir`).
  - `lib.js`: generic helpers (properties, HTTP, Cachearoo, Elmo file API, mail, Drive, sheets, UI,
    triggers). It knows nothing about the schedule and should rarely change.
  - `rundownImages.js`: the seSport-specific code.
  - `appsscript.json`: the manifest (time zone, advanced services such as Sheets, scopes).
- `.clasp.json` links this folder to the Apps Script project.

All files in an Apps Script project share one global scope, so there are no imports. `lib.js` defines
only functions, so the order the files load in does not matter.

## Deploying with clasp

One-time setup:

```sh
npm install -g @google/clasp
clasp login
```

Also turn on "Google Apps Script API" at <https://script.google.com/home/usersettings>.

Then:

```sh
clasp pull     # get what is currently in the Apps Script editor (overwrites local files)
clasp push     # replace the Apps Script project's files with src/
clasp open-script  # open the project in the editor
```

`clasp push` replaces every file in the project: a file that exists only in the editor is deleted.
If anyone edits in the browser, `clasp pull` and commit first.

Secrets (API keys, Cachearoo URL, mail addresses) live in Script Properties, not in the code.
