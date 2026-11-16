# seSport Apps Script

The Google Apps Script apps for seSport and the code they share.

## Layout

```
lib/                     shared code, copied into every app when it is deployed
  lib.js                 generic helpers: properties, HTTP, Cachearoo, Elmo file API, mail,
                         Drive and a folder watcher, sheets, dates, UI, triggers
apps/
  master-schedule/       the script bound to the master schedule spreadsheet
    .clasp.json          links the folder to its Apps Script project (rootDir: build)
    src/                 the app's own code and its appsscript.json manifest
    build/               src/ + lib/, assembled by scripts/push.sh (not in git)
scripts/
  push.sh                deploy an app
  pull.sh                fetch browser-editor changes back into an app's src/
```

All files in an Apps Script project share one global scope, so there are no imports: once
deployed, an app calls `sendMail_(...)`, `cachearooWrite_(...)` and so on directly.

`lib/` must stay generic: nothing about a particular app's sheets, columns or settings, everything
passed in as arguments. Functions there end in `_`, which keeps them out of the editor's Run menu.
lib.js defines only functions, so the order Apps Script loads files in does not matter.

## Deploying

One-time setup:

```sh
npm install -g @google/clasp
clasp login
```

Also turn on "Google Apps Script API" at <https://script.google.com/home/usersettings>.

```sh
scripts/push.sh master-schedule           # build and push to Apps Script
scripts/push.sh master-schedule --build   # build only, to look at apps/master-schedule/build
```

A push replaces every file in the Apps Script project with the build, so code edited in the
browser editor is lost unless it is pulled first:

```sh
scripts/pull.sh master-schedule   # copies the editor's files into src/, then check git diff
```

pull.sh does not copy the lib files back: if the editor's copy of lib.js differs from `lib/`, it
says so and the change has to be moved into `lib/` by hand.

A change in `lib/` reaches an app the next time that app is pushed. Push every app after changing
lib.js, or at least the ones that use what changed.

Secrets (API keys, Cachearoo URL, mail addresses) live in each project's Script Properties, not in
the code.

## Adding an app

1. Create the Apps Script project (for a sheet: Extensions → Apps Script in that sheet) and copy
   its Script ID from Project Settings.
2. Create `apps/<name>/.clasp.json` with that `scriptId` and `"rootDir": "build"`.
3. Fetch its manifest and any code: `scripts/pull.sh <name>` (create `apps/<name>/src` first).
4. Write the app in `apps/<name>/src`, using what lib.js offers, and deploy with
   `scripts/push.sh <name>`.
5. Add the OAuth scopes lib functions need to the app's `appsscript.json` if it lists scopes
   explicitly (see apps/master-schedule/src/appsscript.json), and enable the Sheets advanced service
   if it uses the smart-chip helpers.
