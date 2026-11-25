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
      postDeployTests.js checks run in the Apps Script editor after a deploy
    tests/app.test.js    unit tests for the app (not deployed)
    build/               src/ + lib/, assembled by scripts/push.sh (not in git)
  video-assets/          the script bound to the "Video Assets" spreadsheet: lists the Drive
                         videos, sends clips to playout (DMC), mails about new files in the asset folder
                         (same layout as master-schedule)
tests/
  harness.js             loads Apps Script files into Node with fake Google services
  lib.test.js            unit tests for lib.js
services/
  dmc-transfer/          Cloud Run job sending video clips from Drive to DMC's FTP server
                         (started from Video Assets; see docs/dmc-transfer.md)
  media-probe/           Cloud Run job probing the videos in Inbox/Video with ffprobe
                         (started from the Asset Manager; see docs/media-probe.md)
docs/                    setup guides: dmc-transfer, media-probe, video-validation, inbox-actions,
                         text-check
  user-tests/            checklists of what to test by hand, one per spreadsheet
scripts/
  push.sh                test and deploy an app
  pull.sh                fetch browser-editor changes back into an app's src/
  deploy-dmc-transfer.sh test and deploy the DMC transfer job to Cloud Run
  deploy-media-probe.sh  test and deploy the media probe job to Cloud Run
.github/workflows/       runs the unit tests on every push to GitHub
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
scripts/push.sh video-assets              # the same for the Video Assets spreadsheet
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

### Why there are two spreadsheets

The Video Assets script uses the Google Cloud project that holds the DMC transfer job (it has to,
to start the job), and that project only lets sesport.tv accounts sign in. The script's sign-in
applies to every menu item in its spreadsheet, so anything in it is closed to people outside
sesport.tv. The master schedule is used by people outside sesport.tv too, so it stays on a Cloud
project of its own, with no Cloud scopes.

## Testing

Three layers:

1. **Unit tests** (`npm test`, Node 22 or later). They load lib.js and the app's files into Node
   with fake Google services (`tests/harness.js`) and test the logic: the coming-up feed, image
   choice, Elmo and Cachearoo requests, the folder watcher, sending clips to playout, parsing and
   formatting. The DMC transfer job's tests send real FTPS uploads to a fake server; they need its
   one dependency, installed once with `npm ci --prefix services/dmc-transfer`. The media probe
   job's tests include real ffprobe runs, which are skipped where ffmpeg is not installed.
   `scripts/push.sh` runs the lib and app tests first and pushes nothing if one fails; GitHub runs
   all of them on every push.
2. **Post-deploy tests**, in the live project. After a push, open `postDeployTests` in the Apps Script
   editor (`clasp open-script`) and press Run: `runPostDeployTests` is preselected as the file's first
   function. It checks Script Properties, the sheets and the Kalender columns the code relies on,
   access to Drive, Elmo and the Sheets service, a dry run of the coming-up feed, and the triggers.
   It only reads. Results are in the execution log; the run is marked failed when a check fails.
   Warnings (e.g. a trigger installed by someone else, which the API does not show you) do not fail it.
3. **By hand**: what needs a person in the sheet. The checklists are in `docs/user-tests/`
   (`master-schedule.md`, `asset-manager.md`): tick a test when it has passed, untick it when a
   change affects it. After a deploy that touches them, at least:
   - the menus appear after reloading the spreadsheet;
   - changing Genre in Kalender updates the league and team dropdowns;
   - setting Live/Rerun to Live adds the checkbox; ticking it creates a körschema;
   - changing Content Type in Playout updates the dropdown beside it.

New code in `lib/` gets tests in `tests/lib.test.js`; new app logic in `apps/<app>/tests/`. Logic
that is tangled with reading and writing sheets is easiest to test once it is moved into a function
that takes values and returns values.

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
