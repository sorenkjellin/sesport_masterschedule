# User tests: master schedule

Tests a person does in the master schedule spreadsheet, beyond the unit tests and
`runPostDeployTests`. Tick a test when it has passed; untick it when a change affects it. Tests
marked 👥 need a colleague.

## Access

- [x] 👥 **Outside user.** Someone outside sesport.tv opens the master schedule and uses a menu
      item. *Expect:* "Google hasn't verified this app" once, then the usual permission screen,
      then everything works.

## Editing (the handleEdit trigger)

- [x] 👥 **After the Cloud project switch (2026-11-21).** Change Genre on a Kalender row.
      *Expect:* the league dropdown follows. If nothing happens, the trigger's owner opens the
      spreadsheet and approves the script again.
- [x] **Live.** Set Live/Rerun to Live, then tick the checkbox. *Expect:* a körschema is created
      and linked in column S.
- [x] **Playout.** Change Content Type in Playout. *Expect:* the dropdown beside it follows.

## Housekeeping

- [ ] 👥 **Leftover watcher triggers.** Colleagues report no error emails about
      `checkForNewAssets` from the master schedule for a few weeks. *Then:* the stub in
      `apps/master-schedule/src/assetWatcher.js` can be removed.
