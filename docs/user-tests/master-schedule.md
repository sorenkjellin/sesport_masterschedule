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

## Program: Längd from the körschema

Run `installProgramLengthTrigger` once first (as the account that can open all linked files). These
work in the Program tab, columns F (Längd) and G.

- [x] **Format.** *Expect within a minute:* lengths in F show as 1.30 (hours without a leading zero).
- [ ] **No link.** Type a length in F on a row with nothing in G. *Expect:* it stays.
- [ ] **Link a körschema in G.** *Expect within a minute:* F shows the file's F2 minus E2, in grey,
      with a note "Körschema: längden hämtas från filen i kolumn G."
- [ ] **Change the end time** in the linked file. *Expect within a minute or two:* F follows.
- [ ] **Type over F on a linked row.** *Expect:* it is put back, and a message says the length is
      governed by the körschema.
- [ ] **Remove the link.** *Expect:* the last length stays, the grey and the note go, and F can be
      typed in.
- [ ] **Empty F2** in a linked file you can spare. *Expect:* F keeps its value, and its note says
      "F2 är tom". Put F2 back: the note returns to normal.
- [ ] **Past midnight.** A körschema with E2 23:00 and F2 01:30. *Expect:* 2.30.
- [ ] **Sort or move Program rows.** *Expect:* each length stays with its körschema.
- [ ] 👥 **A colleague types over F on a linked row.** *Expect:* put back. Note whether they see the
      message.

## Housekeeping

- [ ] **Old körschema triggers.** Within a minute or so of the deploy, the Triggers page (as the owner of
      `checkProgramLengths`) lists `syncKorSchemaValueToKalender_` only for körscheman still linked in
      Kalender column L.
- [ ] **Data → Arkivera gamla rader.** Archive rows that have a körschema. *Expect:* its trigger is
      gone at once, or within an hour if someone else archived; a körschema still in Kalender
      still updates Slutar when its end time changes.

- [ ] 👥 **Leftover watcher triggers.** Colleagues report no error emails about
      `checkForNewAssets` from the master schedule for a few weeks. *Then:* the stub in
      `apps/master-schedule/src/assetWatcher.js` can be removed.
