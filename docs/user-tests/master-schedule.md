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
      with a note "Längden styrs av grundkörschemat"
- [ ] **Change the end time** in the linked file. *Expect within a minute or two:* F follows.
- [x] **Type over F on a linked row.** *Expect:* it is put back, and a message says the length is
      governed by the körschema.
- [ ] **Remove the link.** *Expect:* the last length stays, the grey and the note go, and F can be
      typed in.
- [ ] **Empty F2** in a linked file you can spare. *Expect:* F keeps its value, and its note says
      "F2 är tom". Put F2 back: the note returns to normal.
- [ ] **Past midnight.** A körschema with E2 23:00 and F2 01:30. *Expect:* 2.30.
- [ ] **Sort or move Program rows.** *Expect:* each length stays with its körschema.
- [ ] 👥 **A colleague types over F on a linked row.** *Expect:* put back. Note whether they see the
      message.

## Kalender: Slutar from the körschema

- [ ] **No körschema.** Type an end time in C on a row with nothing in L. *Expect:* it stays.
- [x] **Create a körschema** with the checkbox. *Expect:* C shows its end time (F2), in grey, with the
      note "Sluttid styrs av körschemats längd".
- [x] **Type over C on that row.** *Expect:* it is put back, and a message says so.
- [ ] **Change the length in the körschema.** *Expect:* C follows.
- [x] **Change Börjar** on that row. *Expect:* C moves with it and stays there (check again after two
      minutes).
- [x] **Clear L.** *Expect:* the end time stays, the grey and the note go, and C can be typed in.
- [ ] **Two körscheman on one date** (two tabs in one file). *Expect:* each row has its own end time.

## Kalender: Slutar copied from Program's Längd

On a row without a körschema:

- [ ] **Fill in Börjar, Genre and Ligans namn.** *Expect:* Slutar becomes Börjar + the Längd of that
      genre and league in Program, not grey and without a note.
- [ ] **Type over Slutar.** *Expect:* it stays.
- [ ] **Change Börjar, then Genre or Ligans namn.** *Expect:* Slutar is copied again each time.
- [ ] **A league whose Längd is empty in Program.** *Expect:* Slutar is left alone.
- [ ] **Change Genre so that the league no longer fits.**
      *Expect:* Slutar is emptied at once; picking a matching league fills it in again.
- [ ] **Past midnight.** Börjar 23.00. *Expect:* the right time of day.

## Data → Kontrollera texter

Set the key first (docs/text-check.md). Use rows dated today or later.

- [x] **Run it from another tab.** *Expect:* the Kalender tab is shown at once, and the window opens.
- [x] **A misspelt text.** Put one in M and run the check. *Expect:* a window with the change marked;
      the cell is selected and light yellow.
- [x] **Uppdatera.** *Expect:* the cell gets the suggestion and its own colour back; the next cell is
      shown.
- [x] **Edit the suggestion, then Uppdatera.** *Expect:* your version is written.
- [x] **Behåll.** *Expect:* the text stays, and is not suggested when the check is run again.
- [x] **Change the text while its suggestion is shown, then Uppdatera.** *Expect:* it is left alone, and
      the window says so.
- [x] **Close with X.** *Expect:* the yellow goes within about four minutes, and a cell with a colour of
      its own gets that back.
- [x] **Names.** Player, team and league names. *Expect:* not "corrected".
- [x] **Old rows and formulas in M.** *Expect:* never suggested.
- [x] **The summary** after the last suggestion. *Expect:* "Kontrollerade 10 texter och uppdaterade 1."
- [x] **No errors.** Run it again when all is handled. *Expect:* "Inga fel hittades".
- [x] 👥 **A colleague runs it, also while your window is open.** *Expect:* "pågår redan", with Ta över.

## Bilder (faster listing and upload)

- [x] **Lista bilder.** *Expect:* the same rows as before, also from subfolders, and none wrongly
      "Missing in Drive".
- [x] **Ladda upp, with nothing new.** *Expect:* done in seconds; "Uploaded 0"; the execution log says
      "… in the list (… missing in Drive); 0 downloaded" and how long each step took.
- [ ] **Ladda upp, with a new image and a new version of one** (in Drive: Manage versions → Upload
      new version). *Expect:* just those two are "Uploaded …", and the new pictures show in Elmo.
- [ ] 👥 **Two at once.** Start Lista or Ladda upp while another upload runs. *Expect:* a message that
      images are already being listed or uploaded.

## Housekeeping

- [x] **Old körschema triggers.** Within a minute or so of the deploy, the Triggers page (as the owner of
      `checkProgramLengths`) lists `syncKorSchemaValueToKalender_` only for körscheman still linked in
      Kalender column L and not in the trash (within an hour).
- [ ] **A trashed körschema.** Put one that Kalender links to in the trash, and later restore it.
      *Expect:* its trigger goes within an hour, and comes back within an hour of restoring.
- [ ] **Data → Arkivera gamla rader.** Archive rows that have a körschema. *Expect:* its trigger is
      gone at once, or within an hour if someone else archived; a körschema still in Kalender
      still updates Slutar when its end time changes.

- [ ] 👥 **Leftover watcher triggers.** Colleagues report no error emails about
      `checkForNewAssets` from the master schedule for a few weeks. *Then:* the stub in
      `apps/master-schedule/src/assetWatcher.js` can be removed.
