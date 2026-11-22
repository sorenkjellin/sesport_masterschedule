# User tests: Asset Manager

Tests a person does in the seSPORT AssetManager spreadsheet, beyond the unit tests and
`runPostDeployTests`. Tick a test when it has passed; untick it when a change affects it. Tests
marked 👥 need a colleague, ⏸ were postponed until material or people are available.

Suggested order: 4, 5, 16, 17 first (they cover the most with the least material), then 25 when
DMC is staffed.

## New-asset mail and the inbox tabs

- [ ] **4. An image in Inbox/Images.** Upload one. *Expect within a minute:* a mail starting
      "This is an automated message from the seSPORT Asset Manager", with a link to Image Inbox,
      which lists the image.
- [ ] **5. A video in Inbox/Video.** Upload one. *Expect:* the mail links to Video Inbox ("probing
      1 video…"); the row goes Queued → Probed, and is validated within a minute (or at once with
      Inkorg → Validera videor).
- [ ] **6. Not a video.** Drop a .txt into Inbox/Video. *Expect:* Probe status "Not a video", and
      only Delete in Action.

## Probing formats not yet seen live

- [ ] ⏸ **7. Interlaced ProRes,** ideally in MXF. *Expect:* Scan "Interlaced TFF" or "Interlaced
      BFF", not Progressive.
- [ ] ⏸ **8. ProRes 4444 with alpha.** *Expect:* Codec "ProRes 4444", Alpha "Yes", Bit depth 12.
- [ ] ⏸ **9. H.265 .mp4.** *Expect:* Codec "H.265 Main" or "H.265 Main 10", Scan "Progressive".
- [ ] ⏸ **10. Interlaced H.264.** *Expect:* "Interlaced TFF/BFF", or "… (measured)" when the file
      does not say.
- [ ] **11. A large file** (several GB). *Expect:* probed within seconds (only parts are read).

## Validation

- [ ] **12. One clip per profile.** *Expect:* "Approved for:" with the right profiles on each.
- [ ] **13. A shared rule broken** (e.g. 25 fps). *Expect:* the row grey, Frame rate light red,
      note "Frame rate should be 50".
- [ ] **14. Too long, or the wrong bitrate, for its codec.** *Expect:* notes such as "Max duration
      is … when using …" or "Max Mbit is … when using …".
- [ ] **15. Edit a profile** (e.g. Max Mbit). *Expect:* the videos are validated again at once.

## The Action column

- [ ] **16. Move an approved video.** *Expect:* "Moving …" at once; the row goes; the clip shows
      in Video Assets with a black Path and its duration; its values are stored on the file.
- [ ] **17. Delete a failed clip.** *Expect:* the row goes; the file is in the shared drive's
      trash and can be restored.
- [ ] **18. Same name, other content.** Approve and move a different clip with the name of one
      already in production. *Expect:* it arrives as "name-001.mov".
- [ ] **19. Re-upload.** Move an identical copy of a clip already in production. *Expect:* "…
      fanns redan … med samma innehåll"; the inbox copy is in the trash, production untouched.
- [ ] **20. An image to Genre, Series or Episodes.** *Expect:* moved; in the master schedule,
      Bilder → Lista bilder lists it.
- [ ] **21. A new subfolder** under Video or Images. *Expect:* in the dropdowns after the next
      rebuild.

## Stored values and analysis

- [ ] **22. A clip put straight into a production folder,** then Video → Lista videor. *Expect:*
      its Path dark grey.
- [ ] **23. Video → Analysera videor utan metadata,** with Video Assets open. *Expect:* the grey
      rows turn black one by one; the next Lista videor shows their durations (MXF too).
- [ ] **24. Summing durations.** `=SUM(…)` over Duration, formatted `[h]:mm:ss`. *Expect:* the
      right total, also past 24 hours.

## Playout (DMC)

- [ ] ⏸ **25. A test send from the Asset Manager** (when DMC is staffed). In Video Assets, choose
      Send to Playout in a small clip's Action column. *Expect:* "Sending … to playout…" at once,
      then Recipient status Playout: Köad → Playout: Skickar … → Playout: Skickad. If it stays at Köad, note the time and check the
      job's logs.
- [ ] **25b. Send a clip that is already on its way.** *Expect:* nothing is sent, and a message
      says it is on its way.
- [ ] **25c. Delete from Video Assets.** Choose Delete for a clip you can spare. *Expect:* the row
      goes, and the file is in the shared drive's trash (restorable for 30 days).
- [ ] ⏸ **26. The old test file.** Ask DMC to delete "TEST ONLY - FAKE COMMERCIAL - TEST.mp4".

## Other users 👥

- [ ] **27. A colleague picks a folder in Action.** *Expect:* the move works (it runs with the
      trigger owner's access). Note whether they see the "Moving …" message.
- [ ] **28. A colleague chooses Inkorg → Uppdatera inkorgarna** with an unprobed video in the
      inbox. *Expect:* it is probed. If they lack permission on the probe job, the message says
      "analysen startade inte": give them the roles below.
- [ ] **29. A colleague chooses Send to Playout.** *Expect:* it works without any role of their
      own (it runs as the edit trigger's owner), and the job's log names the colleague.

For 28, everyone who refreshes the inbox needs `roles/run.jobsExecutorWithOverrides` and
`roles/run.viewer` on the media-probe job (docs/dmc-transfer.md, step 7), best given through a
Google group.
