# Validating videos against profiles

Which quality profiles each probed video in **Video Inbox** meets (see docs/media-probe.md for the
probing). The profiles are in the **Video Profiles** tab of the Asset Manager; the code is
`apps/video-assets/src/validation.js`.

## The Video Profiles tab

One row per profile, with its name in column A. Add, change or remove profiles at any time; the
videos are validated again as soon as a profile is edited.

The other columns are rules, under headers that name Video Inbox columns:

- **`Min <column>` / `Max <column>`**: a lower or upper limit, e.g. *Max duration* (a duration,
  typed as `0:01:30`), *Min MBit*, *Max Mbit*.
- **Any other header** must be a Video Inbox column (*Frame rate*, *Width*, *Codec*, *Scan*,
  *Bit depth*, *Alpha*, *Audio* …), and the video must have one of the values in the cell.
  Several values are separated by `/`: `8/10`, `none / stereo (2 ch)`.
- **`*`** stands for any text, in any column but Min and Max: `ProRes 422*` allows ProRes 422, ProRes 422 HQ, ProRes 422 LT
  and ProRes 422 Proxy (but not ProRes 4444); `H.26*` allows H.264 and H.265.
- **An empty cell** allows anything.

Upper and lower case do not matter, and decimals may be written with a comma (`29,97`). A column
whose header names no Video Inbox column is ignored; the post-deploy check **Video Profiles** names
it.

## What the videos get

For every video in Video Inbox with *Probed* in **Probe status**:

- **Meets one or more profiles:** a note on its name: *Approved for:* and the profiles, one per line.
- **Meets none:** the row is greyed out, and the cells that make it fail are light red, with a note:
  - a rule every profile has in common: *Frame rate should be 50*;
  - a codec no profile allows: *VC3/DNxHD is not an allowed codec.*;
  - otherwise, compared with the profiles that allow its codec: *Max duration is 2:00:00 when using
    H.264 Main* (the widest limit of those profiles, or all the values they allow).
  - When every property is allowed by some profile with the codec, but no single profile allows them
    all (say, short enough only for one and the right bitrate only for another), the cells the
    nearest profile (the one with the fewest failures) does not allow are marked, naming it:
    *Max duration is 0:00:30 when using ProRes 422 (Graphics)*.

Videos not probed yet, or that could not be probed, are left unmarked.

## When it runs

- When Video Inbox is rebuilt (the asset watcher, or Inkorg → Uppdatera inkorgarna).
- When a profile is edited (by the installed edit trigger, handleEdit: see docs/inbox-actions.md).
- On every run of the asset watcher, every 10 minutes: probe results are written by the probe job
  through the Sheets API, which does not count as an edit, so a new result is validated within
  10 minutes.
- **Inkorg → Validera videor**, any time.

The marks (colours and notes) of Video Inbox's rows are set by validation; colours or notes added
there by hand are replaced at the next run.
