# Probing the videos in Video Inbox

What each video in Inbox/Video is (duration, frame rate, size, codec, bitrate, scan, bit depth, alpha and audio
layout), shown in the Asset Manager's **Video Inbox** tab, and how to set it up.

## How it works

1. When **Video Inbox** is rebuilt, by the asset watcher when new files land in Inbox/Video or by
   **Inkorg → Uppdatera inkorgarna**, the Asset Manager (`apps/video-assets/src/mediaProbe.js`)
   finds the videos whose **Probe status** is empty, starts the Cloud Run job `media-probe` with
   their Drive file IDs and marks them *Queued*.
2. The job (`services/media-probe`) runs **ffprobe** on each video straight from Drive. ffprobe only
   fetches the parts of the file it needs, usually a few MB, so a video takes seconds.
3. The job writes the columns, and *Probed <time>* in **Probe status**, or *Error: <reason>*.

The columns stay with their video when the tab is rebuilt, and a probed video is not probed again.
To probe a video again, clear its **Probe status** and choose Inkorg → Uppdatera inkorgarna.

| Column | Example | Notes |
|---|---|---|
| Duration | 0:01:05:01 | H:MM:SS:FF on the clip's own timebase (30 for 29.97, non-drop). Until the clip is probed, Drive's duration at 25 fps, if Drive has one (it has none for MXF) |
| Frame rate | 25, 29.97, 50 | frames per second (25 for 1080i50) |
| Width, Height | 1920, 1080 | pixels |
| Codec | ProRes 422 HQ, ProRes 4444, H.264 High, H.265 Main 10 | |
| Bitrate | 73.1 Mbit/s, ≈ 73.5 Mbit/s | the video's own; for MXF (which does not give it) the whole file minus the audio, marked ≈ |
| Scan | Progressive, Interlaced TFF, Interlaced BFF | see below |
| Bit depth | 8, 10, 12 | ProRes 4444 is always decoded as 12-bit |
| Alpha | Yes / No | |
| Audio | 8 × mono (8 ch), stereo (2 ch), none | tracks grouped, then the channel total |

### How the scan is found

- **ProRes**: from the first frame. Every ProRes frame says how it is coded; the container may not.
  (FFmpeg reports interlaced ProRes in MXF as progressive at the stream level.)
- **H.265**: always *Progressive*. Interlaced H.265 is rare and often wrongly flagged.
- **H.264 and anything else**: from the stream's field order. When the file does not say, or says
  it ambiguously, the job decodes the first 300 frames and measures (ffmpeg's `idet` filter); the
  cell then ends in *(measured)*. Measuring looks at the picture, so progressive material in an
  interlaced stream is measured as progressive.

Files that are not videos (Drive type `video/…`, or `application/mxf`) are listed but not probed.

## One-time setup

In the same Cloud project as the DMC transfer job (see docs/dmc-transfer.md, step 1 for the APIs).

1. Create the job's service account:

   ```sh
   gcloud iam service-accounts create media-probe --display-name "Media probe" --project PROJECT
   ```

2. Share with `media-probe@PROJECT.iam.gserviceaccount.com`:
   - **Inbox/Video** as **Viewer**;
   - **the Asset Manager spreadsheet** as **Editor**, so the job can write its columns.

3. Deploy the job:

   ```sh
   cp services/media-probe/job.env.example services/media-probe/job.env
   # fill in GCP_PROJECT and SPREADSHEET_ID (job.env is not in git)
   scripts/deploy-media-probe.sh
   ```

4. Deploy the Asset Manager (`scripts/push.sh video-assets`) and run `runPostDeployTests`. The
   checks **Media probe: Cloud Run job** and **Media probe: permissions** should pass. Then choose
   Inkorg → Uppdatera inkorgarna once: every video already in Inbox/Video is probed.

The job is started as the person whose action rebuilt the tab: the watcher's account
(WATCH_RUN_AS) for new files, or whoever chose the menu item. They need permission to start it,
like the DMC job (roles/run.jobsExecutorWithOverrides; docs/dmc-transfer.md, step 7). When it
cannot be started, the statuses stay empty, so the next rebuild tries again, and the mail or the
menu's message says why.

## When something goes wrong

- **Probe status stays at "Queued".** The job stopped before it could write to the sheet: look in
  Cloud Run → Jobs → media-probe → Executions → Logs. Usually the spreadsheet is not shared with
  the job's account. Clear the status and choose Inkorg → Uppdatera inkorgarna to try again.
- **"Error: ffprobe failed: … 403 …".** Inbox/Video is not shared with the job's account.
- **"Error: no video stream …".** ffprobe could not find a video in the file.

## Cost

A few seconds of Cloud Run time per video, and the few MB ffprobe reads. Measuring an H.264 file
reads and decodes 300 frames, which takes longer, but is only done when the file does not say how
it is scanned.
