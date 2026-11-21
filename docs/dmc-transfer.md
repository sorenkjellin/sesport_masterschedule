# Sending clips to DMC

How video clips get from the Drive "Video" folder to DMC's FTP server
(`ftp-no1.dmcproduction.se`, port 21, FTPS with explicit TLS), and how to set it up.

## How it works

1. In the **Video Assets** spreadsheet (its own file, not a tab in the master schedule), someone
   ticks **Skicka till DMC** (column G) for the clips to send and chooses
   **Video → Skicka markerade till DMC**.
2. The app (`apps/video-assets/src/dmcTransfer.js`) checks the ticked clips, sets them to
   *Köad* and starts the Cloud Run job `dmc-transfer` with their Drive file IDs. It uses the
   person's own Google login to do this.
3. The job (`services/dmc-transfer`) streams each clip from Drive to DMC. It uploads to
   `<name>.part`, checks that the size at DMC matches the size in Drive, then renames the file to
   its Drive name, overwriting any existing file with that name. Clips are not put in subfolders.
4. The job writes progress to **DMC-status** (column H): *Skickar … (25 %)* → *Skickad <time>*,
   or *Fel: <reason>*.

Apps Script can't do the transfer itself: it has no FTP support, and a script run is stopped after
6 minutes. A Cloud Run job can run for up to 24 hours per clip.

The FTP password is stored only in Secret Manager. It is not in the code, the repo or the sheet.

## One-time setup

Do these steps as someone who administers the Google Cloud project. The commands use `gcloud`;
install it with `brew install --cask google-cloud-sdk`, or run them in Cloud Shell in the console.
Replace `PROJECT` with the project ID.

### 1. Turn on the APIs

```sh
gcloud config set project PROJECT
gcloud services enable run.googleapis.com drive.googleapis.com sheets.googleapis.com \
  secretmanager.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com \
  iam.googleapis.com
```

Or in the console: APIs & Services → Library, then search for each name.

| API | Service name | Used for |
|---|---|---|
| Cloud Run Admin API | `run.googleapis.com` | Runs the job. The sheet uses it to start the job, and the post-deploy check uses it to read the job's settings. |
| Google Drive API | `drive.googleapis.com` | The job reads the clips from Drive. The app's Drive advanced service needs it too (step 6). |
| Google Sheets API | `sheets.googleapis.com` | The job writes DMC-status to the sheet. |
| Secret Manager API | `secretmanager.googleapis.com` | Stores the FTP password, which the job reads when it starts. |
| Cloud Build API | `cloudbuild.googleapis.com` | Builds the job's container when you deploy. |
| Artifact Registry API | `artifactregistry.googleapis.com` | Stores the built container. |
| Identity and Access Management (IAM) API | `iam.googleapis.com` | Creates the service account (step 2). It is often on already. |

Logging is on in every project already, so you don't need to turn anything on for the job's logs.

### 2. Create the job's service account

```sh
gcloud iam service-accounts create dmc-transfer --display-name "DMC transfer"
```

The job runs as `dmc-transfer@PROJECT.iam.gserviceaccount.com`.

### 3. Store the FTP password

Use the console (Security → Secret Manager → Create secret):

- Name: `dmc-ftp-password`
- Secret value: the password, typed in by you.

Then give the service account access to it:

```sh
gcloud secrets add-iam-policy-binding dmc-ftp-password \
  --member serviceAccount:dmc-transfer@PROJECT.iam.gserviceaccount.com \
  --role roles/secretmanager.secretAccessor
```

If the password changes, add a new version of the secret. The next transfer uses the new version
without redeploying the job.

### 4. Share Drive and the sheet with the service account

- **The "Video" folder:** share it with `dmc-transfer@PROJECT.iam.gserviceaccount.com` as
  **Viewer**. If it is in a shared drive, add the account to that shared drive as a Viewer instead.
- **The Video Assets spreadsheet:** share it with the same address as **Editor**, so the job
  can write DMC-status.

If Drive refuses to share with the address, your Workspace domain blocks sharing outside the
domain. A Workspace admin has to allow sharing with this account (Admin console → Apps → Google
Workspace → Drive and Docs → Sharing settings).

### 5. Deploy the job

```sh
cp services/dmc-transfer/job.env.example services/dmc-transfer/job.env
# fill in GCP_PROJECT, FTP_USER, FTP_DIR and SPREADSHEET_ID (job.env is not in git)
gcloud auth login
scripts/deploy-dmc-transfer.sh
```

The script runs the job's tests, builds the job in Cloud Build and deploys it. The first deploy
asks to create an Artifact Registry repository; answer yes.

Ask DMC whether uploading as `<name>.part` and renaming is fine. If their system can't handle it,
set `FTP_TEMP_SUFFIX=none` in job.env and deploy again.

### 6. Connect the Video Assets script to the Cloud project

The app starts the job by calling the Cloud Run API. That call goes through the Apps Script
project's Cloud project, and the default one, which Google creates and hides, can't have the
Cloud Run API turned on. So the Video Assets script has to use your Cloud project instead:

1. In the Cloud console, go to **Google Auth Platform → Branding** (the OAuth consent screen) and
   set it up as **Internal**.
2. Copy the **project number** from the console's dashboard. This is the number, not the ID.
3. In the Video Assets spreadsheet's Apps Script editor, go to **Project Settings → Google Cloud Platform (GCP) Project →
   Change project** and paste the number.

**Turn on the Drive API (step 1) before you switch.** After the switch, "Lista videor" reads the
videos' durations through your Cloud project.

With Internal sign-in, only accounts in the project's organisation (sesport.tv) can use **any**
menu item of a script linked to it. That is why this lives in its own spreadsheet: never link the
master schedule's script to this project. (A script can't go back to Google's default project once
it has been switched; link it to another standard project instead.)

### 7. Decide who may send clips

Create a Google group, for example `dmc-sending@yourdomain`, add the people who send clips, then
give the group permission on the job:

```sh
for role in roles/run.jobsExecutorWithOverrides roles/run.viewer; do
  gcloud run jobs add-iam-policy-binding dmc-transfer --region europe-north1 \
    --member group:dmc-sending@yourdomain --role $role
done
```

`roles/run.viewer` lets them run the post-deploy check, which reads the job's settings.

### 8. Set up the app

1. In the Video Assets script's Script Properties, set `VIDEO_FOLDER_ID` (the Drive "Video"
   folder) and `DMC_GCP_PROJECT` (the project ID). Also set `DMC_REGION` and `DMC_JOB` if you
   changed them from `europe-north1` and `dmc-transfer`.
2. Deploy the app with `scripts/push.sh video-assets`, then run `runPostDeployTests`. The DMC
   checks should pass.
3. In the sheet, choose **Video → Lista videor i Google Drive** once. This creates the
   "Video Assets" tab with columns A–H.
4. Send a small test clip.

## Permission checks

Before anything is queued, **Skicka markerade till DMC** checks:
- that the person sending may start the job;
- that the Video Assets spreadsheet is shared with the job's service account as Editor;
- that every ticked clip is shared with it, at least as Viewer.

The service account's address is read from the job itself. If something is missing, an alert says
what to share, with which address, and nothing is sent. Sharing that the sender isn't allowed to
see is skipped, not counted as missing.

The post-deploy check **DMC transfer: permissions** checks the same for the Video folder, and also
whether the service account may read the password secret. Only Secret Manager admins can see that
last part; for anyone else it shows as a warning.

Only sharing with the address itself, or with anyone who has the link, is recognised. Sharing
through a Google group that contains the service account is reported as missing, so share
directly with the address.

## When something goes wrong

- **The status stays at "Köad".** The job stopped before it could write to the sheet. The check
  before sending should catch a missing share, so look in the job's logs first: Cloud Run → Jobs
  → dmc-transfer → Executions → Logs. To try again, clear the status cell, tick the clip and send it
  again.
- **"Fel: 530 Login incorrect."** The username in job.env or the password secret is wrong.
- **"Fel: … certificate …".** DMC's TLS certificate isn't trusted. Ask DMC about their
  certificate.
- **"Fel: överföringen startade inte: … 403 …".** The person who sent the clip isn't allowed to
  start the job (step 7), or the Apps Script project isn't using your Cloud project (step 6).
- **"Fel: DMC has … bytes".** The upload was incomplete. The partial file has been removed; send
  the clip again.

A failed clip is retried once automatically (`--max-retries 1`).

## Cost

The main cost is Google's outbound internet traffic, roughly USD 0.10–0.12 per GB, which comes to
about USD 0.50 for a 5 GB clip. Running the job costs a few cents per hour of transfer time. Check
current prices for the region.
