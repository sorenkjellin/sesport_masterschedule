# Moving files out of the inbox

The **Action** column (D) of **Video Inbox** and **Image Inbox** moves a file to a production
folder, or deletes it. The code is `apps/video-assets/src/actions.js`.

## What the dropdown offers

The production folders are **Video** (`VIDEO_FOLDER_ID`) and **Images** (`IMAGES_FOLDER_ID`). Their
subfolders are read from Drive every time the dropdowns are set, so a new subfolder appears without
any change here.

| | Dropdown |
|---|---|
| A video approved for one or more profiles | the Video subfolders |
| A video that failed validation, could not be probed (*Error*), or a file that is not a video (*Not a video*) | Delete |
| A video still being probed (*Queued*, *Probing*) | none yet |
| An image (images are not validated) | the Images subfolders |

## What a choice does

- **A subfolder:** the file is moved there, a video with its probe results and approved profiles
  stored on it (docs/media-probe.md, "Results that go with the file"). The inbox tab is rebuilt,
  so its row goes; a moved video also refreshes Video Assets. A video that is not approved is not
  moved, even if its folder was typed or pasted in. When another file in that folder already has the same name:
  - **the same content** (Drive's MD5 checksum): it is a re-upload of a file already in production.
    That file stays as it is, keeping its Drive ID, links and history, and the inbox copy goes to
    the trash;
  - **other content**: the moved file gets a suffix before its extension, `clip.mov` becoming
    `clip-001.mov` (then `-002` …), since Video Assets and playout cannot tell two files with the same
    name apart.
- **Delete:** the file goes to the trash. In a shared drive, it can be restored from the trash for
  30 days.

When a choice cannot be carried out, it is undone and a message says why.

## Setup

1. Script Property `IMAGES_FOLDER_ID`: the ID of the Images folder (`VIDEO_FOLDER_ID` is already set).
2. Run `installEditTrigger` once in the Apps Script editor. Moving and trashing need full Drive
   access, which a simple edit trigger does not have; the installed trigger, `handleEdit`, acts with
   the access of the person who installed it, for everyone's choices. That person needs Content
   manager or Manager on the shared drive. The trigger also re-validates when a profile is edited.
3. `runPostDeployTests`: **Drive: production folders** lists the subfolders, and **Triggers**
   warns when handleEdit is not installed.
