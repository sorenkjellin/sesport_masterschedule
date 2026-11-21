// Unit tests for the Video Assets app. Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, scriptsIn, plain, FakeSheet, fakeSpreadsheet, fakeFolder, fakeFile, fakeDriveFiles } = require('../../../tests/harness');

// lib.js plus every app file, in one scope, as Apps Script deploys them.
const FILES = ['lib/lib.js', ...scriptsIn('apps/video-assets/src')];

const VIDEO_HEADERS = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Path', 'Duration',
  'Skicka till DMC', 'DMC-status'];

test('lib.js and all app files load together without clashing names', () => {
  const gas = load(FILES);
  ['onOpen', 'listDriveVideos', 'sendCheckedToDmc', 'checkForNewAssets', 'installWatcherTrigger', 'runPostDeployTests']
    .forEach((name) => assert.equal(typeof gas[name], 'function', name));
});

// The DMC job's service account, and the job as the Cloud Run Admin API describes it.
const SA = 'dmc-transfer@sesport.iam.gserviceaccount.com';
const DMC_JOB = {
  name: 'projects/sesport/locations/europe-north1/jobs/dmc-transfer',
  template: { template: { serviceAccount: SA, containers: [{ env: [
    { name: 'FTP_HOST', value: 'ftp-no1.dmcproduction.se' }, { name: 'FTP_USER', value: 'u' },
    { name: 'FTP_PASSWORD', valueSource: { secretKeyRef: { secret: 'dmc-ftp-password', version: 'latest' } } },
    { name: 'SPREADSHEET_ID', value: 'spreadsheet-id' },
  ] }] } },
};

// The media probe job, with its own service account.
const PROBE_SA = 'media-probe@sesport.iam.gserviceaccount.com';
const PROBE_JOB = {
  name: 'projects/sesport/locations/europe-north1/jobs/media-probe',
  template: { template: { serviceAccount: PROBE_SA, containers: [{ env: [{ name: 'SPREADSHEET_ID', value: 'spreadsheet-id' }] }] } },
};

// Answers the Google Cloud requests of the DMC code (reading the job, the caller's permissions on
// it, the password secret's access, starting it),.
function cloudFetch({ job = DMC_JOB, probeJob = PROBE_JOB, granted = true, runCode = 200,
  secretPolicy = { bindings: [{ role: 'roles/secretmanager.secretAccessor', members: ['serviceAccount:' + SA] }] } } = {}) {
  return (url) => {
    if (url.endsWith(':testIamPermissions')) return { code: 200, body: granted ? { permissions: ['run.jobs.runWithOverrides'] } : {} };
    if (url.endsWith(':run')) {
      return runCode === 200 ? { code: 200, body: { metadata: { name: 'executions/e1' } } } : { code: runCode, body: 'Permission denied' };
    }
    if (url.startsWith('https://secretmanager.googleapis.com/')) {
      return secretPolicy ? { code: 200, body: secretPolicy } : { code: 403, body: 'Permission denied' };
    }
    if (url.includes('/jobs/media-probe')) return probeJob ? { code: 200, body: probeJob } : { code: 404, body: 'Job not found' };
    if (url.startsWith('https://run.googleapis.com/')) return job ? { code: 200, body: job } : { code: 404, body: 'Job not found' };
    throw new Error('Unexpected request: ' + url);
  };
}

// The Drive advanced service's sharing: sharing maps a Drive ID (or '*', any other) to its
// permissions; reading the sharing of an ID not in it fails, as for an item one may not manage.
function fakeDriveSharing(sharing) {
  return {
    Permissions: {
      list(id) {
        const permissions = id in sharing ? sharing[id] : sharing['*'];
        if (!permissions) throw new Error('The user does not have sufficient permissions for file ' + id);
        return { permissions };
      },
    },
  };
}
const sharedWith = (role, email = SA) => [{ type: 'user', emailAddress: email, role }];
// The spreadsheet ('spreadsheet-id') and the Video folder ('VID') shared as the job needs them.
const SHARED = {
  'spreadsheet-id': [...sharedWith('writer'), ...sharedWith('writer', PROBE_SA)],
  VID: [...sharedWith('reader'), ...sharedWith('writer', PROBE_SA)], // the probe job stores results there
  INBOXVIDEO: sharedWith('writer', PROBE_SA), // the probe job stores results on the files
};


// --- Video Assets ---------------------------------------------------------------------------

// A project whose Video folder (VIDEO_FOLDER_ID 'VID') holds videos, and whose "Video Assets" sheet
// holds rows (A-F, and G-H when given).
function videoProject({ videos, rows }) {
  const sheet = new FakeSheet('Video Assets', [VIDEO_HEADERS, ...rows]);
  const folder = fakeFolder({ id: 'VID', name: 'Video', files: videos.map((v) => Object.assign({ mimeType: 'video/mp4' }, v)) });
  const gas = load(FILES, {
    spreadsheet: fakeSpreadsheet([sheet]),
    properties: { VIDEO_FOLDER_ID: 'VID', DMC_GCP_PROJECT: 'sesport' },
    folders: { VID: folder },
    fetch: () => ({ code: 200, body: { videoMediaMetadata: { durationMillis: 1000 } } }),
  });
  return { gas, sheet };
}

const sixColumns = (id, name) => [id, name, '', '1', name, '0:00:01:00'];

test('listDriveVideos keeps each row\'s "Skicka till DMC" and "DMC-status" with its video', () => {
  const { gas, sheet } = videoProject({
    videos: [{ id: 'a', name: 'a.mp4' }, { id: 'b', name: 'b.mp4' }, { id: 'c', name: 'c.mp4' }],
    rows: [
      [...sixColumns('b', 'b.mp4'), true, 'Skickad 2026-11-01 10:00'],
      [...sixColumns('gone', 'gone.mp4'), false, 'Fel: x'],
      sixColumns('a', 'a.mp4'), // listed before G-H existed
      [...sixColumns('x', 'x.mp4'), false, ''],
    ],
  });
  sheet.checkboxes.add('4,6'); // the old 4th row's checkbox
  gas.listDriveVideos();
  assert.deepEqual(plain(sheet.data[0]).slice(0, 6), ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Path', 'Duration']);
  assert.equal(sheet.data[1][4], 'a.mp4', 'Path is the path inside the Video folder');
  const rows = plain(sheet.data.slice(1, 5)).map((r) => [r[0], r[6], r[7]]);
  assert.deepEqual(rows, [
    ['a', false, ''],
    ['b', true, 'Skickad 2026-11-01 10:00'],
    ['c', false, ''],
    ['', '', ''], // the rows that are gone are cleared, checkboxes and all
  ]);
  assert.deepEqual([...sheet.checkboxes].sort(), ['1,6', '2,6', '3,6']);
});

function dmcProject(rows, { cloud = {}, sharing = { '*': sharedWith('writer') } } = {}) {
  const sheet = new FakeSheet('Video Assets', [VIDEO_HEADERS, ...rows]);
  const alerts = [];
  const gas = load(FILES, {
    spreadsheet: fakeSpreadsheet([sheet], 'spreadsheet-id', 'Video Assets'),
    properties: { DMC_GCP_PROJECT: 'sesport' },
    fetch: cloudFetch(cloud),
    globals: {
      Session: { getActiveUser: () => ({ getEmail: () => 'producer@example.com' }) },
      Drive: fakeDriveSharing(sharing),
    },
  });
  // The UI, for alerts (the harness spreadsheet has toast but no getUi).
  gas.SpreadsheetApp.getUi = () => ({ alert: (...a) => alerts.push(a), ButtonSet: { OK: 'OK' } });
  return { gas, sheet, alerts };
}

test('"Video Assets" G and H are the columns the DMC code uses', () => {
  const { gas } = dmcProject([]);
  const headers = plain(gas.get('VIDEO_HEADERS'));
  assert.deepEqual(headers, VIDEO_HEADERS);
  assert.equal(headers[gas.get('DMC_SEND_COL') - 1], 'Skicka till DMC');
  assert.equal(headers[gas.get('DMC_STATUS_COL') - 1], 'DMC-status');
});

test('listDriveVideos fails clearly without VIDEO_FOLDER_ID', () => {
  const { gas } = dmcProject([]);
  assert.throws(() => gas.listDriveVideos(), /VIDEO_FOLDER_ID is not set/);
});

// "Video Assets" rows A-H: [id, name, ..., ticked, status]
const videoRow = (id, name, ticked, status) => [...sixColumns(id, name), ticked, status || ''];

test('pickDmcClips_ picks the ticked rows that have a Drive file ID', () => {
  const { gas } = dmcProject([]);
  const pick = gas.pickDmcClips_([
    videoRow('a', 'a.mp4', true),
    videoRow('b', 'b.mp4', false),
    videoRow('', 'no-id.mp4', true),
    videoRow('c', 'c.mp4', true, 'Skickad 2026-11-01 10:00'), // sending again is fine
    videoRow('d', 'd.mp4', 'TRUE'), // text, not a ticked checkbox
  ]);
  assert.deepEqual(plain(pick), {
    clips: [{ row: 2, id: 'a', name: 'a.mp4' }, { row: 5, id: 'c', name: 'c.mp4' }],
    problems: [],
  });
});

test('pickDmcClips_ refuses clips on their way and two clips with the same name', () => {
  const { gas } = dmcProject([]);
  const { problems } = plain(gas.pickDmcClips_([
    videoRow('a', 'a.mp4', true, 'Köad 2026-11-17 14:00'),
    videoRow('b', 'b.mp4', true, 'Skickar 1.2 GB av 4.8 GB (25 %)'),
    videoRow('c1', 'c.mp4', true),
    videoRow('c2', 'c.mp4', true),
  ]));
  assert.equal(problems.length, 3);
  assert.match(problems[0], /^a.mp4 \(rad 2\) är redan på väg: "Köad 2026-11-17 14:00"/);
  assert.match(problems[1], /^b.mp4 \(rad 3\) är redan på väg/);
  assert.match(problems[2], /^c.mp4 finns på rad 4 och 5: samma namn hos DMC/);
});


// --- Sending to DMC ---------------------------------------------------------------------------

test('sendCheckedToDmc starts the job with the ticked clips, one task each, then queues and unticks them', () => {
  const { gas, sheet } = dmcProject([
    videoRow('a', 'a.mp4', true), videoRow('b', 'b.mp4', false), videoRow('c', 'c.mp4', true),
  ]);
  gas.sendCheckedToDmc();
  const req = gas.fakes.urlFetch.requests.find((r) => r.url.endsWith(':run'));
  assert.equal(req.url, 'https://run.googleapis.com/v2/projects/sesport/locations/europe-north1/jobs/dmc-transfer:run');
  assert.deepEqual(JSON.parse(req.options.payload).overrides, {
    containerOverrides: [{ env: [{ name: 'FILE_IDS', value: 'a,c' }, { name: 'REQUESTED_BY', value: 'producer@example.com' }] }],
    taskCount: 2,
  });
  const gh = plain(sheet.data.slice(1)).map((r) => [r[6], r[7]]);
  assert.equal(gh[1].join('|'), 'false|', 'the unticked row is left alone');
  [gh[0], gh[2]].forEach(([ticked, status]) => {
    assert.equal(ticked, false);
    assert.match(status, /^Köad \d{4}-\d{2}-\d{2} \d{2}:\d{2}$/);
  });
});

test('sendCheckedToDmc uses DMC_REGION and DMC_JOB when they are set', () => {
  const { gas } = dmcProject([videoRow('a', 'a.mp4', true)]);
  gas.fakes.properties.setProperty('DMC_REGION', 'europe-west1').setProperty('DMC_JOB', 'dmc-test');
  gas.sendCheckedToDmc();
  assert.ok(gas.fakes.urlFetch.requests.some((r) => /\/locations\/europe-west1\/jobs\/dmc-test:run$/.test(r.url)));
});

test('sendCheckedToDmc keeps the clips ticked, with the reason, when the job does not start', () => {
  const { gas, sheet } = dmcProject([videoRow('a', 'a.mp4', true)], { cloud: { runCode: 403 } });
  assert.throws(() => gas.sendCheckedToDmc(), /Starting Cloud Run job dmc-transfer failed \(403\)/);
  assert.equal(sheet.data[1][6], true);
  assert.match(sheet.data[1][7], /^Fel: överföringen startade inte: Starting Cloud Run job dmc-transfer failed \(403\)/);
});

test('sendCheckedToDmc shows the problems and sends nothing when a clip cannot be sent', () => {
  const { gas, sheet, alerts } = dmcProject([
    videoRow('a', 'a.mp4', true), videoRow('b', 'b.mp4', true, 'Köad 2026-11-17 14:00'),
  ]);
  gas.sendCheckedToDmc();
  assert.equal(gas.fakes.urlFetch.requests.length, 0);
  assert.equal(alerts.length, 1);
  assert.match(alerts[0][1], /b.mp4 \(rad 3\) är redan på väg[\s\S]*Inget skickades\.$/);
  assert.equal(sheet.data[1][6], true, 'nothing was unticked');
});

test('sendCheckedToDmc does nothing when no clip is ticked, and fails clearly without a video list', () => {
  const { gas } = dmcProject([videoRow('a', 'a.mp4', false)]);
  gas.sendCheckedToDmc();
  assert.equal(gas.fakes.urlFetch.requests.length, 0);
  assert.ok(gas.fakes.logs.some((l) => l.includes('Bocka i "Skicka till DMC"')));
  const empty = load(FILES, { spreadsheet: fakeSpreadsheet([]), properties: { DMC_GCP_PROJECT: 'sesport' } });
  assert.throws(() => empty.sendCheckedToDmc(), /Inga videor listade i "Video Assets"/);
});

// --- DMC permissions ------------------------------------------------------------------------------

const TYPO = 'dmc-transfer@project.iam.gserviceaccount.com'; // the mistake made when setting up

test('sendCheckedToDmc sends nothing, and says what to share, when the spreadsheet is not shared with the job', () => {
  const { gas, sheet, alerts } = dmcProject([videoRow('a', 'a.mp4', true)], {
    sharing: { 'spreadsheet-id': sharedWith('writer', TYPO), a: sharedWith('reader') },
  });
  gas.sendCheckedToDmc();
  assert.ok(!gas.fakes.urlFetch.requests.some((r) => r.url.endsWith(':run')), 'the job was not started');
  assert.equal(alerts[0][1], `"Video Assets" är inte delad med ${SA} som redigerare. ` +
    'Dela den så att jobbet kan skriva DMC-status.\n\nInget skickades.');
  assert.deepEqual(plain(sheet.data[1].slice(6)), [true, ''], 'still ticked, not queued');
});

test('sendCheckedToDmc names each clip the job cannot read', () => {
  const { gas, alerts } = dmcProject([videoRow('a', 'a.mp4', true), videoRow('b', 'b.mp4', true)], {
    sharing: { 'spreadsheet-id': sharedWith('writer'), a: sharedWith('reader'), b: sharedWith('reader', TYPO) },
  });
  gas.sendCheckedToDmc();
  assert.match(alerts[0][1], new RegExp(`^"b.mp4" är inte delad med ${SA} som läsare\\. Dela den så att jobbet kan läsa klippet\\.`));
  assert.ok(!alerts[0][1].includes('"a.mp4"'));
});

test('sendCheckedToDmc tells a sender without permission to start the job which role they need', () => {
  const { gas, alerts } = dmcProject([videoRow('a', 'a.mp4', true)], { cloud: { granted: false } });
  gas.sendCheckedToDmc();
  assert.ok(!gas.fakes.urlFetch.requests.some((r) => r.url.endsWith(':run')));
  assert.match(alerts[0][1], /^producer@example.com får inte starta överföringen\. .*"Cloud Run Jobs Executor With Overrides" på jobbet dmc-transfer\./);
});

test('sendCheckedToDmc reports a job it cannot read', () => {
  const { gas, alerts } = dmcProject([videoRow('a', 'a.mp4', true)], { cloud: { job: null } });
  gas.sendCheckedToDmc();
  assert.match(alerts[0][1], /^Kan inte läsa överföringsjobbet dmc-transfer: Reading Cloud Run job dmc-transfer failed \(404\)/);
});

test('sendCheckedToDmc still sends when the sender may not see a clip\'s sharing', () => {
  const { gas, alerts } = dmcProject([videoRow('a', 'a.mp4', true)], {
    sharing: { 'spreadsheet-id': sharedWith('writer') }, // a: sharing not readable
  });
  gas.sendCheckedToDmc();
  assert.equal(alerts.length, 0);
  assert.ok(gas.fakes.urlFetch.requests.some((r) => r.url.endsWith(':run')));
});

test('dmcAccess_ accepts link sharing and a higher role than needed, and matches addresses in any case', () => {
  const { gas } = dmcProject([], {
    sharing: {
      'spreadsheet-id': [{ type: 'user', emailAddress: SA.toUpperCase(), role: 'owner' }],
      a: [{ type: 'anyone', role: 'reader' }],
      b: [{ type: 'user', emailAddress: SA, role: 'commenter' }],
    },
  });
  const access = gas.dmcAccess_([{ id: 'a', name: 'a.mp4', role: 'reader' }, { id: 'b', name: 'b.mp4', role: 'reader' }]);
  assert.deepEqual(plain(access.problems), []);
  assert.equal(access.account, SA);
});

// --- Asset watcher ------------------------------------------------------------------------------

test('checkForNewAssets watches the folder from WATCH_FOLDER_ID and mails WATCH_RECIPIENTS', () => {
  const now = Date.now();
  const folder = fakeFolder({ id: 'W', name: 'Images', files: [{ name: 'new.png', created: now - 60000 }] });
  const gas = load(FILES, {
    folders: { W: folder },
    properties: { WATCH_FOLDER_ID: 'W', WATCH_RECIPIENTS: 'a@example.com, b@example.com' },
  });
  gas.checkForNewAssets();
  const [m] = gas.fakes.mail;
  assert.equal(m.to, 'a@example.com, b@example.com');
  assert.equal(m.subject, '[Drive Alert] New Asset Uploaded: 1 new asset(s) in [Root / Parent Folder]');
  assert.ok(gas.fakes.properties.store.LAST_CHECK_TIME, 'keeps its state under LAST_CHECK_TIME');
});

test('checkForNewAssets uses WATCH_SUBJECT_PREFIX when it is set', () => {
  const folder = fakeFolder({ id: 'W', name: 'Images', files: [{ name: 'new.png', created: Date.now() - 60000 }] });
  const gas = load(FILES, {
    folders: { W: folder },
    properties: { WATCH_FOLDER_ID: 'W', WATCH_RECIPIENTS: 'a@example.com', WATCH_SUBJECT_PREFIX: '[Nytt]' },
  });
  gas.checkForNewAssets();
  assert.match(gas.fakes.mail[0].subject, /^\[Nytt\]: 1 new asset/);
});

test('checkForNewAssets fails clearly, and saves no check time, when its properties are not set', () => {
  const gas = load(FILES, { properties: { WATCH_RECIPIENTS: 'a@example.com' } });
  assert.throws(() => gas.checkForNewAssets(), /The script property WATCH_FOLDER_ID is not set/);
  assert.equal(gas.fakes.properties.store.LAST_CHECK_TIME, undefined);
});

// --- Who runs the watcher (WATCH_RUN_AS) ------------------------------------------------------------

// A project whose watcher runs as `runAs`, run by `me`, who has a checkForNewAssets trigger.
function watcherProject({ runAs, me = 'editor@example.com' }) {
  const folder = fakeFolder({ id: 'W', name: 'Images', files: [{ name: 'new.png', created: Date.now() - 60000 }] });
  const properties = { WATCH_FOLDER_ID: 'W', WATCH_RECIPIENTS: 'a@example.com' };
  if (runAs !== undefined) properties.WATCH_RUN_AS = runAs;
  const gas = load(FILES, {
    folders: { W: folder },
    properties,
    globals: { Session: { getScriptTimeZone: () => 'Europe/Stockholm', getEffectiveUser: () => ({ getEmail: () => me }) } },
  });
  gas.fakes.triggers.push({ getHandlerFunction: () => 'checkForNewAssets', getTriggerSourceId: () => null });
  gas.fakes.triggers.push({ getHandlerFunction: () => 'other', getTriggerSourceId: () => null });
  return gas;
}
const handlers = (gas) => gas.fakes.triggers.map((t) => t.getHandlerFunction());

test('checkForNewAssets run by another account than WATCH_RUN_AS removes that account\'s trigger and does nothing else', () => {
  const gas = watcherProject({ runAs: 'owner@example.com' });
  gas.checkForNewAssets();
  assert.deepEqual(handlers(gas), ['other'], 'only the checkForNewAssets trigger is removed');
  assert.equal(gas.fakes.mail.length, 0);
  assert.equal(gas.fakes.properties.store.LAST_CHECK_TIME, undefined, 'the watcher\'s state is untouched');
  assert.ok(gas.fakes.logs.some((l) => l.includes("runs as owner@example.com, not editor@example.com: removed editor@example.com's checkForNewAssets trigger")));
});

test('checkForNewAssets runs the watcher for WATCH_RUN_AS (any case, spaces trimmed)', () => {
  const gas = watcherProject({ runAs: ' Editor@Example.com ' });
  gas.checkForNewAssets();
  assert.equal(gas.fakes.mail.length, 1);
  assert.deepEqual(handlers(gas), ['checkForNewAssets', 'other']);
});

test('checkForNewAssets runs for everyone without WATCH_RUN_AS, and when Google hides the address', () => {
  const unset = watcherProject({});
  unset.checkForNewAssets();
  assert.equal(unset.fakes.mail.length, 1);
  const hidden = watcherProject({ runAs: 'owner@example.com', me: '' });
  hidden.checkForNewAssets();
  assert.equal(hidden.fakes.mail.length, 1);
  assert.deepEqual(handlers(hidden), ['checkForNewAssets', 'other'], 'the trigger is kept');
});

// --- Post-deploy tests ------------------------------------------------------------------------
// A correctly set up Video Assets project: the sheet, the Video folder, the watched folder, the DMC
// job writing to this spreadsheet, and the runner's checkForNewAssets trigger.
function liveProject({ dmcJob = DMC_JOB, properties = {}, cloud = {}, sharing = SHARED, headers = VIDEO_HEADERS,
  inboxFolders = [{ id: 'INBOXVIDEO', name: 'Video' }, { name: 'Images' }] } = {}) {
  const ss = fakeSpreadsheet([new FakeSheet('Video Assets', [headers, ['a', 'a.mp4']]),
    new FakeSheet('Video Profiles', [PROFILE_HEADERS, ...PROFILES])], 'spreadsheet-id', 'Video Assets');
  const gas = load(FILES, {
    spreadsheet: ss,
    properties: {
      VIDEO_FOLDER_ID: 'VID', IMAGES_FOLDER_ID: 'PRODIMAGES', DMC_GCP_PROJECT: 'sesport',
      WATCH_FOLDER_ID: 'WATCHED', WATCH_RECIPIENTS: 'watch@example.com', WATCH_RUN_AS: 'User@example.com',
      ...properties,
    },
    folders: { VID: fakeFolder({ id: 'VID', name: 'Video', folders: [{ name: 'Editorial' }, { name: 'Commercials' }] }),
      PRODIMAGES: productionFolders().PRODIMAGES,
      WATCHED: fakeFolder({ id: 'WATCHED', name: 'Watched', folders: inboxFolders }) },
    fetch: cloudFetch(Object.assign({ job: dmcJob }, cloud)),
    globals: { Drive: fakeDriveSharing(sharing) },
  });
  gas.installWatcherTrigger();
  gas.installEditTrigger();
  return gas;
}

test('runPostDeployTests passes on a correctly set up project and logs every check', () => {
  const gas = liveProject();
  const summary = gas.runPostDeployTests();
  const failed = gas.fakes.logs.filter((l) => !l.startsWith('PASS') && /^(WARN|FAIL)/.test(l));
  assert.equal(summary, '11 passed, 0 warnings, 0 failed', failed.join('\n'));
  assert.ok(gas.fakes.logs.some((l) => /^PASS +Video Assets sheet +1 videos listed$/.test(l)));
  assert.ok(gas.fakes.logs.some((l) => /^PASS +Drive: watched asset folder +"Watched", mails watch@example.com; lists "Video" in Video Inbox and "Images" in Image Inbox$/.test(l)));
});

test('runPostDeployTests reports every failure in an empty project instead of stopping at the first', () => {
  const gas = load(FILES, { spreadsheet: fakeSpreadsheet([]) });
  assert.throws(() => gas.runPostDeployTests(), /Post-deploy tests failed: Script Properties, Drive: Video folder/);
});

test('runPostDeployTests fails when the sheet\'s columns are not the ones the code uses', () => {
  const headers = VIDEO_HEADERS.slice();
  [headers[6], headers[7]] = [headers[7], headers[6]];
  const gas = liveProject({ headers });
  assert.throws(() => gas.runPostDeployTests(), /Post-deploy tests failed: Video Assets sheet/);
});

test('runPostDeployTests: the DMC check names the job, the server, and that the password is a secret', () => {
  const gas = liveProject();
  gas.runPostDeployTests();
  assert.ok(gas.fakes.logs.some((l) => l.includes('dmc-transfer in sesport/europe-north1 sends to ftp-no1.dmcproduction.se; never run')));
  const req = gas.fakes.urlFetch.requests.find((r) => r.url.startsWith('https://run.googleapis.com/'));
  assert.equal(req.url, 'https://run.googleapis.com/v2/projects/sesport/locations/europe-north1/jobs/dmc-transfer');
  assert.equal(req.options.method, 'get', 'only reads');
});

test('runPostDeployTests: the DMC check fails when the job writes to another spreadsheet', () => {
  const job = JSON.parse(JSON.stringify(DMC_JOB));
  job.template.template.containers[0].env[3].value = 'the-master-schedule';
  const gas = liveProject({ dmcJob: job });
  assert.throws(() => gas.runPostDeployTests(), /Post-deploy tests failed: DMC transfer/);
  assert.ok(gas.fakes.logs.some((l) => l.includes('writes its status to spreadsheet the-master-schedule, not this one (spreadsheet-id); set SPREADSHEET_ID in job.env')));
});

test('runPostDeployTests: the DMC check shows the job\'s last run once it has run', () => {
  const ran = liveProject({ dmcJob: Object.assign({}, DMC_JOB, {
    latestCreatedExecution: { createTime: '2026-11-17T14:20:00Z', completionStatus: 'EXECUTION_SUCCEEDED' },
  }) });
  ran.runPostDeployTests();
  assert.ok(ran.fakes.logs.some((l) => l.endsWith('; last run 2026-11-17T14:20:00Z EXECUTION_SUCCEEDED')));
});

test('runPostDeployTests: the DMC permissions check lists what it verified', () => {
  const gas = liveProject();
  gas.runPostDeployTests();
  assert.ok(gas.fakes.logs.some((l) => /^PASS +DMC transfer: permissions +you may start the job; /.test(l) &&
    l.endsWith(`${SA} can edit this spreadsheet, read "Video" and read the password`)), gas.fakes.logs.join('\n'));
});

test('runPostDeployTests: the DMC permissions check fails on a folder shared with the wrong address', () => {
  const gas = liveProject({ sharing: { 'spreadsheet-id': sharedWith('writer'), VID: sharedWith('reader', TYPO) } });
  assert.throws(() => gas.runPostDeployTests(), /Post-deploy tests failed: DMC transfer: permissions/);
  assert.ok(gas.fakes.logs.some((l) => l.includes(`"Video" is not shared with ${SA} as Viewer`)));
});

test('runPostDeployTests: the DMC permissions check fails when the job may not read the password', () => {
  const gas = liveProject({ cloud: { secretPolicy: { bindings: [
    { role: 'roles/secretmanager.secretAccessor', members: ['serviceAccount:' + TYPO] },
  ] } } });
  assert.throws(() => gas.runPostDeployTests(), /DMC transfer: permissions/);
  assert.ok(gas.fakes.logs.some((l) => l.includes(`${SA} may not read the secret dmc-ftp-password`)));
});

test('runPostDeployTests: the DMC permissions check warns when it cannot see the password\'s access', () => {
  const gas = liveProject({ cloud: { secretPolicy: null } });
  assert.equal(gas.runPostDeployTests(), '10 passed, 1 warnings, 0 failed');
  assert.ok(gas.fakes.logs.some((l) => /^WARN +DMC transfer: permissions .*read "Video"; secret dmc-ftp-password not checked/.test(l)));
});

test('runPostDeployTests: the DMC permissions check fails when the runner may not start the job', () => {
  const gas = liveProject({ cloud: { granted: false } });
  assert.throws(() => gas.runPostDeployTests(), /DMC transfer: permissions/);
  assert.ok(gas.fakes.logs.some((l) => l.includes('user@example.com may not start the job; grant roles/run.jobsExecutorWithOverrides')));
});

test('installWatcherTrigger replaces the runner\'s watcher triggers with one every 10 minutes', () => {
  const gas = liveProject();
  gas.installWatcherTrigger();
  assert.deepEqual(gas.fakes.triggers.map((t) => t.getHandlerFunction()).sort(), ['checkForNewAssets', 'handleEdit'], 'handleEdit is left alone');
  assert.ok(gas.fakes.logs.includes('checkForNewAssets now runs every 10 minutes as user@example.com'));
});

test('runPostDeployTests: the trigger check follows WATCH_RUN_AS', () => {
  const other = liveProject({ properties: { WATCH_RUN_AS: 'owner@example.com' } });
  other.fakes.triggers.splice(0, 1); // the watcher's trigger is the owner's
  assert.equal(other.runPostDeployTests(), '11 passed, 0 warnings, 0 failed');
  assert.ok(other.fakes.logs.some((l) => l.includes('not checkForNewAssets (WATCH_RUN_AS: the watcher runs as owner@example.com)')));

  const mine = liveProject();
  mine.runPostDeployTests();
  assert.ok(mine.fakes.logs.some((l) => /^PASS +Triggers +handleEdit and checkForNewAssets \(WATCH_RUN_AS: the watcher runs as you, User@example.com\)$/.test(l)));

  const missing = liveProject();
  missing.fakes.triggers.splice(0, 1);
  missing.runPostDeployTests();
  assert.ok(missing.fakes.logs.some((l) => /^WARN +Triggers +checkForNewAssets is not installed by you \(run installWatcherTrigger\)/.test(l)));

  const leftover = liveProject({ properties: { WATCH_RUN_AS: 'owner@example.com' } });
  leftover.runPostDeployTests();
  assert.ok(leftover.fakes.logs.some((l) => /^WARN +Triggers +you have a checkForNewAssets trigger, but the watcher runs as owner@example.com/.test(l)));

  const unset = liveProject({ properties: { WATCH_RUN_AS: '' } });
  unset.runPostDeployTests();
  assert.ok(unset.fakes.logs.some((l) => /^WARN +Triggers +WATCH_RUN_AS is not set/.test(l)));
});

test('runPostDeployTests lists the Script Properties by name and warns about unused ones', () => {
  const gas = liveProject({ properties: { LAST_CHECK_TIME: '2026-11-20T10:00:00Z' } });
  gas.runPostDeployTests();
  assert.match(gas.fakes.logs.find((l) => l.includes('Script Properties')),
    /^PASS +Script Properties +7 in all: the 5 required; optional WATCH_RUN_AS; kept by the script LAST_CHECK_TIME$/);
  assert.ok(!gas.fakes.logs.join('\n').includes('2026-11-20T10:00:00Z'), 'no values in the log');

  const old = liveProject({ properties: { DRIVE_FOLDER_ID: 'x' } });
  assert.match(old.runPostDeployTests(), /1 warnings/);
  assert.ok(old.fakes.logs.some((l) => /^WARN +Script Properties .*; not used by the code: DRIVE_FOLDER_ID$/.test(l)));
});

// --- Inbox tabs ------------------------------------------------------------------------------------

const PROBE_COLUMNS = ['Frame rate', 'Width', 'Height', 'Codec', 'Mbit', 'Scan', 'Bit depth', 'Alpha', 'Audio', 'Probe status'];
const MINUTE_AGO = Date.now() - 60000;
const LONG_AGO = Date.now() - 30 * 86400000;

// The production folders, with the subfolders they have in Drive.
const productionFolders = () => ({
  PRODVIDEO: fakeFolder({ id: 'PRODVIDEO', name: 'Video', folders: [{ name: 'Editorial' }, { name: 'Commercials' }, { name: 'Bumpers' }] }),
  PRODIMAGES: fakeFolder({ id: 'PRODIMAGES', name: 'Images', folders: [{ name: 'Episodes' }, { name: 'Series' }, { name: 'Genre' }] }),
});
const PRODUCTION_PROPERTIES = { VIDEO_FOLDER_ID: 'PRODVIDEO', IMAGES_FOLDER_ID: 'PRODIMAGES' };

// The Inbox (WATCH_FOLDER_ID 'INBOX') with Video and Images subfolders, files given per folder.
function inboxProject({ video = [], images = [], root = [], subfolders = ['Video', 'Images'], sheets = [], runCode = 200 } = {}) {
  const content = { Video: video, Images: images };
  const inbox = fakeFolder({ id: 'INBOX', name: 'Inbox', files: root,
    folders: subfolders.map((name) => ({ name, files: content[name].filter((f) => !f.sub),
      folders: content[name].some((f) => f.sub) ? [{ name: 'Matcher', files: content[name].filter((f) => f.sub) }] : [] })) });
  const ss = fakeSpreadsheet(sheets, 'spreadsheet-id', 'seSPORT AssetManager');
  const folders = { INBOX: inbox, ...productionFolders() };
  const gas = load(FILES, {
    spreadsheet: ss,
    properties: { WATCH_FOLDER_ID: 'INBOX', WATCH_RECIPIENTS: 'a@example.com', DMC_GCP_PROJECT: 'sesport', ...PRODUCTION_PROPERTIES },
    folders,
    globals: { Drive: fakeDriveFiles(folders) },
    fetch: (url) => {
      // Starting the probe job.
      if (url.endsWith(':run')) return runCode === 200 ? { code: 200, body: { metadata: { name: 'executions/p1' } } } : { code: runCode, body: 'Permission denied' };
      // Drive has a duration for the files whose IDs start with v (the videos), none for the others.
      return { code: 200, body: /\/files\/v/.test(url) ? { videoMediaMetadata: { durationMillis: 2000 } } : {} };
    },
  });
  const tab = (name) => { const t = ss.getSheetByName(name); return t ? plain(t.data) : null; };
  return { gas, tab };
}

test('a new file in Inbox/Video lists all of Inbox/Video in "Video Inbox", sorted, with durations', () => {
  const { gas, tab } = inboxProject({
    video: [
      { id: 'v2', name: 'b.mov', mimeType: 'video/quicktime', created: MINUTE_AGO, updated: MINUTE_AGO, size: 2048 },
      { id: 'v1', name: 'a.mp4', mimeType: 'video/mp4', created: LONG_AGO, updated: LONG_AGO, size: 1024 },
      { id: 'v3', name: 'c.mp4', mimeType: 'video/mp4', created: LONG_AGO },
      { id: 'n1', name: 'notes.txt', mimeType: 'text/plain', created: LONG_AGO },
    ],
    images: [{ id: 'i1', name: 'x.png', mimeType: 'image/png', created: LONG_AGO }],
  });
  gas.checkForNewAssets();
  assert.equal(gas.fakes.mail.length, 1, 'the mail still goes out');
  const rows = tab('Video Inbox');
  assert.deepEqual(rows[0], ['Drive file ID', 'Name', 'Drive modified', 'Action', 'Size', 'Duration', ...PROBE_COLUMNS]);
  assert.deepEqual(rows.slice(1).map((r) => [r[0], r[1], r[5]]), [
    ['v1', 'a.mp4', 2 / 86400], // Drive's 2000 ms, as a Sheets duration (until probed)
    ['v2', 'b.mov', 2 / 86400],
    ['v3', 'c.mp4', 2 / 86400],
    ['n1', 'notes.txt', ''], // every file is listed; only videos get a duration
  ]);
  const shown = gas.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Video Inbox').getRange(2, 6, 1, 1).getDisplayValues();
  assert.deepEqual(shown, [['0:00:02']], 'shown as h:mm:ss');
  assert.equal(rows[2][4], '2048');
  assert.equal(tab('Image Inbox'), null, 'Images had nothing new');
  assert.ok(gas.fakes.logs.includes('Video Inbox: 4 files'));
});

test('a new file in Inbox/Images lists Inbox/Images in "Image Inbox", without a Duration column', () => {
  const { gas, tab } = inboxProject({
    video: [{ id: 'v1', name: 'a.mp4', mimeType: 'video/mp4', created: LONG_AGO }],
    images: [{ id: 'i2', name: 'b.png', mimeType: 'image/png', created: MINUTE_AGO },
             { id: 'i1', name: 'a.jpg', mimeType: 'image/jpeg', created: LONG_AGO }],
  });
  gas.checkForNewAssets();
  assert.deepEqual(tab('Image Inbox').map((r) => [r[0], r[1], r.filter((c) => c !== '').length]), [
    ['Drive file ID', 'Name', 5], ['i1', 'a.jpg', 4], ['i2', 'b.png', 4], // Action is empty
  ]);
  assert.equal(tab('Video Inbox'), null);
  assert.equal(gas.fakes.urlFetch.requests.length, 0, 'no duration lookups for images');
});

test('new files in both subfolders refresh both tabs; a new file directly in Inbox refreshes none', () => {
  const both = inboxProject({
    video: [{ id: 'v1', name: 'a.mp4', mimeType: 'video/mp4', created: MINUTE_AGO }],
    images: [{ id: 'i1', name: 'a.png', mimeType: 'image/png', created: MINUTE_AGO }],
  });
  both.gas.checkForNewAssets();
  assert.equal(both.tab('Video Inbox').length, 2);
  assert.equal(both.tab('Image Inbox').length, 2);

  const root = inboxProject({ root: [{ id: 'r', name: 'loose.mp4', created: MINUTE_AGO }] });
  root.gas.checkForNewAssets();
  assert.equal(root.gas.fakes.mail.length, 1);
  assert.equal(root.tab('Video Inbox'), null);
  assert.equal(root.tab('Image Inbox'), null);
});

test('nothing new: no mail and the tabs are left alone', () => {
  const { gas, tab } = inboxProject({ video: [{ id: 'v1', name: 'a.mp4', created: LONG_AGO }] });
  gas.checkForNewAssets();
  assert.equal(gas.fakes.mail.length, 0);
  assert.equal(tab('Video Inbox'), null);
});

test('refreshInboxLists rebuilds both tabs, clearing rows of files that are gone', () => {
  // In the layout before Path was dropped: the old column E goes too.
  const old = new FakeSheet('Image Inbox', [['Drive file ID', 'Name', 'Drive modified', 'Size', 'Path'],
    ['gone1', 'old.png', '', '1', 'old.png'], ['gone2', 'older.png', '', '1', 'older.png']]);
  const { gas, tab } = inboxProject({
    video: [{ id: 'v1', name: 'a.mp4', mimeType: 'video/mp4', created: LONG_AGO }],
    images: [{ id: 'i1', name: 'a.png', mimeType: 'image/png', created: LONG_AGO }],
    sheets: [old],
  });
  gas.refreshInboxLists();
  assert.deepEqual(tab('Image Inbox').slice(1).map((r) => r[0]).filter(Boolean), ['i1'], 'the old rows are cleared');
  assert.deepEqual(tab('Image Inbox')[0], ['Drive file ID', 'Name', 'Drive modified', 'Action', 'Size'], 'the old Path column is gone');
  assert.ok(tab('Image Inbox').every((r) => !r[5]), 'nothing is left right of Size');
  assert.deepEqual(tab('Video Inbox').slice(1).map((r) => r[0]), ['v1']);
  assert.ok(gas.fakes.logs.includes('Video Inbox: 1 filer (1 analyseras), Image Inbox: 1 filer'));
});

test('refreshInboxLists says which subfolder is missing instead of failing', () => {
  const { gas, tab } = inboxProject({ subfolders: ['Video'] });
  gas.refreshInboxLists();
  assert.ok(gas.fakes.logs.includes('Video Inbox: 0 filer, Image Inbox: ingen mapp "Images" i "Inbox"'));
  assert.equal(tab('Image Inbox'), null);
});

test('runPostDeployTests warns when the watched folder has no Video or Images subfolder', () => {
  const gas = liveProject({ inboxFolders: [{ id: 'INBOXVIDEO', name: 'Video' }] });
  assert.equal(gas.runPostDeployTests(), '10 passed, 1 warnings, 0 failed');
  assert.ok(gas.fakes.logs.some((l) => /^WARN +Drive: watched asset folder .*; no subfolder "Images" \(for Image Inbox\)$/.test(l)));
});

test('the mail goes out after the tabs are updated and links to each updated tab', () => {
  const both = inboxProject({
    video: [{ id: 'v1', name: 'a.mp4', mimeType: 'video/mp4', created: MINUTE_AGO }],
    images: [{ id: 'i1', name: 'a.png', mimeType: 'image/png', created: MINUTE_AGO },
             { id: 'i2', name: 'b.png', mimeType: 'image/png', created: LONG_AGO }],
  });
  const ss = both.gas.SpreadsheetApp.getActiveSpreadsheet();
  both.gas.checkForNewAssets();
  const url = (name) => 'https://docs.google.com/spreadsheets/d/spreadsheet-id/edit#gid=' + ss.getSheetByName(name).getSheetId();
  const body = both.gas.fakes.mail[0].body;
  assert.ok(body.startsWith('Hello,\n\nThis is an automated message from the seSPORT Asset Manager\n' +
    '2 new asset(s) were uploaded'), body);
  assert.ok(body.includes('Updated in seSPORT AssetManager:\n' +
    `  Video Inbox (1 file): ${url('Video Inbox')} (probing 1 video; the results appear in the tab)\n` +
    `  Image Inbox (2 files): ${url('Image Inbox')}\n\n=====`), body);
  assert.ok(body.indexOf('Updated in') < body.indexOf('SUBFOLDER:'), 'the links come before the file list');

  const one = inboxProject({ images: [{ id: 'i1', name: 'a.png', mimeType: 'image/png', created: MINUTE_AGO }] });
  one.gas.checkForNewAssets();
  assert.ok(one.gas.fakes.mail[0].body.includes('  Image Inbox (1 file): '));
  assert.ok(!one.gas.fakes.mail[0].body.includes('Video Inbox'), 'only the tab that was updated');
});

test('when a tab cannot be updated, the mail still goes out, says so, and has no link for it', () => {
  const { gas } = inboxProject({
    video: [{ id: 'v1', name: 'a.mp4', mimeType: 'video/mp4', created: MINUTE_AGO }],
    images: [{ id: 'i1', name: 'a.png', mimeType: 'image/png', created: MINUTE_AGO }],
  });
  gas.UrlFetchApp.fetchAll = () => { throw new Error('Drive is down'); }; // durations fail: Video Inbox only
  gas.checkForNewAssets();
  const body = gas.fakes.mail[0].body;
  assert.match(body, /Image Inbox \(1 file\): https:/);
  assert.match(body, /Video Inbox could not be updated: Drive is down/);
  assert.ok(!/Video Inbox \(/.test(body));
});

test('a new file directly in Inbox gives the usual mail, with nothing about the tabs', () => {
  const { gas } = inboxProject({ root: [{ id: 'r', name: 'loose.mp4', created: MINUTE_AGO }] });
  gas.checkForNewAssets();
  assert.match(gas.fakes.mail[0].body, /in "Inbox"\.\n\n=====/);
});

// --- Probing the videos in Video Inbox (mediaProbe.js) ------------------------------------------------

const runRequests = (gas) => gas.fakes.urlFetch.requests.filter((r) => r.url.endsWith(':run'));
const O = 15; // "Probe status", column P

test('the videos not probed yet are sent to the probe job, one execution, and marked Queued', () => {
  const { gas, tab } = inboxProject({ video: [
    { id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: MINUTE_AGO },
    { id: 'v2', name: 'b.mxf', mimeType: 'application/mxf', created: LONG_AGO }, // Drive's type for MXF
    { id: 'n1', name: 'notes.txt', mimeType: 'text/plain', created: LONG_AGO },
  ] });
  gas.checkForNewAssets();
  const [req] = runRequests(gas);
  assert.equal(runRequests(gas).length, 1);
  assert.match(req.url, /\/projects\/sesport\/locations\/europe-north1\/jobs\/media-probe:run$/);
  assert.deepEqual(JSON.parse(req.options.payload).overrides, {
    containerOverrides: [{ env: [{ name: 'FILE_IDS', value: 'v1,v2' }, { name: 'REQUESTED_BY', value: 'user@example.com' }] }],
    taskCount: 1,
  });
  const rows = tab('Video Inbox').slice(1);
  assert.deepEqual(rows.map((r) => [r[0], String(r[O] || '').replace(/\d{4}-\d\d-\d\d \d\d:\d\d$/, 'T')]),
    [['v1', 'Queued T'], ['v2', 'Queued T'], ['n1', 'Not a video']], 'the text file is not probed');
});

test('probe results stay with their video when the tab is rebuilt, and probed videos are not sent again', () => {
  // Video Inbox in an older layout: with Path, without Mbit. The results must land
  // under their own headers in the new layout.
  const OLD_PROBE = ['Frame rate', 'Width', 'Height', 'Codec', 'Scan', 'Bit depth', 'Alpha', 'Audio', 'Probe status'];
  const probedOld = ['25', 1920, 1080, 'ProRes 422 HQ', 'Interlaced TFF', 10, 'No', '8 × mono (8 ch)', 'Probed 2026-11-21 17:02'];
  const old = new FakeSheet('Video Inbox', [['Drive file ID', 'Name', 'Drive modified', 'Size', 'Path', 'Duration', ...OLD_PROBE],
    ['v9', 'z.mov', '', '1', 'z.mov', '', ...probedOld],
    ['gone', 'old.mov', '', '1', 'old.mov', '', ...probedOld]]);
  const { gas, tab } = inboxProject({ sheets: [old], video: [
    { id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: MINUTE_AGO },
    { id: 'v9', name: 'z.mov', mimeType: 'video/quicktime', created: LONG_AGO },
  ] });
  gas.refreshInboxLists();
  const rows = tab('Video Inbox');
  assert.deepEqual(rows[0], ['Drive file ID', 'Name', 'Drive modified', 'Action', 'Size', 'Duration', ...PROBE_COLUMNS]);
  assert.deepEqual(rows[2].slice(0, 1).concat(rows[2].slice(6)),
    ['v9', '25', 1920, 1080, 'ProRes 422 HQ', '', 'Interlaced TFF', 10, 'No', '8 × mono (8 ch)', 'Probed 2026-11-21 17:02'],
    'z.mov moved down a row with its results, each under its header; no Mbit until probed again');
  assert.match(rows[1][O], /^Queued /);
  assert.equal(rows[1][9], '', 'a.mov has no results yet');
  assert.ok(!rows[3] || rows[3].every((c) => c === ''), 'the row of the file that is gone is cleared');
  assert.equal(JSON.parse(runRequests(gas)[0].options.payload).overrides.containerOverrides[0].env[0].value, 'v1');
});

test('nothing to probe: no job is started', () => {
  const { gas } = inboxProject({ video: [{ id: 'n1', name: 'notes.txt', mimeType: 'text/plain', created: MINUTE_AGO }] });
  gas.checkForNewAssets();
  assert.equal(runRequests(gas).length, 0);
  assert.ok(!gas.fakes.mail[0].body.includes('probing'));
});

test('when the probe job does not start, the statuses stay empty (so the next rebuild retries) and the mail says why', () => {
  const { gas, tab } = inboxProject({ runCode: 403, video: [{ id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: MINUTE_AGO }] });
  gas.checkForNewAssets();
  assert.equal(tab('Video Inbox')[1][O] || '', '');
  const body = gas.fakes.mail[0].body;
  assert.match(body, /Video Inbox \(1 file\): https:\S+\n/);
  assert.match(body, /The videos in Video Inbox could not be probed: Starting Cloud Run job media-probe failed \(403\)/);

  const manual = inboxProject({ runCode: 403, video: [{ id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: LONG_AGO }] });
  manual.gas.refreshInboxLists();
  assert.ok(manual.gas.fakes.logs.some((l) => l.startsWith('Video Inbox: 1 filer (analysen startade inte: Starting Cloud Run job media-probe failed (403)')));
});

test('PROBE_JOB names another job', () => {
  const { gas } = inboxProject({ video: [{ id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: MINUTE_AGO }] });
  gas.fakes.properties.setProperty('PROBE_JOB', 'probe-test');
  gas.checkForNewAssets();
  assert.match(runRequests(gas)[0].url, /\/jobs\/probe-test:run$/);
});

test('runPostDeployTests: the probe job checks name the job and what its account may do', () => {
  const gas = liveProject();
  gas.runPostDeployTests();
  assert.ok(gas.fakes.logs.some((l) => /^PASS +Media probe: Cloud Run job +media-probe in sesport\/europe-north1 probes the videos in Video Inbox; never run$/.test(l)));
  assert.ok(gas.fakes.logs.some((l) => l.endsWith(`you may start the job; ${PROBE_SA} can edit this spreadsheet, edit Inbox/"Video" and edit "Video"`)), gas.fakes.logs.join('\n'));
});

test('runPostDeployTests: the probe checks fail when Inbox/Video is not shared with the probe account, or the job writes elsewhere', () => {
  const unshared = liveProject({ sharing: Object.assign({}, SHARED, { INBOXVIDEO: sharedWith('reader', PROBE_SA) }) });
  assert.throws(() => unshared.runPostDeployTests(), /Post-deploy tests failed: Media probe: permissions$/);
  assert.ok(unshared.fakes.logs.some((l) => l.includes(`"Video" is not shared with ${PROBE_SA} as Editor (for media-probe)`)), 'Viewer is not enough any more');

  const job = JSON.parse(JSON.stringify(PROBE_JOB));
  job.template.template.containers[0].env[0].value = 'elsewhere';
  const elsewhere = liveProject({ cloud: { probeJob: job } });
  assert.throws(() => elsewhere.runPostDeployTests(), /Media probe: Cloud Run job/);
  assert.ok(elsewhere.fakes.logs.some((l) => l.includes('set SPREADSHEET_ID in job.env and run scripts/deploy-media-probe.sh')));
});

test('Duration: Drive\'s until the probe has written its own, which is then kept; an MXF waits for the probe', () => {
  const header = ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Duration', ...PROBE_COLUMNS];
  // v1 probed at 24 fps (Drive's 25 fps timecode would differ); v2 an MXF Drive has no duration for.
  const probedRow = ['v1', 'a.mp4', '', '1', '0:00:14:12', 24, 1920, 1080, 'H.264 Main', 14.9, 'Progressive', 8, 'No', 'stereo (2 ch)', 'Probed 2026-11-21 17:31'];
  const { gas, tab } = inboxProject({ sheets: [new FakeSheet('Video Inbox', [header, probedRow])], video: [
    { id: 'v1', name: 'a.mp4', mimeType: 'video/mp4', created: LONG_AGO },
    { id: 'm2', name: 'b.mxf', mimeType: 'application/mxf', created: LONG_AGO },
    { id: 'v3', name: 'c.mp4', mimeType: 'video/mp4', created: LONG_AGO },
  ] });
  gas.refreshInboxLists();
  const rows = tab('Video Inbox').slice(1);
  assert.deepEqual(rows.map((r) => [r[0], r[5]]), [
    ['v1', 14 / 86400], // the probe's (an older timecode, 0:00:14:12, frames dropped), not Drive's 2 s
    ['m2', ''],         // nothing from Drive for MXF: the probe fills it in
    ['v3', 2 / 86400],  // Drive's, until probed
  ]);
  assert.equal(JSON.parse(runRequests(gas)[0].options.payload).overrides.containerOverrides[0].env[0].value, 'm2,v3');
});

// --- Names link to their files ---------------------------------------------------------------------

const DRIVE_URL = (id) => `https://drive.google.com/file/d/${id}/view`;

test('the Drive modified time in Video Inbox and Image Inbox links to the file\'s preview, blue without underline', () => {
  const { gas, tab } = inboxProject({
    video: [{ id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: LONG_AGO }],
    images: [{ id: 'i1', name: 'b.png', mimeType: 'image/png', created: LONG_AGO }, { id: 'i2', name: 'c.png', mimeType: 'image/png', created: LONG_AGO }],
  });
  gas.refreshInboxLists();
  const ss = gas.SpreadsheetApp.getActiveSpreadsheet();
  assert.deepEqual([...ss.getSheetByName('Video Inbox').links], [['1,2', DRIVE_URL('v1')]]);
  assert.deepEqual([...ss.getSheetByName('Image Inbox').links], [['1,2', DRIVE_URL('i1')], ['2,2', DRIVE_URL('i2')]]);
  assert.deepEqual(ss.getSheetByName('Image Inbox').linkStyles.get('1,2'), { underline: false, bold: true, color: '#1155cc' });
  assert.equal(tab('Image Inbox')[1][2], new Date(LONG_AGO).toISOString(), 'the value is the time itself');
});

test('a link goes with its row: rows of files that are gone lose theirs; names linked before lose their link', () => {
  const old = new FakeSheet('Image Inbox', [['Drive file ID', 'Name', 'Drive modified', 'Size'], ['i1', 'b.png', '', '1']]);
  old.links.set('1,1', DRIVE_URL('i1')); // linked on the name, before the link moved
  const { gas } = inboxProject({ sheets: [old], images: [{ id: 'i1', name: 'b.png', mimeType: 'image/png', created: LONG_AGO }, { id: 'i2', name: 'c.png', mimeType: 'image/png', created: LONG_AGO }] });
  gas.refreshInboxLists();
  assert.deepEqual([...old.links].sort(), [['1,2', DRIVE_URL('i1')], ['2,2', DRIVE_URL('i2')]]);
  // i1 is deleted from Drive.
  gas.DriveApp.getFolderById('INBOX').getFoldersByName('Images').next().getFiles = (() => {
    let done = false;
    return () => ({ hasNext: () => !done, next: () => { done = true; return { getId: () => 'i2', getName: () => 'c.png', getMimeType: () => 'image/png', getLastUpdated: () => new Date(0), getSize: () => 1 }; } });
  })();
  gas.refreshInboxLists();
  assert.deepEqual([...old.links], [['1,2', DRIVE_URL('i2')]]);
});

test('Video Assets: the Drive modified time links to the preview; sending to DMC still reads the plain name', () => {
  const { gas, sheet } = videoProject({ videos: [{ id: 'a', name: 'a.mp4' }], rows: [] });
  gas.listDriveVideos();
  assert.deepEqual([...sheet.links], [['1,2', DRIVE_URL('a')]]);
  assert.equal(plain(gas.pickDmcClips_(plain(sheet.data.slice(1)).map((r) => (r[6] = true, r)))).clips[0].name, 'a.mp4');
});

// --- Durations are Sheets durations ------------------------------------------------------------------

test('a probed duration survives rebuilds as a duration: read as shown (h:mm:ss), stored as a number', () => {
  const { gas, tab } = inboxProject({ video: [{ id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: LONG_AGO }] });
  gas.refreshInboxLists();
  const sheet = gas.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Video Inbox');
  // The job writes 1:30:05 (as a number, RAW); Apps Script would read it back as a Date.
  sheet.getRange(2, 6).setValue(5405 / 86400);
  sheet.getRange(2, 16).setValue('Probed 2026-11-21 18:00');
  gas.refreshInboxLists();
  gas.refreshInboxLists();
  assert.equal(tab('Video Inbox')[1][5], 5405 / 86400);
  assert.deepEqual(sheet.getRange(2, 6).getDisplayValues(), [['1:30:05']]);
  assert.equal(sheet.formats.get('1,5'), '[h]:mm:ss');
});

test('Video Assets: Drive\'s duration as a Sheets duration, in whole seconds', () => {
  const { gas, sheet } = videoProject({ videos: [{ id: 'a', name: 'a.mp4' }], rows: [] });
  gas.listDriveVideos();
  assert.equal(sheet.data[1][5], 1 / 86400, 'the fake Drive says 1000 ms');
  assert.deepEqual(sheet.getRange(2, 6).getDisplayValues(), [['0:00:01']]);
  assert.equal(sheet.data[1][4], 'a.mp4', 'Path is still text in E');
});

test('Video Assets lists MXF clips (Drive calls them application/mxf), and no other files', () => {
  const { gas, sheet } = videoProject({ videos: [
    { id: 'a', name: 'a.mp4' },
    { id: 'b', name: 'b.mxf', mimeType: 'application/mxf' },
    { id: 'c', name: 'notes.txt', mimeType: 'text/plain' },
  ], rows: [] });
  gas.listDriveVideos();
  assert.deepEqual(plain(sheet.data.slice(1)).map((r) => r[1]), ['a.mp4', 'b.mxf']);
});

test('parseDuration_ and durationDays_: h:mm:ss or a timecode, frames and fractions dropped', () => {
  const { gas } = inboxProject();
  assert.equal(gas.parseDuration_('1:30:05'), 5405 / 86400);
  assert.equal(gas.parseDuration_('0:01:30:24'), 90 / 86400);
  assert.equal(gas.parseDuration_('12:00:00'), 0.5);
  ['', 'Probing', '1:30', null].forEach((x) => assert.equal(gas.parseDuration_(x), '', String(x)));
  assert.equal(gas.durationDays_(89.96), 89 / 86400);
  assert.equal(gas.durationDays_(''), '');
});

// --- Validation against Video Profiles (validation.js) --------------------------------------------------

// The profiles as set up in the sheet (2026-11-21). Durations as typed: shown as h:mm:ss.
const PROFILE_HEADERS = ['Profile name', 'Max duration', 'Frame rate', 'Width', 'Height', 'Codec', 'Min MBit', 'Max Mbit',
  'Scan', 'Bit depth', 'Alpha', 'Audio'];
const PROFILES = [
  ['Graphics with alpha', '0:00:15', 50, 1920, 1080, 'ProRes 4444', 100, 150, 'Progressive', '8/10', 'Yes', 'none / stereo (2 ch)'],
  ['Graphics', '0:00:30', 50, 1920, 1080, 'ProRes 422', 75, 125, 'Progressive', '8/10', 'No', 'none / stereo (2 ch)'],
  ['Promos / Commercials', '0:01:30', 50, 1920, 1080, 'ProRes 422', 50, 100, 'Progressive', '8/10', 'No', 'stereo (2 ch)'],
  ['Editorial short stories', '0:15:00', 50, 1920, 1080, 'H.264 Main', 35, 50, 'Progressive', 8, 'No', 'stereo (2 ch)'],
  ['Editorial', '2:00:00', 50, 1920, 1080, 'H.264 Main', 25, 40, 'Progressive', 8, 'No', 'stereo (2 ch)'],
];
const INBOX_COLUMNS = ['Drive file ID', 'Name', 'Drive modified', 'Action', 'Size', 'Duration', ...PROBE_COLUMNS];

// A probed video as Video Inbox holds it: o overrides a valid "Graphics" ProRes 422 clip.
function probedVideo(id, o = {}) {
  const v = Object.assign({ seconds: 20, fps: 50, w: 1920, h: 1080, codec: 'ProRes 422', mbit: 80, scan: 'Progressive',
    bits: 10, alpha: 'No', audio: 'stereo (2 ch)', status: 'Probed 2026-11-21 18:00' }, o);
  return [id, id + '.mov', '', '', '1', v.seconds / 86400, v.fps, v.w, v.h, v.codec, v.mbit, v.scan, v.bits, v.alpha, v.audio, v.status];
}

function validationProject(videos, profiles = PROFILES) {
  const inbox = new FakeSheet('Video Inbox', [INBOX_COLUMNS, ...videos]);
  videos.forEach((_, i) => inbox.formats.set(`${i + 1},5`, '[h]:mm:ss'));
  const profileSheet = new FakeSheet('Video Profiles', [PROFILE_HEADERS, ...profiles]);
  const gas = load(FILES, { spreadsheet: fakeSpreadsheet([inbox, profileSheet], 'spreadsheet-id', 'seSPORT AssetManager'),
    properties: PRODUCTION_PROPERTIES, folders: productionFolders() });
  const marks = (row) => {
    const out = {};
    INBOX_COLUMNS.forEach((h, c) => {
      const key = `${row},${c}`;
      const m = [inbox.cellBackgrounds.get(key), inbox.notes.get(key)].filter(Boolean);
      if (m.length) out[h] = m.join(' | ');
    });
    return out;
  };
  return { gas, inbox, profileSheet, marks };
}

const GREY = '#efefef';
const RED = '#f4cccc';
// What a failed row's marks look like: every cell grey, the given ones red with their notes.
const failed = (red) => {
  const out = {};
  INBOX_COLUMNS.forEach((h) => { out[h] = GREY; });
  Object.keys(red).forEach((h) => { out[h] = RED + ' | ' + red[h]; });
  return out;
};

test('a video meeting profiles gets a note on its name listing them, and no colours', () => {
  const { gas, inbox, marks } = validationProject([
    probedVideo('g'), // 20 s, 80 Mbit ProRes 422: Graphics and Promos
    probedVideo('a', { codec: 'ProRes 4444', alpha: 'Yes', mbit: 120, seconds: 10, audio: 'none' }),
    probedVideo('e', { codec: 'H.264 Main', bits: 8, mbit: 38, seconds: 600 }), // both editorial profiles
  ]);
  assert.deepEqual(plain(gas.validateVideoInbox_()), { approved: 3, failed: 0, unprobed: 0 });
  assert.deepEqual(marks(1), { Name: 'Approved for:\nGraphics\nPromos / Commercials' });
  assert.deepEqual(marks(2), { Name: 'Approved for:\nGraphics with alpha' });
  assert.deepEqual(marks(3), { Name: 'Approved for:\nEditorial short stories\nEditorial' });
  assert.equal(inbox.fontColors.size, 0);
});

test('a rule every profile shares: "<name> should be <value>"', () => {
  const { gas, marks, inbox } = validationProject([probedVideo('x', { fps: 25 })]);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), failed({ 'Frame rate': 'Frame rate should be 50' }));
  assert.equal(inbox.fontColors.get('1,0'), '#666666', 'dark grey text');
});

test('a codec no profile allows: "<codec> is not an allowed codec."', () => {
  const { gas, marks } = validationProject([probedVideo('x', { codec: 'VC3/DNxHD DNXHR HQX' })]);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), failed({ Codec: 'VC3/DNxHD DNXHR HQX is not an allowed codec.' }));
});

test('what the profiles with the codec require: the widest limit, "… when using <codec>"', () => {
  const { gas, marks } = validationProject([
    probedVideo('long', { codec: 'H.264 Main', bits: 8, mbit: 38, seconds: 3 * 3600 }),
    probedVideo('fat', { mbit: 150 }),
    probedVideo('thin', { mbit: 40, bits: 12 }),
  ]);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), failed({ Duration: 'Max duration is 2:00:00 when using H.264 Main' }));
  assert.deepEqual(marks(2), failed({ Mbit: 'Max Mbit is 125 when using ProRes 422' }));
  assert.deepEqual(marks(3), failed({ Mbit: 'Min MBit is 50 when using ProRes 422', 'Bit depth': 'Bit depth is 8 / 10 when using ProRes 422' }));
});

test('allowed property by property but by no single profile: the nearest profile\'s requirements, named', () => {
  // 60 s is too long for Graphics (but fine for Promos); 110 Mbit too much for Promos (but fine for Graphics).
  const { gas, marks } = validationProject([probedVideo('x', { seconds: 60, mbit: 110 })]);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), failed({ Duration: 'Max duration is 0:00:30 when using ProRes 422 (Graphics)' }));
});

test('global and codec problems are reported together, several notes on one cell joined', () => {
  const { gas, marks } = validationProject([probedVideo('x', { fps: 25, scan: 'Interlaced TFF', mbit: 10 })]);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), failed({
    'Frame rate': 'Frame rate should be 50',
    Scan: 'Scan should be Progressive',
    Mbit: 'Min MBit is 50 when using ProRes 422',
  }));
});

test('videos not probed are left unmarked, and old marks are cleared when a video passes', () => {
  const { gas, marks, inbox } = validationProject([probedVideo('x', { status: 'Queued 2026-11-21 18:00', fps: 25 }), probedVideo('y', { fps: 25 })]);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), {});
  assert.equal(marks(2).Name, GREY);
  inbox.data[2][6] = 50; // fixed
  assert.deepEqual(plain(gas.validateVideoInbox_()), { approved: 1, failed: 0, unprobed: 1 });
  assert.deepEqual(marks(2), { Name: 'Approved for:\nGraphics\nPromos / Commercials' });
});

test('editing a profile re-validates (handleEdit); other edits do not', () => {
  const { gas, profileSheet, marks } = validationProject([probedVideo('x', { fps: 25 })]);
  gas.handleEdit({ range: gas.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Video Inbox').getRange(2, 2), value: 'x' });
  assert.deepEqual(marks(1), {}, 'not validated by an edit in Video Inbox');
  profileSheet.data[2][2] = '25/50'; // Graphics now takes 25 fps
  gas.handleEdit({ range: profileSheet.getRange(3, 3) });
  assert.deepEqual(marks(1), { Name: 'Approved for:\nGraphics' });
});

test('profile rules: lists with "/", Swedish decimals, case, empty cells allow anything; durations typed as durations', () => {
  const profiles = [['Any 29.97', '', '29,97 / 59,94', '', '', 'prores 422', '', '', '', '', '', '']];
  const { gas, marks, profileSheet } = validationProject([probedVideo('x', { fps: 29.97, mbit: 999, seconds: 9999 })], profiles);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), { Name: 'Approved for:\nAny 29.97' });
  // A max duration typed into the sheet is a duration value, shown as h:mm:ss.
  profileSheet.data[1][1] = 9000 / 86400;
  profileSheet.formats.set('1,1', '[h]:mm:ss');
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), failed({ Duration: 'Max duration should be 2:30:00' }), 'a single profile: every rule is global');
});

test('validation runs after Video Inbox is rebuilt, and on watcher runs without new files', () => {
  const { gas, tab } = inboxProject({
    sheets: [new FakeSheet('Video Profiles', [PROFILE_HEADERS, ...PROFILES])],
    video: [{ id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: LONG_AGO }],
  });
  gas.refreshInboxLists();
  const inbox = gas.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Video Inbox');
  // The job writes a result that fails (25 fps), as the Sheets API does: no edit event.
  inbox.getRange(2, 6, 1, 11).setValues([[20 / 86400, 25, 1920, 1080, 'ProRes 422', 80, 'Progressive', 10, 'No', 'stereo (2 ch)', 'Probed 2026-11-21 18:00']]);
  assert.equal(inbox.notes.size, 0);
  gas.checkForNewAssets(); // nothing new: validates
  assert.equal(inbox.notes.get('1,6'), 'Frame rate should be 50');
  assert.ok(tab('Video Inbox'));
});

test('runPostDeployTests: Video Profiles lists the profiles and what they check, and warns about unknown columns', () => {
  const gas = liveProject();
  gas.runPostDeployTests();
  assert.ok(gas.fakes.logs.some((l) => /^PASS +Video Profiles +5 profiles \(Graphics with alpha, Graphics, Promos \/ Commercials, Editorial short stories, Editorial\), checking Max duration, Frame rate, Width, Height, Codec, Min MBit, Max Mbit, Scan, Bit depth, Alpha, Audio$/.test(l)),
    gas.fakes.logs.join('\n'));
  const ss = gas.SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName('Video Profiles').data[0].push('Colour space');
  assert.match(gas.runPostDeployTests(), /1 warnings/);
  assert.ok(gas.fakes.logs.some((l) => /^WARN +Video Profiles .*columns that name no Video Inbox column \(ignored\): Colour space$/.test(l)));
});

test('a * in a rule stands for any text: "ProRes 422*" allows the whole ProRes 422 family', () => {
  const profiles = PROFILES.map((p) => p.slice());
  profiles[2][5] = 'ProRes 422*'; // Promos / Commercials
  const { gas, marks } = validationProject([
    probedVideo('lt', { codec: 'ProRes 422 LT', mbit: 81.2, seconds: 90 }),
    probedVideo('hq', { codec: 'prores 422 hq', mbit: 90, seconds: 60 }),
    probedVideo('x4', { codec: 'ProRes 4444 XQ', mbit: 90, seconds: 60 }), // not a 422
    probedVideo('lt2', { codec: 'ProRes 422 LT', mbit: 120, seconds: 90 }),
  ], profiles);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), { Name: 'Approved for:\nPromos / Commercials' });
  assert.deepEqual(marks(2), { Name: 'Approved for:\nPromos / Commercials' }, 'case does not matter');
  assert.deepEqual(marks(3), failed({ Codec: 'ProRes 4444 XQ is not an allowed codec.' }));
  assert.deepEqual(marks(4), failed({ Mbit: 'Max Mbit is 100 when using ProRes 422 LT' }), 'compared with the profiles whose pattern allows it');
  // Only * is special: other characters are taken as they are.
  assert.equal(gas.wildcardRegExp_('h.264 (main)*').test('h.264 (main) 4:2:0'), true);
  assert.equal(gas.wildcardRegExp_('h.264*').test('hx264 main'), false);
});

test('a * works in any value column: text (Audio, Scan) and numbers (Width)', () => {
  const profiles = [['Loose', '', 50, '19*', '', 'ProRes*', '', '', 'Interlaced*', '', '', '* (2 ch) / none']];
  const { gas, marks } = validationProject([
    probedVideo('ok', { scan: 'Interlaced TFF', audio: 'stereo (2 ch)' }),
    probedVideo('no', { w: 1280, scan: 'Progressive', audio: '5.1 (6 ch)' }),
  ], profiles);
  gas.validateVideoInbox_();
  assert.deepEqual(marks(1), { Name: 'Approved for:\nLoose' });
  assert.deepEqual(marks(2), failed({ Width: 'Width should be 19*', Scan: 'Scan should be Interlaced*', Audio: 'Audio should be * (2 ch) / none' }));
});

// --- The Action column: moving approved files, deleting the rest (actions.js) ---------------------------

// An inbox with the given videos and images, the profiles, and the videos' probe results written
// as the job writes them; validated. results: { videoId: probedVideo overrides, or a status text }.
function actionProject({ video = [], images = [], results = {} } = {}) {
  const project = inboxProject({ video, images, sheets: [new FakeSheet('Video Profiles', [PROFILE_HEADERS, ...PROFILES])] });
  const { gas } = project;
  gas.refreshInboxLists();
  const ss = gas.SpreadsheetApp.getActiveSpreadsheet();
  const inbox = ss.getSheetByName('Video Inbox');
  plain(inbox.data).forEach((row, i) => {
    const r = results[row[0]];
    if (i === 0 || r === undefined) return;
    if (typeof r === 'string') { inbox.getRange(i + 1, 16).setValue(r); return; }
    inbox.getRange(i + 1, 6, 1, 11).setValues([probedVideo(row[0], r).slice(5)]);
  });
  gas.validateVideoInbox_();
  const rowOf = (sheet, id) => plain(sheet.data).findIndex((r) => r[0] === id) + 1;
  const choose = (sheetName, id, value) => {
    const sheet = ss.getSheetByName(sheetName);
    const range = sheet.getRange(rowOf(sheet, id), 4);
    range.setValue(value);
    gas.handleEdit({ range, value });
  };
  const dropdown = (sheetName, id) => {
    const sheet = ss.getSheetByName(sheetName);
    const rule = sheet.validations.get(`${rowOf(sheet, id) - 1},3`);
    return rule ? plain(rule.values) : null;
  };
  const folder = (id) => gas.DriveApp.getFolderById(id);
  const sub = (id, name) => folder(id).getFoldersByName(name).next();
  const names = (f) => f.files.map((x) => x.getName());
  return Object.assign(project, { ss, inbox, choose, dropdown, folder, sub, names, rowOf });
}

const VIDEOS = [
  { id: 'v1', name: 'ok.mov', mimeType: 'video/quicktime', created: LONG_AGO },
  { id: 'v2', name: 'bad.mov', mimeType: 'video/quicktime', created: LONG_AGO },
  { id: 'v3', name: 'busy.mov', mimeType: 'video/quicktime', created: LONG_AGO },
  { id: 'v4', name: 'broken.mov', mimeType: 'video/quicktime', created: LONG_AGO },
  { id: 'n1', name: 'notes.txt', mimeType: 'text/plain', created: LONG_AGO },
];
const RESULTS = { v1: {}, v2: { fps: 25 }, v4: 'Error: ffprobe failed: Invalid data' };

test('the Action dropdowns: Video subfolders for approved videos, Delete for the rest, nothing while probing; Images subfolders for images', () => {
  const p = actionProject({ video: VIDEOS, images: [{ id: 'i1', name: 'a.png', mimeType: 'image/png', created: LONG_AGO }], results: RESULTS });
  assert.deepEqual(p.dropdown('Video Inbox', 'v1'), ['Bumpers', 'Commercials', 'Editorial'], 'approved: the subfolders, sorted');
  assert.deepEqual(p.dropdown('Video Inbox', 'v2'), ['Delete'], 'failed validation');
  assert.equal(p.dropdown('Video Inbox', 'v3'), null, 'still being probed (Queued)');
  assert.deepEqual(p.dropdown('Video Inbox', 'v4'), ['Delete'], 'could not be probed');
  assert.deepEqual(p.dropdown('Video Inbox', 'n1'), ['Delete'], 'not a video');
  assert.deepEqual(p.dropdown('Image Inbox', 'i1'), ['Episodes', 'Genre', 'Series']);
});

test('a new subfolder shows up in the dropdowns without any change', () => {
  const p = actionProject({ video: VIDEOS, results: RESULTS });
  p.folder('PRODVIDEO').folders.push(fakeFolder({ name: 'Trailers' }));
  p.gas.validateVideoInbox_();
  assert.deepEqual(p.dropdown('Video Inbox', 'v1'), ['Bumpers', 'Commercials', 'Editorial', 'Trailers']);
});

test('choosing a subfolder moves the approved video there; the row goes, and Video Assets lists it', () => {
  const p = actionProject({ video: VIDEOS, results: RESULTS });
  p.choose('Video Inbox', 'v1', 'Commercials');
  assert.deepEqual(p.names(p.sub('PRODVIDEO', 'Commercials')), ['ok.mov']);
  assert.equal(p.rowOf(p.inbox, 'v1'), 0, 'gone from Video Inbox');
  assert.equal(p.ss.getSheetByName('Video Assets').data[1][1], 'ok.mov', 'Video Assets was refreshed');
  assert.ok(p.gas.fakes.logs.includes('"ok.mov" är flyttad till Commercials'));
});

test('Delete trashes the file and the row goes; images move to their subfolders', () => {
  const p = actionProject({ video: VIDEOS, images: [{ id: 'i1', name: 'Handboll.png', mimeType: 'image/png', created: LONG_AGO }], results: RESULTS });
  const bad = p.gas.DriveApp.getFileById('v2');
  p.choose('Video Inbox', 'v2', 'Delete');
  assert.equal(bad.isTrashed(), true);
  assert.equal(p.rowOf(p.inbox, 'v2'), 0);
  p.choose('Image Inbox', 'i1', 'Genre');
  assert.deepEqual(p.names(p.sub('PRODIMAGES', 'Genre')), ['Handboll.png']);
  assert.equal(p.ss.getSheetByName('Image Inbox').data.slice(1).filter((r) => r[0]).length, 0);
});

test('a name already taken by another file in the folder gets -001, -002 … before its extension', () => {
  // Three uploads named ok.mov, each with content of its own, none the same as the one in Editorial.
  const p = actionProject({ video: [...VIDEOS.map((v) => (v.id === 'v1' ? Object.assign({}, v, { md5: 'first' }) : v)),
    { id: 'v5', name: 'ok.mov', mimeType: 'video/quicktime', created: LONG_AGO, md5: 'second' },
    { id: 'v6', name: 'ok.mov', mimeType: 'video/quicktime', created: LONG_AGO, md5: 'third' }], results: Object.assign({ v5: {}, v6: {} }, RESULTS) });
  const editorial = p.sub('PRODVIDEO', 'Editorial');
  editorial.files.push(fakeFile({ id: 'other', name: 'ok.mov', mimeType: 'video/quicktime', md5: 'other content' }, editorial));
  p.choose('Video Inbox', 'v1', 'Editorial');
  p.choose('Video Inbox', 'v5', 'Editorial');
  p.choose('Video Inbox', 'v6', 'Editorial');
  assert.deepEqual(p.names(editorial), ['ok.mov', 'ok-001.mov', 'ok-002.mov', 'ok-003.mov']);
  assert.ok(p.gas.fakes.logs.includes('"ok.mov" är flyttad till Editorial som "ok-001.mov"'));
  assert.equal(p.ss.getSheetByName('Video Assets').data.slice(1).map((r) => r[1]).sort().join(), 'ok-001.mov,ok-002.mov,ok-003.mov,ok.mov');
});

test('a re-upload (same name, same content) is not moved: the copy in production stays, the inbox copy goes to the trash', () => {
  const p = actionProject({ video: [
    { id: 'v1', name: 'ok.mov', mimeType: 'video/quicktime', created: LONG_AGO, md5: 'abc' },
    { id: 'v2', name: 'new.mov', mimeType: 'video/quicktime', created: LONG_AGO, md5: 'abc' }, // same content, other name
  ], results: { v1: {}, v2: {} } });
  const editorial = p.sub('PRODVIDEO', 'Editorial');
  const original = fakeFile({ id: 'prod', name: 'ok.mov', mimeType: 'video/quicktime', md5: 'abc' }, editorial);
  editorial.files.push(original);
  const upload = p.gas.DriveApp.getFileById('v1');
  p.choose('Video Inbox', 'v1', 'Editorial');
  assert.equal(upload.isTrashed(), true, 'the inbox copy is in the trash');
  assert.deepEqual(editorial.files.map((f) => f.getId()), ['prod'], 'production is untouched: the same file, same ID');
  assert.equal(p.rowOf(p.inbox, 'v1'), 0, 'its row is gone');
  assert.ok(p.gas.fakes.logs.includes('"ok.mov" fanns redan i Editorial med samma innehåll: inkorgens kopia ligger nu i papperskorgen'));

  // The same content under another name is not a re-upload of that file: moved as usual.
  p.choose('Video Inbox', 'v2', 'Editorial');
  assert.deepEqual(p.names(editorial), ['ok.mov', 'new.mov']);
});

test('files without a checksum are never taken as identical', () => {
  const p = actionProject({ video: [{ id: 'v1', name: 'ok.mov', mimeType: 'video/quicktime', created: LONG_AGO }], results: { v1: {} } });
  const editorial = p.sub('PRODVIDEO', 'Editorial');
  editorial.files.push(fakeFile({ id: 'prod', name: 'ok.mov', mimeType: 'video/quicktime' }, editorial));
  p.choose('Video Inbox', 'v1', 'Editorial');
  assert.deepEqual(p.names(editorial), ['ok.mov', 'ok-001.mov']);
});

test('freeFileName_: the file\'s own name when only it has it; a name without extension; a dot at the start', () => {
  const p = actionProject();
  const folder = (files) => ({ getName: () => 'F', getFilesByName: (n) => { const l = files.filter((f) => f.getName() === n); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; } });
  const f = (id, name) => ({ getId: () => id, getName: () => name });
  assert.equal(p.gas.freeFileName_(folder([f('me', 'a.mov')]), f('me', 'a.mov')), 'a.mov', 'the same Drive ID: kept as it is');
  assert.equal(p.gas.freeFileName_(folder([f('x', 'README')]), f('me', 'README')), 'README-001');
  assert.equal(p.gas.freeFileName_(folder([f('x', '.hidden')]), f('me', '.hidden')), '.hidden-001');
  assert.equal(p.gas.freeFileName_(folder([f('x', 'a.b.mov'), f('y', 'a.b-001.mov')]), f('me', 'a.b.mov')), 'a.b-002.mov');
});

test('a video that is not approved is not moved, even if a folder is pasted in; the choice is undone', () => {
  const p = actionProject({ video: VIDEOS, results: RESULTS });
  p.choose('Video Inbox', 'v2', 'Commercials'); // pasted past the dropdown
  assert.equal(p.gas.DriveApp.getFileById('v2').getParent().getName(), 'Video');
  assert.equal(p.inbox.data[p.rowOf(p.inbox, 'v2') - 1][3], '', 'the choice is cleared');
  assert.ok(p.gas.fakes.logs.some((l) => l.includes('"bad.mov" är inte godkänd för någon profil')));
});

test('handleEdit ignores other columns, several cells at once, and an emptied cell', () => {
  const p = actionProject({ video: VIDEOS, results: RESULTS });
  const before = p.names(p.folder('INBOX').getFoldersByName('Video').next());
  const ok = p.rowOf(p.inbox, 'v1'); // the approved video: any of these would otherwise move it
  p.gas.handleEdit({ range: p.inbox.getRange(ok, 2), value: 'Commercials' });
  p.gas.handleEdit({ range: p.inbox.getRange(ok, 4, 2, 1), value: 'Commercials' });
  p.gas.handleEdit({ range: p.inbox.getRange(ok, 4), value: '' });
  assert.deepEqual(p.names(p.folder('INBOX').getFoldersByName('Video').next()), before);
  assert.deepEqual(p.names(p.sub('PRODVIDEO', 'Commercials')), []);
});

test('installEditTrigger installs handleEdit for the spreadsheet, once', () => {
  const p = actionProject();
  p.gas.installEditTrigger();
  p.gas.installEditTrigger();
  assert.deepEqual(p.gas.fakes.triggers.map((t) => [t.getHandlerFunction(), t.getTriggerSourceId()]), [['handleEdit', 'spreadsheet-id']]);
});

test('runPostDeployTests: the production folders and handleEdit', () => {
  const gas = liveProject();
  gas.runPostDeployTests();
  assert.ok(gas.fakes.logs.some((l) => /^PASS +Drive: production folders \(Action dropdowns\) +Video Inbox → "Video": Commercials, Editorial; Image Inbox → "Images": Episodes, Genre, Series$/.test(l)), gas.fakes.logs.join('\n'));
  gas.fakes.triggers.splice(gas.fakes.triggers.findIndex((t) => t.getHandlerFunction() === 'handleEdit'), 1);
  gas.runPostDeployTests();
  assert.ok(gas.fakes.logs.some((l) => /^WARN +Triggers +handleEdit is not installed by you \(run installEditTrigger\)/.test(l)));
});

test('when the last row goes, nothing of it stays: no note, colour or dropdown', () => {
  const p = actionProject({ video: [{ id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: LONG_AGO },
    { id: 'v2', name: 'z.mov', mimeType: 'video/quicktime', created: LONG_AGO }], results: { v1: {}, v2: { fps: 25 } } });
  const last = p.rowOf(p.inbox, 'v2') - 1; // z.mov, failed: grey, red, noted, Delete
  assert.equal(last, 2);
  assert.ok(p.inbox.notes.has(`${last},6`) && p.inbox.cellBackgrounds.has(`${last},0`) && p.inbox.validations.has(`${last},3`));
  p.choose('Video Inbox', 'v2', 'Delete');
  const left = (map) => [...map.keys()].filter((k) => k.startsWith(`${last},`));
  assert.deepEqual([left(p.inbox.notes), left(p.inbox.cellBackgrounds), left(p.inbox.fontColors), left(p.inbox.validations)], [[], [], [], []]);
  assert.equal(p.inbox.notes.get('1,1'), 'Approved for:\nGraphics\nPromos / Commercials', 'the row left keeps its marks');
});

test('a choice says straight away what is happening, for 5 seconds', () => {
  const p = actionProject({ video: VIDEOS, results: RESULTS });
  const toasts = [];
  p.ss.toast = (msg, title, seconds) => toasts.push([msg, title, seconds]);
  p.choose('Video Inbox', 'v1', 'Commercials');
  p.choose('Video Inbox', 'v2', 'Delete');
  assert.deepEqual(toasts[0], ['Moving "ok.mov" to Commercials…', 'Video Inbox', 5]);
  assert.ok(toasts.some((t) => t[0] === 'Deleting "bad.mov"…'));
  assert.ok(toasts.indexOf(toasts.find((t) => t[0].startsWith('"ok.mov" är flyttad'))) > 0, 'then what was done');
});

// --- Probe results stored on the files (Drive properties) ------------------------------------------------

test('the Asset Manager and the probe job store results under the same Drive property keys', () => {
  const job = require('../../../services/media-probe/src/properties');
  const gas = load(FILES);
  assert.deepEqual(plain(gas.get('PROBE_PROPERTY_KEYS')), job.PROPERTY_KEYS);
  assert.equal(gas.get('PROBED_PROPERTY'), job.PROBED_KEY);
});

test('moving a video stores its probe results and approved profiles on the file; Video Assets shows the stored duration', () => {
  const p = actionProject({ video: VIDEOS, results: RESULTS });
  const file = p.gas.DriveApp.getFileById('v1');
  p.choose('Video Inbox', 'v1', 'Commercials');
  assert.deepEqual(file.properties, {
    'sesport.duration': '20', 'sesport.frameRate': '50', 'sesport.width': '1920', 'sesport.height': '1080',
    'sesport.codec': 'ProRes 422', 'sesport.mbit': '80', 'sesport.scan': 'Progressive', 'sesport.bitDepth': '10',
    'sesport.alpha': 'No', 'sesport.audio': 'stereo (2 ch)', 'sesport.probed': '2026-11-21 18:00',
    'sesport.approvedFor': 'Graphics, Promos / Commercials',
  });
  const assets = p.ss.getSheetByName('Video Assets');
  assert.equal(assets.data[1][5], 20 / 86400, 'the probed 20 s, not Drive\'s 2 s');
});

test('Video Assets reads the stored results with one Drive request per folder, and falls back to Drive\'s duration', () => {
  const p = actionProject();
  const editorial = p.sub('PRODVIDEO', 'Editorial');
  editorial.files.push(fakeFile({ id: 'vs', name: 'stored.mov', mimeType: 'video/quicktime', properties: { 'sesport.duration': '95' } }, editorial));
  editorial.files.push(fakeFile({ id: 'vd', name: 'drive.mov', mimeType: 'video/quicktime' }, editorial));
  const drive = p.gas.Drive;
  const before = drive.lists;
  p.gas.listDriveVideos();
  assert.equal(drive.lists - before, 4, 'Video and its three subfolders: one request each');
  const rows = plain(p.ss.getSheetByName('Video Assets').data.slice(1));
  assert.deepEqual(rows.map((r) => [r[1], r[5]]), [['drive.mov', 2 / 86400], ['stored.mov', 95 / 86400]]);
});

test('a move whose results cannot be stored does not happen; the choice is undone', () => {
  const p = actionProject({ video: VIDEOS, results: RESULTS });
  p.gas.Drive.Files.update = () => { throw new Error('Insufficient permissions'); };
  p.choose('Video Inbox', 'v1', 'Commercials');
  assert.equal(p.gas.DriveApp.getFileById('v1').getParent().getName(), 'Video', 'still in the inbox');
  assert.ok(p.gas.fakes.logs.some((l) => l.includes('Commercials failed: Insufficient permissions')));
});

test('Analysera videor utan metadata sends only the videos without stored results, to Drive properties', () => {
  const p = actionProject();
  const bumpers = p.sub('PRODVIDEO', 'Bumpers');
  bumpers.files.push(fakeFile({ id: 'has', name: 'has.mov', mimeType: 'video/quicktime', properties: { 'sesport.duration': '10' } }, bumpers));
  bumpers.files.push(fakeFile({ id: 'needs', name: 'needs.mov', mimeType: 'video/quicktime' }, bumpers));
  bumpers.files.push(fakeFile({ id: 'mxf', name: 'needs.mxf', mimeType: 'application/mxf' }, bumpers));
  bumpers.files.push(fakeFile({ id: 'txt', name: 'notes.txt', mimeType: 'text/plain' }, bumpers));
  assert.deepEqual(plain(p.gas.probeAssetsWithoutMetadata()), { started: 2, left: 0 });
  const [req] = runRequests(p.gas);
  assert.match(req.url, /\/jobs\/media-probe:run$/);
  const overrides = JSON.parse(req.options.payload).overrides;
  assert.deepEqual(overrides.containerOverrides[0].env.slice(0, 2), [{ name: 'FILE_IDS', value: 'needs,mxf' }, { name: 'TARGET', value: 'drive' }]);
  assert.equal(overrides.taskCount, 1);
});

test('Analysera videor utan metadata: 500 at a time, in up to 4 tasks; nothing to do, no job', () => {
  const p = actionProject();
  const editorial = p.sub('PRODVIDEO', 'Editorial');
  for (let i = 0; i < 501; i++) editorial.files.push(fakeFile({ id: 'f' + i, name: `c${i}.mov`, mimeType: 'video/quicktime' }, editorial));
  assert.deepEqual(plain(p.gas.probeAssetsWithoutMetadata()), { started: 500, left: 1 });
  const overrides = JSON.parse(runRequests(p.gas)[0].options.payload).overrides;
  assert.equal(overrides.containerOverrides[0].env[0].value.split(',').length, 500);
  assert.equal(overrides.taskCount, 4);

  const done = actionProject();
  assert.deepEqual(plain(done.gas.probeAssetsWithoutMetadata()), { started: 0, left: 0 });
  assert.equal(runRequests(done.gas).length, 0);
});

test('Video Assets: Path is dark grey for a video without stored results, black (no colour) for one with them', () => {
  const p = actionProject();
  const editorial = p.sub('PRODVIDEO', 'Editorial');
  editorial.files.push(fakeFile({ id: 'a', name: 'a.mov', mimeType: 'video/quicktime', properties: { 'sesport.probed': '2026-11-21T20:59:11Z', 'sesport.duration': '90' } }, editorial));
  editorial.files.push(fakeFile({ id: 'b', name: 'b.mov', mimeType: 'video/quicktime' }, editorial));
  editorial.files.push(fakeFile({ id: 'c', name: 'c.mov', mimeType: 'video/quicktime' }, editorial));
  p.gas.listDriveVideos();
  const assets = p.ss.getSheetByName('Video Assets');
  assert.deepEqual([1, 2, 3].map((r) => assets.fontColors.get(`${r},4`) || null), [null, '#666666', '#666666']);
  // c gets its results, and b goes: c turns black, and b's old row keeps no colour.
  const c = p.gas.DriveApp.getFileById('c');
  c.properties['sesport.probed'] = '2026-11-22T08:00:00Z';
  p.gas.DriveApp.getFileById('b').setTrashed(true);
  p.gas.listDriveVideos();
  assert.deepEqual([1, 2, 3].map((r) => assets.fontColors.get(`${r},4`) || null), [null, null, null]);
});
