// Unit tests for the Video Assets app. Run with: npm test
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load, scriptsIn, plain, FakeSheet, fakeSpreadsheet, fakeFolder } = require('../../../tests/harness');

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
  VID: sharedWith('reader'),
  INBOXVIDEO: sharedWith('reader', PROBE_SA),
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
  const ss = fakeSpreadsheet([new FakeSheet('Video Assets', [headers, ['a', 'a.mp4']])], 'spreadsheet-id', 'Video Assets');
  const gas = load(FILES, {
    spreadsheet: ss,
    properties: {
      VIDEO_FOLDER_ID: 'VID', DMC_GCP_PROJECT: 'sesport',
      WATCH_FOLDER_ID: 'WATCHED', WATCH_RECIPIENTS: 'watch@example.com', WATCH_RUN_AS: 'User@example.com',
      ...properties,
    },
    folders: { VID: fakeFolder({ id: 'VID', name: 'Video' }), WATCHED: fakeFolder({ id: 'WATCHED', name: 'Watched', folders: inboxFolders }) },
    fetch: cloudFetch(Object.assign({ job: dmcJob }, cloud)),
    globals: { Drive: fakeDriveSharing(sharing) },
  });
  gas.installWatcherTrigger();
  return gas;
}

test('runPostDeployTests passes on a correctly set up project and logs every check', () => {
  const gas = liveProject();
  const summary = gas.runPostDeployTests();
  const failed = gas.fakes.logs.filter((l) => !l.startsWith('PASS') && /^(WARN|FAIL)/.test(l));
  assert.equal(summary, '9 passed, 0 warnings, 0 failed', failed.join('\n'));
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
  assert.equal(gas.runPostDeployTests(), '8 passed, 1 warnings, 0 failed');
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
  assert.deepEqual(gas.fakes.triggers.map((t) => t.getHandlerFunction()), ['checkForNewAssets']);
  assert.ok(gas.fakes.logs.includes('checkForNewAssets now runs every 10 minutes as user@example.com'));
});

test('runPostDeployTests: the trigger check follows WATCH_RUN_AS', () => {
  const other = liveProject({ properties: { WATCH_RUN_AS: 'owner@example.com' } });
  other.fakes.triggers.splice(0); // the watcher's trigger is the owner's
  assert.equal(other.runPostDeployTests(), '9 passed, 0 warnings, 0 failed');
  assert.ok(other.fakes.logs.some((l) => l.includes('not checkForNewAssets (WATCH_RUN_AS: the watcher runs as owner@example.com)')));

  const mine = liveProject();
  mine.runPostDeployTests();
  assert.ok(mine.fakes.logs.some((l) => /^PASS +Triggers +checkForNewAssets \(WATCH_RUN_AS: the watcher runs as you, User@example.com\)$/.test(l)));

  const missing = liveProject();
  missing.fakes.triggers.splice(0);
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
    /^PASS +Script Properties +6 in all: the 4 required; optional WATCH_RUN_AS; kept by the script LAST_CHECK_TIME$/);
  assert.ok(!gas.fakes.logs.join('\n').includes('2026-11-20T10:00:00Z'), 'no values in the log');

  const old = liveProject({ properties: { DRIVE_FOLDER_ID: 'x' } });
  assert.match(old.runPostDeployTests(), /1 warnings/);
  assert.ok(old.fakes.logs.some((l) => /^WARN +Script Properties .*; not used by the code: DRIVE_FOLDER_ID$/.test(l)));
});

// --- Inbox tabs ------------------------------------------------------------------------------------

const PROBE_COLUMNS = ['Frame rate', 'Width', 'Height', 'Codec', 'Mbit', 'Scan', 'Bit depth', 'Alpha', 'Audio', 'Probe status'];
const MINUTE_AGO = Date.now() - 60000;
const LONG_AGO = Date.now() - 30 * 86400000;

// The Inbox (WATCH_FOLDER_ID 'INBOX') with Video and Images subfolders, files given per folder.
function inboxProject({ video = [], images = [], root = [], subfolders = ['Video', 'Images'], sheets = [], runCode = 200 } = {}) {
  const content = { Video: video, Images: images };
  const inbox = fakeFolder({ id: 'INBOX', name: 'Inbox', files: root,
    folders: subfolders.map((name) => ({ name, files: content[name].filter((f) => !f.sub),
      folders: content[name].some((f) => f.sub) ? [{ name: 'Matcher', files: content[name].filter((f) => f.sub) }] : [] })) });
  const ss = fakeSpreadsheet(sheets, 'spreadsheet-id', 'seSPORT AssetManager');
  const gas = load(FILES, {
    spreadsheet: ss,
    properties: { WATCH_FOLDER_ID: 'INBOX', WATCH_RECIPIENTS: 'a@example.com', DMC_GCP_PROJECT: 'sesport' },
    folders: { INBOX: inbox },
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
  assert.deepEqual(rows[0], ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Duration', ...PROBE_COLUMNS]);
  assert.deepEqual(rows.slice(1).map((r) => [r[0], r[1], r[4]]), [
    ['v1', 'a.mp4', '0:00:02:00'],
    ['v2', 'b.mov', '0:00:02:00'],
    ['v3', 'c.mp4', '0:00:02:00'],
    ['n1', 'notes.txt', ''], // every file is listed; only videos get a duration
  ]);
  assert.equal(rows[2][3], '2048');
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
    ['Drive file ID', 'Name', 4], ['i1', 'a.jpg', 4], ['i2', 'b.png', 4],
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
  assert.ok(tab('Image Inbox').every((r) => !r[4]), 'nothing is left in the old Path column, header included');
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
  assert.equal(gas.runPostDeployTests(), '8 passed, 1 warnings, 0 failed');
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
const O = 14; // "Probe status", column O

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
    [['v1', 'Queued T'], ['v2', 'Queued T'], ['n1', '']], 'the text file is not probed');
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
  assert.deepEqual(rows[0], ['Drive file ID', 'Name', 'Drive modified', 'Size', 'Duration', ...PROBE_COLUMNS]);
  assert.deepEqual(rows[2].slice(0, 1).concat(rows[2].slice(5)),
    ['v9', '25', 1920, 1080, 'ProRes 422 HQ', '', 'Interlaced TFF', 10, 'No', '8 × mono (8 ch)', 'Probed 2026-11-21 17:02'],
    'z.mov moved down a row with its results, each under its header; no Mbit until probed again');
  assert.match(rows[1][O], /^Queued /);
  assert.equal(rows[1][8], '', 'a.mov has no results yet');
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
  assert.ok(gas.fakes.logs.some((l) => l.endsWith(`you may start the job; ${PROBE_SA} can edit this spreadsheet and read Inbox/"Video"`)), gas.fakes.logs.join('\n'));
});

test('runPostDeployTests: the probe checks fail when Inbox/Video is not shared with the probe account, or the job writes elsewhere', () => {
  const unshared = liveProject({ sharing: Object.assign({}, SHARED, { INBOXVIDEO: sharedWith('reader') }) });
  assert.throws(() => unshared.runPostDeployTests(), /Post-deploy tests failed: Media probe: permissions$/);
  assert.ok(unshared.fakes.logs.some((l) => l.includes(`"Video" is not shared with ${PROBE_SA} as Viewer (for media-probe)`)));

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
  assert.deepEqual(rows.map((r) => [r[0], r[4]]), [
    ['v1', '0:00:14:12'], // the probe's, not Drive's 0:00:02:00
    ['m2', ''],           // nothing from Drive for MXF: the probe fills it in
    ['v3', '0:00:02:00'], // Drive's, until probed
  ]);
  assert.equal(JSON.parse(runRequests(gas)[0].options.payload).overrides.containerOverrides[0].env[0].value, 'm2,v3');
});

// --- Names link to their files ---------------------------------------------------------------------

const DRIVE_URL = (id) => `https://drive.google.com/file/d/${id}/view`;

test('the names in Video Inbox and Image Inbox link to their files in Drive, and stay plain text', () => {
  const { gas, tab } = inboxProject({
    video: [{ id: 'v1', name: 'a.mov', mimeType: 'video/quicktime', created: LONG_AGO }],
    images: [{ id: 'i1', name: 'b.png', mimeType: 'image/png', created: LONG_AGO }, { id: 'i2', name: 'c.png', mimeType: 'image/png', created: LONG_AGO }],
  });
  gas.refreshInboxLists();
  const ss = gas.SpreadsheetApp.getActiveSpreadsheet();
  assert.deepEqual([...ss.getSheetByName('Video Inbox').links], [['1,1', DRIVE_URL('v1')]]);
  assert.deepEqual([...ss.getSheetByName('Image Inbox').links], [['1,1', DRIVE_URL('i1')], ['2,1', DRIVE_URL('i2')]]);
  assert.equal(tab('Image Inbox')[2][1], 'c.png', 'the value is the name itself');
});

test('a link goes with its row: rows of files that are gone lose theirs', () => {
  const { gas } = inboxProject({ images: [{ id: 'i1', name: 'b.png', mimeType: 'image/png', created: LONG_AGO }, { id: 'i2', name: 'c.png', mimeType: 'image/png', created: LONG_AGO }] });
  gas.refreshInboxLists();
  const sheet = gas.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Image Inbox');
  // i1 is deleted from Drive.
  gas.DriveApp.getFolderById('INBOX').getFoldersByName('Images').next().getFiles = (() => {
    let done = false;
    return () => ({ hasNext: () => !done, next: () => { done = true; return { getId: () => 'i2', getName: () => 'c.png', getMimeType: () => 'image/png', getLastUpdated: () => new Date(0), getSize: () => 1 }; } });
  })();
  gas.refreshInboxLists();
  assert.deepEqual([...sheet.links], [['1,1', DRIVE_URL('i2')]]);
});

test('the names in Video Assets link to their files, and sending to DMC still reads the plain name', () => {
  const { gas, sheet } = videoProject({ videos: [{ id: 'a', name: 'a.mp4' }], rows: [] });
  gas.listDriveVideos();
  assert.deepEqual([...sheet.links], [['1,1', DRIVE_URL('a')]]);
  assert.equal(plain(gas.pickDmcClips_(plain(sheet.data.slice(1)).map((r) => (r[6] = true, r)))).clips[0].name, 'a.mp4');
});
