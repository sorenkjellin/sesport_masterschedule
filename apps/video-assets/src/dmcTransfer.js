/**
 * DMC, the recipient behind Playout: clips go to DMC's FTP server (FTPS, explicit TLS on port 21)
 * through the Cloud Run job dmc-transfer (services/dmc-transfer in the repo), which writes each
 * clip's progress to "Recipient status", as "Playout: …". Users choose "Send to Playout" in Video Assets' Action
 * column; the sending itself is generic (recipients.js). A file with the same name at DMC is
 * overwritten: all clips go to the same folder there.
 *
 * Script Properties:
 *   DMC_GCP_PROJECT  ID of the Google Cloud project the job is in (the probe job's too)
 *   DMC_REGION       the job's region (optional, default europe-north1)
 *   DMC_JOB          the job's name (optional, default dmc-transfer)
 *
 * Setting up the job, its password and who may send: docs/dmc-transfer.md in the repo.
 */

/** Playout, delivered by DMC (a recipient description: see recipients.js). */
function dmcRecipient_() {
  return {
    choice: 'Send to Playout',
    label: 'Playout',
    formerHeaders: ['DMC-status'], // the column before Recipient status: only Playout's statuses
    job: dmcJob_,
    checks: {
      requiredEnv: ['FTP_HOST', 'FTP_USER', 'FTP_PASSWORD'],
      secretEnv: 'FTP_PASSWORD',
      deployScript: 'scripts/deploy-dmc-transfer.sh',
      describe: function (env) { return 'sends to ' + env.FTP_HOST + (env.FTP_DIR ? '/' + env.FTP_DIR : ''); },
    },
  };
}

/** The Cloud Run job that delivers to DMC (the `job` argument of the Cloud Run API in lib.js). */
function dmcJob_() {
  return {
    project: requiredProperty_('DMC_GCP_PROJECT'),
    region: optionalProperty_('DMC_REGION', 'europe-north1'),
    name: optionalProperty_('DMC_JOB', 'dmc-transfer'),
  };
}
