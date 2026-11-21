// Drive and Sheets for the probe job, called over their REST APIs as the job's service account.
// Without a client library: the token comes from the Cloud Run metadata server, and requests use
// fetch. (The same approach as services/dmc-transfer/src/google.js.)
'use strict';

const DRIVE = 'https://www.googleapis.com/drive/v3/files/';
const SHEETS = 'https://sheets.googleapis.com/v4/spreadsheets/';
// Drive with write access: the job stores its results on files as properties (properties.js). What it
// may touch is still decided by what is shared with its service account.
const SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'];
const METADATA_TOKEN = 'http://metadata.google.internal/computeMetadata/v1/instance/service-accounts/default/token';

/** An access token for the job's service account (Cloud Run hands these out to the job). */
async function metadataToken(fetchFn) {
  const res = await fetchFn(METADATA_TOKEN + '?scopes=' + encodeURIComponent(SCOPES.join(',')),
    { headers: { 'Metadata-Flavor': 'Google' } });
  if (!res.ok) throw new Error(`No access token from the metadata server (${res.status})`);
  return (await res.json()).access_token;
}

/** 1 -> "A", 27 -> "AA" */
function columnLetter(col) {
  let s = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** The URL ffprobe reads a Drive file from (with the token in an Authorization header). */
function driveDownloadUrl(id) {
  return DRIVE + encodeURIComponent(id) + '?alt=media&supportsAllDrives=true';
}

/**
 * options.fetch and options.token (an async function returning an access token) replace the real
 * ones in tests.
 */
function googleApi(options = {}) {
  const fetchFn = options.fetch || fetch;
  const token = options.token || (() => metadataToken(fetchFn));

  async function call(url, init = {}) {
    const headers = Object.assign({}, init.headers, { Authorization: 'Bearer ' + (await token()) });
    const res = await fetchFn(url, Object.assign({}, init, { headers }));
    if (!res.ok) {
      const body = (await res.text()).slice(0, 300);
      throw new Error(`${init.method || 'GET'} ${url.split('?')[0]} failed (${res.status}): ${body}`);
    }
    return res;
  }

  /**
   * A write(fileId, values) function: values maps column headers to what goes in them, written in
   * one request to the row whose column A holds fileId. Headers and rows are looked up on every
   * write, so the row is found even after the tab has been rebuilt and re-sorted.
   */
  function sheetRow({ spreadsheetId, sheetName }) {
    const base = SHEETS + encodeURIComponent(spreadsheetId) + '/values';
    const range = (r) => `'${sheetName.replace(/'/g, "''")}'!${r}`;
    return async function write(fileId, values) {
      const res = await call(`${base}:batchGet?ranges=${encodeURIComponent(range('1:1'))}&ranges=${encodeURIComponent(range('A:A'))}`);
      const [headerRows, ids] = (await res.json()).valueRanges.map((v) => v.values || []);
      const headers = (headerRows[0] || []).map((h) => String(h).trim().toLowerCase());
      const row = ids.findIndex((v, i) => i > 0 && v[0] === fileId);
      if (row < 0) throw new Error(`${fileId} is not in ${sheetName}`);
      const data = Object.keys(values).map((name) => {
        const col = headers.indexOf(name.toLowerCase());
        if (col < 0) throw new Error(`No "${name}" column in ${sheetName}`);
        return { range: range(columnLetter(col + 1) + (row + 1)), values: [[values[name]]] };
      });
      await call(`${base}:batchUpdate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ valueInputOption: 'RAW', data }),
      });
    };
  }

  /** Stores properties (key -> text) on a Drive file, alongside any it has. */
  async function setFileProperties(fileId, properties) {
    await call(DRIVE + encodeURIComponent(fileId) + '?supportsAllDrives=true&fields=id', {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ properties }),
    });
  }

  return { token, sheetRow, setFileProperties };
}

module.exports = { googleApi, columnLetter, driveDownloadUrl };
