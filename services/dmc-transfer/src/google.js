// Drive and Sheets, called over their REST APIs as the job's service account. Without a client
// library: the token comes from the Cloud Run metadata server, and requests use fetch.
'use strict';
const { Readable } = require('node:stream');

const DRIVE = 'https://www.googleapis.com/drive/v3/files/';
const SHEETS = 'https://sheets.googleapis.com/v4/spreadsheets/';
const SCOPES = ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets'];
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

/**
 * options.fetch and options.token (an async function returning an access token) replace the real
 * ones in tests. A fresh token is asked for on every request: transfers outlast a token's hour, and
 * the metadata server caches them.
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

  const drive = {
    /** { name, size } of a Drive file (also in shared drives). */
    async metadata(id) {
      const res = await call(DRIVE + encodeURIComponent(id) + '?fields=name,size&supportsAllDrives=true');
      const json = await res.json();
      return { name: json.name, size: json.size == null ? NaN : Number(json.size) };
    },
    /** The file's content as a stream, so a 5 GB clip passes through without being stored. */
    async download(id) {
      const res = await call(DRIVE + encodeURIComponent(id) + '?alt=media&supportsAllDrives=true');
      return Readable.fromWeb(res.body);
    },
  };

  /**
   * A status(fileId, text) function writing text to the `header` column of the row whose column A
   * holds fileId. The row is looked up on every write, so it is found even after the list has been
   * re-sorted during a transfer.
   */
  function sheetStatus({ spreadsheetId, sheetName, header }) {
    const base = SHEETS + encodeURIComponent(spreadsheetId) + '/values';
    const range = (r) => encodeURIComponent(`'${sheetName.replace(/'/g, "''")}'!${r}`);
    return async function status(fileId, text) {
      const res = await call(`${base}:batchGet?ranges=${range('1:1')}&ranges=${range('A:A')}`);
      const [headers, ids] = (await res.json()).valueRanges.map((v) => v.values || []);
      const col = (headers[0] || []).findIndex((h) => String(h).trim().toLowerCase() === header.toLowerCase());
      if (col < 0) throw new Error(`No "${header}" column in ${sheetName}`);
      const row = ids.findIndex((v, i) => i > 0 && v[0] === fileId);
      if (row < 0) throw new Error(`${fileId} is not in ${sheetName}`);
      await call(`${base}/${range(columnLetter(col + 1) + (row + 1))}?valueInputOption=RAW`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ values: [[text]] }),
      });
    };
  }

  return { drive, sheetStatus };
}

module.exports = { googleApi, columnLetter };
