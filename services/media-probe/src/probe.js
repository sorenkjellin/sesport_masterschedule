// What a video file is: frame rate, size, codec, scan (progressive or interlaced), bit depth,
// alpha and audio layout, from ffprobe (and, when the file does not say how it is scanned, from
// ffmpeg's idet filter, which measures it on decoded frames).
//
// ffprobe reads the file straight from Drive over HTTP, with range requests, so only the parts it
// needs are fetched. The commands run through `run`, which the tests replace.
'use strict';
const { execFile } = require('node:child_process');

// The columns this fills in, in "Video Inbox" order (the Asset Manager's PROBE_HEADERS).
const RESULT_HEADERS = ['Frame rate', 'Width', 'Height', 'Codec', 'Scan', 'Bit depth', 'Alpha', 'Audio'];

/** The ffprobe arguments that print streams, the container, and the first video frame as JSON. */
function ffprobeArgs(url, token) {
  return ['-v', 'error', '-rw_timeout', '60000000',
    '-headers', `Authorization: Bearer ${token}\r\n`,
    '-print_format', 'json', '-show_streams', '-show_format',
    '-show_entries', 'frame=media_type,stream_index,interlaced_frame,top_field_first',
    '-read_intervals', '%+#1', url];
}

/** The ffmpeg arguments that measure interlacing on the first `frames` frames (idet). */
function idetArgs(url, token, frames = 300) {
  return ['-hide_banner', '-nostats', '-rw_timeout', '60000000',
    '-headers', `Authorization: Bearer ${token}\r\n`,
    '-i', url, '-map', '0:v:0', '-an', '-frames:v', String(frames), '-vf', 'idet', '-f', 'null', '-'];
}

/** "30000/1001" -> "29.97", "25/1" -> "25", "24000/1001" -> "23.976"; '' when unknown. */
function frameRate(ratio) {
  const [n, d] = String(ratio || '').split('/').map(Number);
  if (!(n > 0) || !(d > 0)) return '';
  return String(Number((n / d).toFixed(3)));
}

// ffprobe's ProRes profile names, as Apple names the flavours.
const PRORES = { Proxy: '422 Proxy', LT: '422 LT', Standard: '422', HQ: '422 HQ', '4444': '4444', XQ: '4444 XQ' };

/** "ProRes 422 HQ", "H.264 High", "H.265 Main 10", or ffprobe's own name for anything else. */
function codecName(v) {
  const profile = v.profile && v.profile !== 'unknown' ? v.profile : '';
  if (v.codec_name === 'prores') return 'ProRes ' + (PRORES[profile] || profile || '(unknown flavour)');
  if (v.codec_name === 'h264') return ('H.264 ' + profile).trim();
  if (v.codec_name === 'hevc') return ('H.265 ' + profile).trim();
  return [v.codec_long_name || v.codec_name || 'unknown', profile].filter(Boolean).join(' ');
}

/** 8, 10, 12 … from bits_per_raw_sample, else from the pixel format (yuv422p10le -> 10). */
function bitDepth(v) {
  const raw = Number(v.bits_per_raw_sample);
  if (raw > 0) return raw;
  const m = /p(\d+)(le|be)$/.exec(v.pix_fmt || '');
  if (m) return Number(m[1]);
  return v.pix_fmt ? 8 : '';
}

/** Whether the pixel format carries alpha (yuva444p12le, rgba, gbrap …). */
function hasAlpha(pixFmt) {
  return /^(yuva|rgba|bgra|argb|abgr|gbrap|ya\d)/.test(pixFmt || '');
}

/**
 * The audio tracks: identical neighbours grouped, then the channel total, e.g. "8 × mono (8 ch)",
 * "stereo (2 ch)", "5.1 + stereo (8 ch)", or "none".
 */
function audioLayout(streams) {
  const audio = streams.filter((s) => s.codec_type === 'audio');
  if (!audio.length) return 'none';
  const names = audio.map((s) => {
    const layout = s.channel_layout && s.channel_layout !== 'unknown' ? s.channel_layout.replace(/\(side\)$/, '') : '';
    return layout || (s.channels === 1 ? 'mono' : s.channels === 2 ? 'stereo' : `${s.channels} ch`);
  });
  const groups = [];
  names.forEach((n) => {
    const last = groups[groups.length - 1];
    if (last && last.name === n) last.count++;
    else groups.push({ name: n, count: 1 });
  });
  const total = audio.reduce((sum, s) => sum + (Number(s.channels) || 0), 0);
  return groups.map((g) => (g.count > 1 ? `${g.count} × ${g.name}` : g.name)).join(' + ') + ` (${total} ch)`;
}

/**
 * How the video is scanned, from what the file says: "Progressive", "Interlaced TFF" or
 * "Interlaced BFF", or null when it has to be measured.
 *   H.265: always progressive (interlaced H.265 is rare and often wrongly flagged).
 *   ProRes: from the first frame; every ProRes frame says how it is coded, while the container
 *           may not (FFmpeg reports interlaced ProRes in MXF as progressive).
 *   Others: from the stream's field order; tb and bt are ambiguous across containers, so measured.
 */
function declaredScan(v, firstFrame) {
  if (v.codec_name === 'hevc') return 'Progressive';
  if (v.codec_name === 'prores' && firstFrame) {
    if (!firstFrame.interlaced_frame) return 'Progressive';
    return firstFrame.top_field_first ? 'Interlaced TFF' : 'Interlaced BFF';
  }
  switch (v.field_order) {
    case 'progressive': return 'Progressive';
    case 'tt': return 'Interlaced TFF';
    case 'bb': return 'Interlaced BFF';
    default: return null;
  }
}

/**
 * idet's verdict from ffmpeg's output: the "Multi frame detection" counts (more reliable than the
 * single-frame ones), as "Progressive (measured)", "Interlaced TFF (measured)" …, or "Unknown".
 */
function measuredScan(stderr) {
  // ffmpeg can print the counts more than once (once per filter instance, some all zeros): use the
  // report that counted the most frames.
  const re = /Multi frame detection:\s*TFF:\s*(\d+)\s*BFF:\s*(\d+)\s*Progressive:\s*(\d+)/g;
  let best = [0, 0, 0];
  for (let m; (m = re.exec(stderr || ''));) {
    const counts = m.slice(1, 4).map(Number);
    if (counts[0] + counts[1] + counts[2] > best[0] + best[1] + best[2]) best = counts;
  }
  const [tff, bff, prog] = best;
  if (!(tff + bff + prog)) return 'Unknown';
  const verdict = prog >= tff && prog >= bff ? 'Progressive' : tff >= bff ? 'Interlaced TFF' : 'Interlaced BFF';
  return verdict + ' (measured)';
}

/**
 * ffprobe's JSON as the sheet's values (RESULT_HEADERS), plus `needsMeasuring` when the scan has to
 * be measured. Throws when the file has no video.
 */
function describe(json) {
  const streams = json.streams || [];
  const v = streams.find((s) => s.codec_type === 'video' && !(s.disposition && s.disposition.attached_pic));
  if (!v) throw new Error('no video stream (ffprobe sees: ' + ((json.format && json.format.format_name) || 'nothing') + ')');
  const firstFrame = (json.frames || []).find((f) => f.media_type === 'video' && f.stream_index === v.index);
  const scan = declaredScan(v, firstFrame);
  return {
    values: {
      'Frame rate': frameRate(v.avg_frame_rate) || frameRate(v.r_frame_rate),
      Width: v.width || '',
      Height: v.height || '',
      Codec: codecName(v),
      Scan: scan || '',
      'Bit depth': bitDepth(v),
      Alpha: hasAlpha(v.pix_fmt) ? 'Yes' : 'No',
      Audio: audioLayout(streams),
    },
    needsMeasuring: !scan,
  };
}

/** Runs a command; resolves with { stdout, stderr }, rejects with its error output on failure. */
function runCommand(cmd, args, { timeoutMs = 300000 } = {}) {
  return new Promise((resolve, reject) => {
    execFile(cmd, args, { timeout: timeoutMs, maxBuffer: 32 * 1024 * 1024 }, (err, stdout, stderr) => {
      if (err) {
        // Never pass the arguments on: they hold the access token.
        const reason = err.killed ? `timed out after ${timeoutMs / 1000} s` : String(stderr || err.message).trim().split('\n').pop();
        reject(new Error(`${cmd} failed: ${reason}`));
      } else {
        resolve({ stdout, stderr });
      }
    });
  });
}

/**
 * Probes one Drive file: { values } for the sheet. deps: url (the file's download URL), token (an
 * access token for it), run (runCommand in production).
 */
async function probeFile({ url, token, run = runCommand }) {
  const { stdout } = await run('ffprobe', ffprobeArgs(url, token), { timeoutMs: 180000 });
  const { values, needsMeasuring } = describe(JSON.parse(stdout));
  if (needsMeasuring) {
    // idet prints its counts on stderr, also when ffmpeg succeeds.
    const { stderr } = await run('ffmpeg', idetArgs(url, token), { timeoutMs: 600000 });
    values.Scan = measuredScan(stderr);
  }
  return { values };
}

module.exports = {
  RESULT_HEADERS, ffprobeArgs, idetArgs, frameRate, codecName, bitDepth, hasAlpha, audioLayout,
  declaredScan, measuredScan, describe, runCommand, probeFile,
};
