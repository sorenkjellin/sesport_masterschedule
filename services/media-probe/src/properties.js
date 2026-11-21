// The probe results as Drive file properties: small key/value pairs stored on the file itself, so
// they go wherever the file is moved (its Drive ID never changes). The Asset Manager writes the same
// keys when it moves a video out of the inbox (PROBE_PROPERTY_KEYS in apps/video-assets/src/
// mediaProbe.js; a test keeps the two the same), and reads them when it lists Video Assets.
'use strict';

// Column header -> property key. Values are strings: Drive properties are text.
const PROPERTY_KEYS = {
  Duration: 'sesport.duration', // whole seconds
  'Frame rate': 'sesport.frameRate',
  Width: 'sesport.width',
  Height: 'sesport.height',
  Codec: 'sesport.codec',
  Mbit: 'sesport.mbit',
  Scan: 'sesport.scan',
  'Bit depth': 'sesport.bitDepth',
  Alpha: 'sesport.alpha',
  Audio: 'sesport.audio',
};
const PROBED_KEY = 'sesport.probed';

/**
 * The results (values by column header, as probe.describe gives them; Duration as a fraction of a
 * day) as Drive properties, with when they were probed. Empty values are left out. A key and its
 * value may be 124 bytes together, so values are cut to 100 characters.
 */
function toProperties(values, probedAt) {
  const props = {};
  Object.keys(PROPERTY_KEYS).forEach((header) => {
    let v = values[header];
    if (v === '' || v == null) return;
    if (header === 'Duration') v = Math.round(Number(v) * 86400);
    props[PROPERTY_KEYS[header]] = String(v).slice(0, 100);
  });
  props[PROBED_KEY] = probedAt.toISOString();
  return props;
}

module.exports = { PROPERTY_KEYS, PROBED_KEY, toProperties };
