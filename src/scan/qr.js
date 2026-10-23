// Finds the sheet's QR code (ER1|<sheet id>) on a page image, trying harder when the first try fails (Session 6I):
// the top-right corner where it is printed, then the whole page; at several scales; plain, contrast-stretched and
// thresholded. Phone photos are rarely sharp, level and evenly lit. Returns the canonical sheet id, or null.

const jsQR = require('jsqr');
const { parseSheetId } = require('../sheetPrint/ids');

function grey(rgba) {
  const out = new Uint8ClampedArray(rgba.width * rgba.height);
  for (let i = 0, j = 0; i < out.length; i += 1, j += 4) out[i] = (rgba.data[j] * 299 + rgba.data[j + 1] * 587 + rgba.data[j + 2] * 114) / 1000;
  return out;
}

const toRgba = (values, width, height) => {
  const data = new Uint8ClampedArray(width * height * 4);
  for (let i = 0; i < values.length; i += 1) {
    data[i * 4] = values[i];
    data[i * 4 + 1] = values[i];
    data[i * 4 + 2] = values[i];
    data[i * 4 + 3] = 255;
  }
  return { data, width, height };
};

function crop(values, width, x0, y0, w, h) {
  const out = new Uint8ClampedArray(w * h);
  for (let y = 0; y < h; y += 1) out.set(values.subarray((y0 + y) * width + x0, (y0 + y) * width + x0 + w), y * w);
  return { values: out, width: w, height: h };
}

// area-average down, bilinear up
function scale(image, factor) {
  if (factor === 1) return image;
  const width = Math.max(8, Math.round(image.width * factor));
  const height = Math.max(8, Math.round(image.height * factor));
  const out = new Uint8ClampedArray(width * height);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const sx = (x + 0.5) / factor - 0.5;
      const sy = (y + 0.5) / factor - 0.5;
      const x0 = Math.min(image.width - 1, Math.max(0, Math.floor(sx)));
      const y0 = Math.min(image.height - 1, Math.max(0, Math.floor(sy)));
      const x1 = Math.min(image.width - 1, x0 + 1);
      const y1 = Math.min(image.height - 1, y0 + 1);
      const fx = Math.min(1, Math.max(0, sx - x0));
      const fy = Math.min(1, Math.max(0, sy - y0));
      const v = (image.values[y0 * image.width + x0] * (1 - fx) + image.values[y0 * image.width + x1] * fx) * (1 - fy) + (image.values[y1 * image.width + x0] * (1 - fx) + image.values[y1 * image.width + x1] * fx) * fy;
      out[y * width + x] = v;
    }
  }
  return { values: out, width, height };
}

function stretch(image) {
  const sorted = Uint8ClampedArray.from(image.values).sort();
  const low = sorted[Math.floor(sorted.length * 0.02)];
  const high = sorted[Math.floor(sorted.length * 0.98)];
  const range = Math.max(1, high - low);
  return { ...image, values: image.values.map((v) => Math.min(255, Math.max(0, ((v - low) * 255) / range))) };
}

function otsu(image) {
  const histogram = new Array(256).fill(0);
  image.values.forEach((v) => { histogram[Math.round(v)] += 1; });
  const total = image.values.length;
  let sum = 0;
  histogram.forEach((count, level) => { sum += level * count; });
  let background = 0;
  let weightB = 0;
  let best = 0;
  let threshold = 128;
  histogram.forEach((count, level) => {
    weightB += count;
    if (!weightB) return;
    const weightF = total - weightB;
    if (!weightF) return;
    background += level * count;
    const between = weightB * weightF * ((background / weightB - (sum - background) / weightF) ** 2);
    if (between > best) { best = between; threshold = level; }
  });
  return { ...image, values: image.values.map((v) => (v > threshold ? 255 : 0)) };
}

const PREPS = [['raw', (image) => image], ['stretch', stretch], ['otsu', (image) => otsu(stretch(image))]];
const SCALES = [1, 0.6, 1.5, 0.4, 2.2];

// rgba: { data, width, height }. Returns { sheetId, attempt } or null.
function findSheetId(rgba) {
  const full = { values: grey(rgba), width: rgba.width, height: rgba.height };
  const regions = [
    ['top-right', crop(full.values, full.width, Math.floor(full.width * 0.45), 0, full.width - Math.floor(full.width * 0.45), Math.floor(full.height * 0.4))],
    ['page', full]
  ];
  for (const [regionName, region] of regions) {
    for (const factor of SCALES) {
      const scaled = scale(region, factor);
      for (const [prepName, prep] of PREPS) {
        const image = prep(scaled);
        const code = jsQR(toRgba(image.values, image.width, image.height).data, image.width, image.height, { inversionAttempts: 'attemptBoth' });
        const sheetId = code && parseSheetId(code.data);
        if (sheetId) return { sheetId, attempt: `${regionName} x${factor} ${prepName}` };
      }
    }
  }
  return null;
}

module.exports = { findSheetId };
