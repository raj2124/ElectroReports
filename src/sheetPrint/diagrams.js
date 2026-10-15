// SET-UP diagrams for the printed measurement sheets. A one-to-one port of
// reference/generators/diagrams.py (arrangements follow IS 3043:2018 Cl. 40-42 and
// IEEE 81-2012 Cl. 9-10); each function draws inside (x, y, w, h), y up.
const NAVY = '#0B2A67';
const GREY = '#9AA3B2';
const RED = '#B42318';
const EARTH = '#8A6D3B';
const GREEN = '#2E7D32';
const BLACK = '#000000';
const WHITE = '#FFFFFF';

function t(c, x, y, s, size = 5.8, bold = false, col = BLACK, anchor = 'l') {
  c.setFont(bold ? 'Arial-Bold' : 'Arial', size);
  c.setFillColor(col);
  c.text(x, y, s, { anchor });
}

function ground(c, x1, x2, y) {
  c.setStrokeColor(EARTH);
  c.setLineWidth(1);
  c.line(x1, y, x2, y);
  c.setLineWidth(0.4);
  for (let xx = x1; xx < x2 - 3; xx += 6) {
    c.line(xx, y, xx + 3, y - 3);
  }
  c.setStrokeColor(BLACK);
}

function probe(c, x, yg, depth, label, col = NAVY) {
  c.setStrokeColor(col);
  c.setLineWidth(1.4);
  c.line(x, yg + 4, x, yg - depth);
  c.setLineWidth(0.5);
  t(c, x, yg + 6, label, 5.8, true, col, 'c');
}

function dim(c, x1, x2, y, label, size = 5.5) {
  c.setStrokeColor(GREY);
  c.setLineWidth(0.5);
  c.line(x1, y, x2, y);
  [[x1, 1], [x2, -1]].forEach(([xx, d]) => {
    c.line(xx, y, xx + 2.5 * d, y + 1.5);
    c.line(xx, y, xx + 2.5 * d, y - 1.5);
    c.line(xx, y - 3, xx, y + 3);
  });
  t(c, (x1 + x2) / 2, y + 2, label, size, false, BLACK, 'c');
}

function box(c, x, y, w, h, label) {
  c.setStrokeColor(NAVY);
  c.setFillColor(WHITE);
  c.setLineWidth(0.8);
  c.rect(x, y, w, h, { fill: true });
  t(c, x + w / 2, y + h / 2 - 2, label, 5.8, true, NAVY, 'c');
}

function wire(c, pts, col = NAVY, dash = null) {
  c.setStrokeColor(col);
  c.setLineWidth(0.6);
  if (dash) {
    c.setDash(dash[0], dash[1]);
  }
  c.polyline(pts);
  c.setDash();
}

function wenner(c, x, y, w, h) {
  const yg = y + h * 0.52;
  const sp = (w * 0.62) / 3;
  const x0 = x + 12;
  const xs = [0, 1, 2, 3].map((i) => x0 + i * sp);
  ground(c, x0 - 8, xs[3] + 8, yg);
  ['C1', 'P1', 'P2', 'C2'].forEach((lab, i) => probe(c, xs[i], yg, 7, lab, lab[0] === 'P' ? NAVY : RED));
  const bx = (xs[1] + xs[2]) / 2 - 20;
  const by = y + h - 12;
  box(c, bx, by, 40, 11, 'Tester');
  wire(c, [[bx, by + 5], [xs[0], by + 5], [xs[0], yg + 13]], RED);
  wire(c, [[bx + 40, by + 5], [xs[3], by + 5], [xs[3], yg + 13]], RED);
  wire(c, [[bx + 8, by], [xs[1], yg + 13]]);
  wire(c, [[bx + 32, by], [xs[2], yg + 13]]);
  for (let i = 0; i < 3; i += 1) {
    dim(c, xs[i], xs[i + 1], yg - 15, 'a');
  }
  t(c, x + 2, y + 1, 'Equal spacing a · 10–15 cm deep · I via C1–C2, V across P1–P2 · ρ = 2πaR', 5.2, true, NAVY);
  const cx = x + w * 0.89;
  const cy = y + h * 0.55;
  const r = Math.min(h * 0.28, 15);
  wire(c, [[cx - r, cy], [cx + r, cy]], NAVY);
  wire(c, [[cx, cy - r], [cx, cy + r]], NAVY);
  wire(c, [[cx - r * 0.7, cy - r * 0.7], [cx + r * 0.7, cy + r * 0.7]], GREY, [1, 1.5]);
  wire(c, [[cx - r * 0.7, cy + r * 0.7], [cx + r * 0.7, cy - r * 0.7]], GREY, [1, 1.5]);
  t(c, cx + r + 2, cy - 2, 'D1', 5.5, true, NAVY);
  t(c, cx, cy + r + 2, 'D2', 5.5, true, NAVY, 'c');
  t(c, cx, cy - r - 7, 'D3/D4 dashed', 5, false, GREY, 'c');
}

function fop(c, x, y, w, h, grid = false) {
  const yg = y + h * 0.58;
  const xe = x + 18;
  const xc = x + w - 12;
  const D = xc - xe;
  ground(c, x + 4, x + w - 4, yg);
  if (grid) {
    c.setStrokeColor(NAVY);
    c.setLineWidth(0.6);
    c.rect(xe - 12, yg - 6, 20, 4);
  } else {
    c.setStrokeColor(NAVY);
    c.setLineWidth(1.6);
    c.line(xe, yg, xe, yg - 9);
    c.setLineWidth(0.6);
    c.setStrokeColor(RED);
    c.rect(xe - 4, yg + 1, 8, 4);
  }
  t(c, xe + (grid ? 8 : 7), yg + 3, grid ? 'Grid L' : 'E (link open)', 5.2, true, grid ? NAVY : RED);
  [[0.5, '50%'], [0.618, '61.8%'], [0.7, '70%']].forEach(([f, lab]) => {
    const xp = xe + f * D;
    if (f === 0.618) {
      probe(c, xp, yg, 6, 'P', NAVY);
    } else {
      c.setStrokeColor(GREY);
      c.setDash(1, 1.2);
      c.line(xp, yg + 3, xp, yg - 6);
      c.setDash();
    }
    t(c, xp, yg - 11, lab, 5.2, f === 0.618, f === 0.618 ? NAVY : BLACK, 'c');
  });
  probe(c, xc, yg, 7, 'C', RED);
  const bx = xe + D * 0.2;
  const by = y + h - 12;
  box(c, bx, by, 36, 11, 'Tester');
  wire(c, [[bx, by + 5], [xe, by + 5], [xe, yg + 6]]);
  wire(c, [[bx + 30, by], [xe + 0.618 * D, yg + 13]]);
  wire(c, [[bx + 36, by + 5], [xc, by + 5], [xc, yg + 13]], RED);
  dim(c, xe, xc, yg - 20, grid ? 'D per Table 17 (L 5 / 10 / 20 m → D 100 / 150 / 200 m)' : 'D (E to C) ≥ 30 m for rod / pipe / plate – Cl. 41.1.5', 5);
  t(c, x + 2, y + 1, 'Read at 61.8%, confirm at 50% and 70% – values must agree (IS 3043:2018 Cl. 41.1.6)', 5.2, true, NAVY);
}

function continuity(c, x, y, w, h) {
  const yb = y + h * 0.42;
  box(c, x + 6, yb - 6, 44, 12, 'Earth bar / ref.');
  c.setStrokeColor(NAVY);
  c.setLineWidth(1.2);
  c.line(x + 50, yb, x + w - 58, yb);
  t(c, (x + w) / 2 - 10, yb + 2, 'protective / bonding conductor, length l', 5.2, false, BLACK, 'c');
  box(c, x + w - 58, yb - 10, 50, 20, 'Equipment');
  t(c, x + w - 33, yb - 17, 'frame / body', 5);
  const bx = (x + w) / 2 - 22;
  const by = y + h - 15;
  box(c, bx, by, 44, 12, 'Low-Ω tester');
  wire(c, [[bx, by + 4], [x + 28, by + 4], [x + 28, yb + 6]]);
  wire(c, [[bx + 44, by + 4], [x + w - 33, by + 4], [x + w - 33, yb + 10]]);
  t(c, x + 4, y + 9, '1. Short the leads and null (record lead R)  2. Clean paint/rust at both points', 5.2);
  t(c, x + 4, y + 1, '3. Read R between reference and tested point (IEEE 81-2012 Cl. 10)', 5.2, true, NAVY);
}

function loop(c, x, y, w, h, pfc = false) {
  const top = y + h - 8;
  const bot = y + 13;
  const mid = (top + bot) / 2;
  box(c, x + 4, mid - 9, 30, 18, 'Tx');
  c.setStrokeColor(NAVY);
  c.setLineWidth(0.9);
  c.line(x + 34, top - 4, x + w * 0.62, top - 4);
  t(c, x + 40, top - 1, 'L (phase)', 5.2, true, RED);
  c.line(x + 34, mid, x + w * 0.62, mid);
  t(c, x + 40, mid + 2, 'N', 5.2, true, NAVY);
  c.setStrokeColor(GREEN);
  c.line(x + 34, bot + 4, x + w * 0.62, bot + 4);
  t(c, x + 40, bot + 6, 'PE / earth', 5.2, true, GREEN);
  const dx = x + w * 0.62;
  box(c, dx, bot - 2, 30, top - bot, 'Device');
  const tx = x + w * 0.8;
  box(c, tx, mid - 8, 42, 16, '4118A');
  wire(c, [[dx + 30, top - 4], [tx + 6, top - 4], [tx + 6, mid + 8]], RED);
  if (pfc) {
    wire(c, [[dx + 30, mid], [tx, mid]], NAVY);
  }
  wire(c, [[dx + 30, bot + 4], [tx + 36, bot + 4], [tx + 36, mid - 8]], GREEN);
  if (pfc) {
    t(c, x + 4, y + 1, 'L-PE → PFC · L-N → PSC · device breaking capacity (kA) must exceed the higher', 5.2, true, NAVY);
  } else {
    c.setStrokeColor(RED);
    c.setDash(2, 1.5);
    c.setLineWidth(0.6);
    c.rect(x + 22, bot + 1, dx - x - 6, top - bot - 2);
    c.setDash();
    t(c, x + 4, y + 1, 'Zs = whole fault loop: phase → fault → PE → source (IS 3043:2018 Cl. 42)', 5.2, true, NAVY);
  }
}

function riser(c, x, y, w, h) {
  const yg = y + h * 0.46;
  ground(c, x + 4, x + w - 4, yg);
  c.setStrokeColor(NAVY);
  c.setLineWidth(1.2);
  c.line(x + 10, yg - 7, x + w - 10, yg - 7);
  t(c, x + w - 12, yg - 14, 'buried earth grid', 5.2, false, NAVY, 'r');
  const xa = x + w * 0.2;
  const xr = x + w * 0.62;
  [[xa, 'Reference riser'], [xr, 'Riser under test']].forEach(([xx, lab]) => {
    c.setStrokeColor(NAVY);
    c.setLineWidth(1.2);
    c.line(xx, yg - 7, xx, y + h - 10);
    t(c, xx, y + h - 7, lab, 5.2, true, NAVY, 'c');
  });
  box(c, xr + 14, y + h - 24, 38, 12, 'Equipment');
  c.setLineWidth(1.2);
  c.line(xr, y + h - 18, xr + 14, y + h - 18);
  const mp = [xr, yg + 8];
  c.setFillColor(RED);
  c.circle(mp[0], mp[1], 2, { fill: true });
  wire(c, [mp, [xr + 33, y + h - 24]], RED, [2, 1.5]);
  t(c, xr + 38, yg + 8, '(1) towards equipment', 5.3, true, RED);
  wire(c, [mp, [xa, yg + 8]], RED, [2, 1.5]);
  t(c, (xa + xr) / 2, yg + 11, '(2) towards grid', 5.3, true, RED, 'c');
  t(c, x + 2, y + 1, 'Null leads first · same reference riser for the area · compare risers (IEEE 81-2012 Cl. 10)', 5.2, true, NAVY);
}

function stepTouch(c, x, y, w, h) {
  const yg = y + h * 0.56;
  ground(c, x + 4, x + w - 4, yg);
  const sx = x + 18;
  c.setStrokeColor(NAVY);
  c.setLineWidth(1.6);
  c.line(sx, yg, sx, y + h - 2);
  if (sx - 3 > x + 30) {
    t(c, sx - 3, y + h - 8, 'structure', 5, true, NAVY, 'r');
  } else {
    t(c, sx + 3, yg + 3, 'structure', 5, true, NAVY);
  }
  const pt = sx + w * 0.15;
  probe(c, pt, yg, 5, '', NAVY);
  box(c, sx + 10, y + h - 13, 30, 11, 'V touch');
  wire(c, [[sx, y + h - 8], [sx + 10, y + h - 8]]);
  wire(c, [[sx + 40, y + h - 8], [pt, y + h - 8], [pt, yg + 4]]);
  dim(c, sx, pt, yg - 10, '≤ 1 m', 5);
  const p1 = x + w * 0.42;
  const p2 = x + w * 0.42 + w * 0.15;
  probe(c, p1, yg, 5, '', NAVY);
  probe(c, p2, yg, 5, '', NAVY);
  box(c, (p1 + p2) / 2 - 15, y + h - 13, 30, 11, 'V step');
  wire(c, [[p1, yg + 4], [p1, y + h - 8], [(p1 + p2) / 2 - 15, y + h - 8]]);
  wire(c, [[p2, yg + 4], [p2, y + h - 8], [(p1 + p2) / 2 + 15, y + h - 8]]);
  dim(c, p1, p2, yg - 10, '1 m', 5);
  const xc = x + w - 12;
  probe(c, xc, yg, 7, 'C', RED);
  t(c, xc - 7, yg + 3, 'remote current electrode', 5, false, RED, 'r');
  t(c, x + 2, y + 1, 'Record It per reading · V actual = V measured × If / It → compare with tolerable limit (IEEE 81-2012 Cl. 9.3)', 5.1, true, NAVY);
}

function tower(c, x, y, w, h) {
  const cx = x + w * 0.24;
  const cy = y + h * 0.55;
  const s = Math.min(h * 0.62, 34);
  c.setStrokeColor(NAVY);
  c.setLineWidth(0.8);
  c.rect(cx - s / 2, cy - s / 2, s, s);
  c.line(cx - s / 2, cy - s / 2, cx + s / 2, cy + s / 2);
  c.line(cx - s / 2, cy + s / 2, cx + s / 2, cy - s / 2);
  [[-1, 1], [1, 1], [1, -1], [-1, -1]].forEach(([dx, dy], index) => {
    const fx = cx + (dx * s) / 2;
    const fy = cy + (dy * s) / 2;
    c.setFillColor(NAVY);
    c.circle(fx, fy, 2.2, { fill: true });
    c.setStrokeColor(RED);
    c.setLineWidth(0.9);
    c.circle(fx, fy, 5);
    t(c, fx + dx * 9, fy + (dy > 0 ? 3 : -7), `F-${index + 1}`, 5.5, true, NAVY, 'c');
  });
  t(c, cx, y + 2, 'plan view – clamp (red) on each foot earth connection', 5, false, RED, 'c');
  const xp = x + w * 0.62;
  const xc = x + w - 12;
  c.setStrokeColor(GREY);
  c.setLineWidth(0.5);
  c.setDash(2, 1.5);
  c.line(cx + s / 2, cy, xc, cy);
  c.setDash();
  probe(c, xp, cy, 0, 'P', NAVY);
  probe(c, xc, cy, 0, 'C', RED);
  c.setFillColor(NAVY);
  c.circle(xp, cy, 2, { fill: true });
  c.setFillColor(RED);
  c.circle(xc, cy, 2, { fill: true });
  t(c, (xp + xc) / 2, cy - 10, 'auxiliary probes along a straight line', 5, false, BLACK, 'c');
  t(c, x + w * 0.5, y + 10, 'Earth wire stays connected · one reading per foot', 5.2);
  t(c, x + w * 0.5, y + 2, 'Zt = Σ 4 feet vs Zsat (acceptance list)', 5.2, true, NAVY);
}

const FIGURES = {
  soilResistivity: wenner,
  electrodeResistance: fop,
  continuityTest: continuity,
  loopImpedanceTest: loop,
  prospectiveFaultCurrent: (c, x, y, w, h) => loop(c, x, y, w, h, true),
  riserIntegrityTest: riser,
  groundResistance: (c, x, y, w, h) => fop(c, x, y, w, h, true),
  stepTouchVoltage: stepTouch,
  towerFootingResistance: tower
};

module.exports = { FIGURES, NAVY, GREY };
