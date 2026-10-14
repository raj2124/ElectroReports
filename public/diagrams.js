// SET-UP diagrams for the measurement sheets, as inline SVG.
// A port of reference/generators/diagrams.py (the figures on the printed sheets): the same
// arrangements, geometry, colours and captions (IS 3043:2018 Cl. 40-42, IEEE 81-2012 Cl. 9-10).
// Each function draws into a box (x, y, w, h) with the origin at the bottom left, as the
// printed version does; the canvas flips y for SVG.

const DIAGRAM_COLORS = Object.freeze({ navy: '#0B2A67', grey: '#9AA3B2', red: '#B42318', earth: '#8A6D3B', green: '#2E7D32', black: '#000000', white: '#FFFFFF' });
const DIAGRAM_WIDTH = 300;
const DIAGRAM_HEIGHT = 60;

function diagramCanvas(width, height) {
  const parts = [];
  const fmt = (value) => Number(value.toFixed(2));
  const flip = (y) => fmt(height - y);
  const stroke = (color, lineWidth, dash) => `stroke="${color}" stroke-width="${lineWidth}"${dash ? ` stroke-dasharray="${dash.join(' ')}"` : ''}`;
  return {
    line(x1, y1, x2, y2, color = DIAGRAM_COLORS.black, lineWidth = 0.5, dash = null) {
      parts.push(`<line x1="${fmt(x1)}" y1="${flip(y1)}" x2="${fmt(x2)}" y2="${flip(y2)}" ${stroke(color, lineWidth, dash)} />`);
    },
    polyline(points, color = DIAGRAM_COLORS.navy, lineWidth = 0.6, dash = null) {
      parts.push(`<polyline points="${points.map(([px, py]) => `${fmt(px)},${flip(py)}`).join(' ')}" fill="none" ${stroke(color, lineWidth, dash)} />`);
    },
    rect(x, y, w, h, { color = DIAGRAM_COLORS.navy, lineWidth = 0.6, fill = 'none', dash = null } = {}) {
      parts.push(`<rect x="${fmt(x)}" y="${flip(y + h)}" width="${fmt(w)}" height="${fmt(h)}" fill="${fill}" ${stroke(color, lineWidth, dash)} />`);
    },
    circle(cx, cy, r, { color = DIAGRAM_COLORS.navy, lineWidth = 0.6, fill = 'none' } = {}) {
      parts.push(`<circle cx="${fmt(cx)}" cy="${flip(cy)}" r="${fmt(r)}" fill="${fill}" ${stroke(color, lineWidth)} />`);
    },
    text(x, y, label, size = 5.8, bold = false, color = DIAGRAM_COLORS.black, anchor = 'l') {
      const anchors = { l: 'start', c: 'middle', r: 'end' };
      parts.push(`<text x="${fmt(x)}" y="${flip(y)}" font-size="${size}" font-family="Arial, Helvetica, sans-serif" font-weight="${bold ? 700 : 400}" fill="${color}" text-anchor="${anchors[anchor]}">${diagramEscape(label)}</text>`);
    },
    svg(label) {
      return `<svg class="setup-diagram" viewBox="0 0 ${width} ${height}" role="img" aria-label="${diagramEscape(label)}" xmlns="http://www.w3.org/2000/svg">${parts.join('')}</svg>`;
    }
  };
}

function diagramEscape(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const DG = {
  ground(c, x1, x2, y) {
    c.line(x1, y, x2, y, DIAGRAM_COLORS.earth, 1);
    for (let xx = x1; xx < x2 - 3; xx += 6) {
      c.line(xx, y, xx + 3, y - 3, DIAGRAM_COLORS.earth, 0.4);
    }
  },
  probe(c, x, yg, depth, label, color = DIAGRAM_COLORS.navy) {
    c.line(x, yg + 4, x, yg - depth, color, 1.4);
    c.text(x, yg + 6, label, 5.8, true, color, 'c');
  },
  dim(c, x1, x2, y, label, size = 5.5) {
    c.line(x1, y, x2, y, DIAGRAM_COLORS.grey, 0.5);
    [[x1, 1], [x2, -1]].forEach(([xx, d]) => {
      c.line(xx, y, xx + 2.5 * d, y + 1.5, DIAGRAM_COLORS.grey, 0.5);
      c.line(xx, y, xx + 2.5 * d, y - 1.5, DIAGRAM_COLORS.grey, 0.5);
      c.line(xx, y - 3, xx, y + 3, DIAGRAM_COLORS.grey, 0.5);
    });
    c.text((x1 + x2) / 2, y + 2, label, size, false, DIAGRAM_COLORS.black, 'c');
  },
  box(c, x, y, w, h, label) {
    c.rect(x, y, w, h, { color: DIAGRAM_COLORS.navy, lineWidth: 0.8, fill: DIAGRAM_COLORS.white });
    c.text(x + w / 2, y + h / 2 - 2, label, 5.8, true, DIAGRAM_COLORS.navy, 'c');
  },
  wire(c, points, color = DIAGRAM_COLORS.navy, dash = null) {
    c.polyline(points, color, 0.6, dash);
  },

  wenner(c, x, y, w, h) {
    const { navy, red, grey } = DIAGRAM_COLORS;
    const yg = y + h * 0.52;
    const sp = (w * 0.62) / 3;
    const x0 = x + 12;
    const xs = [0, 1, 2, 3].map((i) => x0 + i * sp);
    DG.ground(c, x0 - 8, xs[3] + 8, yg);
    ['C1', 'P1', 'P2', 'C2'].forEach((label, i) => DG.probe(c, xs[i], yg, 7, label, label[0] === 'P' ? navy : red));
    const bx = (xs[1] + xs[2]) / 2 - 20;
    const by = y + h - 12;
    DG.box(c, bx, by, 40, 11, 'Tester');
    DG.wire(c, [[bx, by + 5], [xs[0], by + 5], [xs[0], yg + 13]], red);
    DG.wire(c, [[bx + 40, by + 5], [xs[3], by + 5], [xs[3], yg + 13]], red);
    DG.wire(c, [[bx + 8, by], [xs[1], yg + 13]]);
    DG.wire(c, [[bx + 32, by], [xs[2], yg + 13]]);
    for (let i = 0; i < 3; i += 1) {
      DG.dim(c, xs[i], xs[i + 1], yg - 15, 'a');
    }
    c.text(x + 2, y + 1, 'Equal spacing a · 10–15 cm deep · I via C1–C2, V across P1–P2 · ρ = 2πaR', 5.2, true, navy);
    const cx = x + w * 0.89;
    const cy = y + h * 0.55;
    const r = Math.min(h * 0.28, 15);
    DG.wire(c, [[cx - r, cy], [cx + r, cy]], navy);
    DG.wire(c, [[cx, cy - r], [cx, cy + r]], navy);
    DG.wire(c, [[cx - r * 0.7, cy - r * 0.7], [cx + r * 0.7, cy + r * 0.7]], grey, [1, 1.5]);
    DG.wire(c, [[cx - r * 0.7, cy + r * 0.7], [cx + r * 0.7, cy - r * 0.7]], grey, [1, 1.5]);
    c.text(cx + r + 2, cy - 2, 'D1', 5.5, true, navy);
    c.text(cx, cy + r + 2, 'D2', 5.5, true, navy, 'c');
    c.text(cx, cy - r - 7, 'D3/D4 dashed', 5, false, grey, 'c');
  },

  fop(c, x, y, w, h, grid = false) {
    const { navy, red, grey, black } = DIAGRAM_COLORS;
    const yg = y + h * 0.58;
    const xe = x + 18;
    const xc = x + w - 12;
    const D = xc - xe;
    DG.ground(c, x + 4, x + w - 4, yg);
    if (grid) {
      c.rect(xe - 12, yg - 6, 20, 4, { color: navy, lineWidth: 0.6 });
    } else {
      c.line(xe, yg, xe, yg - 9, navy, 1.6);
      c.rect(xe - 4, yg + 1, 8, 4, { color: red, lineWidth: 0.6 });
    }
    c.text(xe + (grid ? 8 : 7), yg + 3, grid ? 'Grid L' : 'E (link open)', 5.2, true, grid ? navy : red);
    [[0.5, '50%'], [0.618, '61.8%'], [0.7, '70%']].forEach(([f, label]) => {
      const xp = xe + f * D;
      if (f === 0.618) {
        DG.probe(c, xp, yg, 6, 'P', navy);
      } else {
        c.line(xp, yg + 3, xp, yg - 6, grey, 0.5, [1, 1.2]);
      }
      c.text(xp, yg - 11, label, 5.2, f === 0.618, f === 0.618 ? navy : black, 'c');
    });
    DG.probe(c, xc, yg, 7, 'C', red);
    const bx = xe + D * 0.2;
    const by = y + h - 12;
    DG.box(c, bx, by, 36, 11, 'Tester');
    DG.wire(c, [[bx, by + 5], [xe, by + 5], [xe, yg + 6]]);
    DG.wire(c, [[bx + 30, by], [xe + 0.618 * D, yg + 13]]);
    DG.wire(c, [[bx + 36, by + 5], [xc, by + 5], [xc, yg + 13]], red);
    DG.dim(c, xe, xc, yg - 20, grid ? 'D per Table 17 (L 5 / 10 / 20 m → D 100 / 150 / 200 m)' : 'D (E to C) ≥ 30 m for rod / pipe / plate – Cl. 41.1.5', 5);
    c.text(x + 2, y + 1, 'Read at 61.8%, confirm at 50% and 70% – values must agree (IS 3043:2018 Cl. 41.1.6)', 5.2, true, navy);
  },

  continuity(c, x, y, w, h) {
    const { navy } = DIAGRAM_COLORS;
    const yb = y + h * 0.42;
    DG.box(c, x + 6, yb - 6, 44, 12, 'Earth bar / ref.');
    c.line(x + 50, yb, x + w - 58, yb, navy, 1.2);
    c.text((x + w) / 2 - 10, yb + 2, 'protective / bonding conductor, length l', 5.2, false, DIAGRAM_COLORS.black, 'c');
    DG.box(c, x + w - 58, yb - 10, 50, 20, 'Equipment');
    c.text(x + w - 33, yb - 17, 'frame / body', 5);
    const bx = (x + w) / 2 - 22;
    const by = y + h - 15;
    DG.box(c, bx, by, 44, 12, 'Low-Ω tester');
    DG.wire(c, [[bx, by + 4], [x + 28, by + 4], [x + 28, yb + 6]]);
    DG.wire(c, [[bx + 44, by + 4], [x + w - 33, by + 4], [x + w - 33, yb + 10]]);
    c.text(x + 4, y + 9, '1. Short the leads and null (record lead R)  2. Clean paint/rust at both points', 5.2);
    c.text(x + 4, y + 1, '3. Read R between reference and tested point (IEEE 81-2012 Cl. 10)', 5.2, true, navy);
  },

  loop(c, x, y, w, h, pfc = false) {
    const { navy, red, green } = DIAGRAM_COLORS;
    const top = y + h - 8;
    const bot = y + 13;
    const mid = (top + bot) / 2;
    DG.box(c, x + 4, mid - 9, 30, 18, 'Tx');
    c.line(x + 34, top - 4, x + w * 0.62, top - 4, navy, 0.9);
    c.text(x + 40, top - 1, 'L (phase)', 5.2, true, red);
    c.line(x + 34, mid, x + w * 0.62, mid, navy, 0.9);
    c.text(x + 40, mid + 2, 'N', 5.2, true, navy);
    c.line(x + 34, bot + 4, x + w * 0.62, bot + 4, green, 0.9);
    c.text(x + 40, bot + 6, 'PE / earth', 5.2, true, green);
    const dx = x + w * 0.62;
    DG.box(c, dx, bot - 2, 30, top - bot, 'Device');
    const tx = x + w * 0.8;
    DG.box(c, tx, mid - 8, 42, 16, '4118A');
    DG.wire(c, [[dx + 30, top - 4], [tx + 6, top - 4], [tx + 6, mid + 8]], red);
    if (pfc) {
      DG.wire(c, [[dx + 30, mid], [tx, mid]], navy);
    }
    DG.wire(c, [[dx + 30, bot + 4], [tx + 36, bot + 4], [tx + 36, mid - 8]], green);
    if (pfc) {
      c.text(x + 4, y + 1, 'L-PE → PFC · L-N → PSC · device breaking capacity (kA) must exceed the higher', 5.2, true, navy);
    } else {
      c.rect(x + 22, bot + 1, dx - x - 6, top - bot - 2, { color: red, lineWidth: 0.6, dash: [2, 1.5] });
      c.text(x + 4, y + 1, 'Zs = whole fault loop: phase → fault → PE → source (IS 3043:2018 Cl. 42)', 5.2, true, navy);
    }
  },

  riser(c, x, y, w, h) {
    const { navy, red } = DIAGRAM_COLORS;
    const yg = y + h * 0.46;
    DG.ground(c, x + 4, x + w - 4, yg);
    c.line(x + 10, yg - 7, x + w - 10, yg - 7, navy, 1.2);
    c.text(x + w - 12, yg - 14, 'buried earth grid', 5.2, false, navy, 'r');
    const xa = x + w * 0.2;
    const xr = x + w * 0.62;
    [[xa, 'Reference riser'], [xr, 'Riser under test']].forEach(([xx, label]) => {
      c.line(xx, yg - 7, xx, y + h - 10, navy, 1.2);
      c.text(xx, y + h - 7, label, 5.2, true, navy, 'c');
    });
    DG.box(c, xr + 14, y + h - 24, 38, 12, 'Equipment');
    c.line(xr, y + h - 18, xr + 14, y + h - 18, navy, 1.2);
    const mp = [xr, yg + 8];
    c.circle(mp[0], mp[1], 2, { color: red, fill: red });
    DG.wire(c, [mp, [xr + 33, y + h - 24]], red, [2, 1.5]);
    c.text(xr + 38, yg + 8, '(1) towards equipment', 5.3, true, red);
    DG.wire(c, [mp, [xa, yg + 8]], red, [2, 1.5]);
    c.text((xa + xr) / 2, yg + 11, '(2) towards grid', 5.3, true, red, 'c');
    c.text(x + 2, y + 1, 'Null leads first · same reference riser for the area · compare risers (IEEE 81-2012 Cl. 10)', 5.2, true, navy);
  },

  stepTouch(c, x, y, w, h) {
    const { navy, red } = DIAGRAM_COLORS;
    const yg = y + h * 0.56;
    DG.ground(c, x + 4, x + w - 4, yg);
    const sx = x + 18;
    c.line(sx, yg, sx, y + h - 2, navy, 1.6);
    if (sx - 3 > x + 30) {
      c.text(sx - 3, y + h - 8, 'structure', 5, true, navy, 'r');
    } else {
      c.text(sx + 3, yg + 3, 'structure', 5, true, navy);
    }
    const pt = sx + w * 0.15;
    DG.probe(c, pt, yg, 5, '', navy);
    DG.box(c, sx + 10, y + h - 13, 30, 11, 'V touch');
    DG.wire(c, [[sx, y + h - 8], [sx + 10, y + h - 8]]);
    DG.wire(c, [[sx + 40, y + h - 8], [pt, y + h - 8], [pt, yg + 4]]);
    DG.dim(c, sx, pt, yg - 10, '≤ 1 m', 5);
    const p1 = x + w * 0.42;
    const p2 = p1 + w * 0.15;
    DG.probe(c, p1, yg, 5, '', navy);
    DG.probe(c, p2, yg, 5, '', navy);
    DG.box(c, (p1 + p2) / 2 - 15, y + h - 13, 30, 11, 'V step');
    DG.wire(c, [[p1, yg + 4], [p1, y + h - 8], [(p1 + p2) / 2 - 15, y + h - 8]]);
    DG.wire(c, [[p2, yg + 4], [p2, y + h - 8], [(p1 + p2) / 2 + 15, y + h - 8]]);
    DG.dim(c, p1, p2, yg - 10, '1 m', 5);
    const xc = x + w - 12;
    DG.probe(c, xc, yg, 7, 'C', red);
    c.text(xc - 7, yg + 3, 'remote current electrode', 5, false, red, 'r');
    c.text(x + 2, y + 1, 'Record It per reading · V actual = V measured × If / It → compare with tolerable limit (IEEE 81-2012 Cl. 9.3)', 5.1, true, navy);
  },

  tower(c, x, y, w, h) {
    const { navy, red, grey, black } = DIAGRAM_COLORS;
    const cx = x + w * 0.24;
    const cy = y + h * 0.55;
    const s = Math.min(h * 0.62, 34);
    c.rect(cx - s / 2, cy - s / 2, s, s, { color: navy, lineWidth: 0.8 });
    c.line(cx - s / 2, cy - s / 2, cx + s / 2, cy + s / 2, navy, 0.8);
    c.line(cx - s / 2, cy + s / 2, cx + s / 2, cy - s / 2, navy, 0.8);
    [[-1, 1], [1, 1], [1, -1], [-1, -1]].forEach(([dx, dy], index) => {
      const fx = cx + (dx * s) / 2;
      const fy = cy + (dy * s) / 2;
      c.circle(fx, fy, 2.2, { color: navy, fill: navy });
      c.circle(fx, fy, 5, { color: red, lineWidth: 0.9 });
      c.text(fx + dx * 9, fy + (dy > 0 ? 3 : -7), `F-${index + 1}`, 5.5, true, navy, 'c');
    });
    c.text(cx, y + 2, 'plan view – clamp (red) on each foot earth connection', 5, false, red, 'c');
    const xp = x + w * 0.62;
    const xc = x + w - 12;
    c.line(cx + s / 2, cy, xc, cy, grey, 0.5, [2, 1.5]);
    DG.probe(c, xp, cy, 0, 'P', navy);
    DG.probe(c, xc, cy, 0, 'C', red);
    c.circle(xp, cy, 2, { color: navy, fill: navy });
    c.circle(xc, cy, 2, { color: red, fill: red });
    c.text((xp + xc) / 2, cy - 10, 'auxiliary probes along a straight line', 5, false, black, 'c');
    c.text(x + w * 0.5, y + 10, 'Earth wire stays connected · one reading per foot', 5.2);
    c.text(x + w * 0.5, y + 2, 'Zt = Σ 4 feet vs Zsat (acceptance list)', 5.2, true, navy);
  }
};

// One figure per sheet, as on the printed sheets.
const SHEET_DIAGRAMS = {
  soilResistivity: { label: 'Wenner four-probe set-up', draw: (c, w, h) => DG.wenner(c, 0, 0, w, h) },
  electrodeResistance: { label: 'Fall-of-potential set-up for an electrode', draw: (c, w, h) => DG.fop(c, 0, 0, w, h) },
  continuityTest: { label: 'Earth continuity set-up', draw: (c, w, h) => DG.continuity(c, 0, 0, w, h) },
  loopImpedanceTest: { label: 'Earth fault loop impedance set-up', draw: (c, w, h) => DG.loop(c, 0, 0, w, h) },
  prospectiveFaultCurrent: { label: 'Prospective fault current set-up', draw: (c, w, h) => DG.loop(c, 0, 0, w, h, true) },
  riserIntegrityTest: { label: 'Riser integrity set-up', draw: (c, w, h) => DG.riser(c, 0, 0, w, h) },
  groundResistance: { label: 'Fall-of-potential set-up for a grid', draw: (c, w, h) => DG.fop(c, 0, 0, w, h, true) },
  stepTouchVoltage: { label: 'Step and touch voltage set-up', draw: (c, w, h) => DG.stepTouch(c, 0, 0, w, h) },
  towerFootingResistance: { label: 'Tower footing clamp set-up', draw: (c, w, h) => DG.tower(c, 0, 0, w, h) }
};

function sheetDiagram(sheetId) {
  const entry = SHEET_DIAGRAMS[sheetId];
  if (!entry) {
    return '';
  }
  const canvas = diagramCanvas(DIAGRAM_WIDTH, DIAGRAM_HEIGHT);
  entry.draw(canvas, DIAGRAM_WIDTH, DIAGRAM_HEIGHT);
  return canvas.svg(`SET-UP: ${entry.label}`);
}
