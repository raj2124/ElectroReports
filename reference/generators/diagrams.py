"""Setup diagrams for ESHA measurement sheets. Each function draws inside (x, y, w, h).
Arrangements follow IS 3043:2018 Cl. 40-42 and IEEE 81-2012 Cl. 9-10."""
from reportlab.lib.colors import HexColor, black, white
NAVY = HexColor("#0B2A67"); GREY = HexColor("#9AA3B2"); RED = HexColor("#B42318"); EARTH = HexColor("#8A6D3B")

def _t(c, x, y, s, size=5.8, bold=False, col=black, anchor="l"):
    c.setFont("Arial-Bold" if bold else "Arial", size); c.setFillColor(col)
    {"l": c.drawString, "c": c.drawCentredString, "r": c.drawRightString}[anchor](x, y, s)

def _ground(c, x1, x2, y):
    c.setStrokeColor(EARTH); c.setLineWidth(1); c.line(x1, y, x2, y); c.setLineWidth(0.4)
    xx = x1
    while xx < x2 - 3: c.line(xx, y, xx + 3, y - 3); xx += 6
    c.setStrokeColor(black)

def _probe(c, x, yg, depth, label, col=NAVY):
    c.setStrokeColor(col); c.setLineWidth(1.4); c.line(x, yg + 4, x, yg - depth); c.setLineWidth(0.5)
    _t(c, x, yg + 6, label, 5.8, True, col, "c")

def _dim(c, x1, x2, y, label, size=5.5):
    c.setStrokeColor(GREY); c.setLineWidth(0.5); c.line(x1, y, x2, y)
    for xx, d in ((x1, 1), (x2, -1)):
        c.line(xx, y, xx + 2.5 * d, y + 1.5); c.line(xx, y, xx + 2.5 * d, y - 1.5); c.line(xx, y - 3, xx, y + 3)
    _t(c, (x1 + x2) / 2, y + 2, label, size, False, black, "c")

def _box(c, x, y, w, h, label):
    c.setStrokeColor(NAVY); c.setFillColor(white); c.setLineWidth(0.8); c.rect(x, y, w, h, fill=1)
    _t(c, x + w / 2, y + h / 2 - 2, label, 5.8, True, NAVY, "c")

def _wire(c, pts, col=NAVY, dash=None):
    c.setStrokeColor(col); c.setLineWidth(0.6)
    if dash: c.setDash(*dash)
    p = c.beginPath(); p.moveTo(*pts[0])
    for q in pts[1:]: p.lineTo(*q)
    c.drawPath(p, stroke=1, fill=0); c.setDash()

def wenner(c, x, y, w, h):
    yg = y + h * 0.52; sp = (w * 0.62) / 3; x0 = x + 12
    xs = [x0 + i * sp for i in range(4)]
    _ground(c, x0 - 8, xs[3] + 8, yg)
    for xx, lab in zip(xs, ["C1", "P1", "P2", "C2"]): _probe(c, xx, yg, 7, lab, NAVY if lab[0] == "P" else RED)
    bx, by = (xs[1] + xs[2]) / 2 - 20, y + h - 12; _box(c, bx, by, 40, 11, "Tester")
    _wire(c, [(bx, by + 5), (xs[0], by + 5), (xs[0], yg + 13)], RED)
    _wire(c, [(bx + 40, by + 5), (xs[3], by + 5), (xs[3], yg + 13)], RED)
    _wire(c, [(bx + 8, by), (xs[1], yg + 13)]); _wire(c, [(bx + 32, by), (xs[2], yg + 13)])
    for i in range(3): _dim(c, xs[i], xs[i + 1], yg - 15, "a")
    _t(c, x + 2, y + 1, "Equal spacing a · 10–15 cm deep · I via C1–C2, V across P1–P2 · ρ = 2πaR", 5.2, True, NAVY)
    cx, cy, r = x + w * 0.89, y + h * 0.55, min(h * 0.28, 15)
    _wire(c, [(cx - r, cy), (cx + r, cy)], NAVY); _wire(c, [(cx, cy - r), (cx, cy + r)], NAVY)
    _wire(c, [(cx - r * .7, cy - r * .7), (cx + r * .7, cy + r * .7)], GREY, (1, 1.5))
    _wire(c, [(cx - r * .7, cy + r * .7), (cx + r * .7, cy - r * .7)], GREY, (1, 1.5))
    _t(c, cx + r + 2, cy - 2, "D1", 5.5, True, NAVY); _t(c, cx, cy + r + 2, "D2", 5.5, True, NAVY, "c")
    _t(c, cx, cy - r - 7, "D3/D4 dashed", 5, False, GREY, "c")

def fop(c, x, y, w, h, grid=False):
    yg = y + h * 0.58; xe = x + 18; xc = x + w - 12; D = xc - xe
    _ground(c, x + 4, x + w - 4, yg)
    if grid:
        c.setStrokeColor(NAVY); c.setLineWidth(0.6); c.rect(xe - 12, yg - 6, 20, 4)
    else:
        c.setStrokeColor(NAVY); c.setLineWidth(1.6); c.line(xe, yg, xe, yg - 9)
        c.setLineWidth(0.6); c.setStrokeColor(RED); c.rect(xe - 4, yg + 1, 8, 4)
    _t(c, xe + (8 if grid else 7), yg + 3, "Grid L" if grid else "E (link open)", 5.2, True, RED if not grid else NAVY)
    for f, lab in ((0.5, "50%"), (0.618, "61.8%"), (0.7, "70%")):
        xp = xe + f * D
        if f == 0.618: _probe(c, xp, yg, 6, "P", NAVY)
        else:
            c.setStrokeColor(GREY); c.setDash(1, 1.2); c.line(xp, yg + 3, xp, yg - 6); c.setDash()
        _t(c, xp, yg - 11, lab, 5.2, f == 0.618, NAVY if f == 0.618 else black, "c")
    _probe(c, xc, yg, 7, "C", RED)
    bx, by = xe + D * 0.20, y + h - 12; _box(c, bx, by, 36, 11, "Tester")
    _wire(c, [(bx, by + 5), (xe, by + 5), (xe, yg + 6)]); _wire(c, [(bx + 30, by), (xe + .618 * D, yg + 13)])
    _wire(c, [(bx + 36, by + 5), (xc, by + 5), (xc, yg + 13)], RED)
    _dim(c, xe, xc, yg - 20, "D per Table 17 (L 5 / 10 / 20 m → D 100 / 150 / 200 m)" if grid else "D (E to C) ≥ 30 m for rod / pipe / plate – Cl. 41.1.5", 5)
    _t(c, x + 2, y + 1, "Read at 61.8%, confirm at 50% and 70% – values must agree (IS 3043:2018 Cl. 41.1.6)", 5.2, True, NAVY)

def continuity(c, x, y, w, h):
    yb = y + h * 0.42
    _box(c, x + 6, yb - 6, 44, 12, "Earth bar / ref.")
    c.setStrokeColor(NAVY); c.setLineWidth(1.2); c.line(x + 50, yb, x + w - 58, yb)
    _t(c, (x + w) / 2 - 10, yb + 2, "protective / bonding conductor, length l", 5.2, False, black, "c")
    _box(c, x + w - 58, yb - 10, 50, 20, "Equipment"); _t(c, x + w - 33, yb - 17, "frame / body", 5)
    bx, by = (x + w) / 2 - 22, y + h - 15; _box(c, bx, by, 44, 12, "Low-Ω tester")
    _wire(c, [(bx, by + 4), (x + 28, by + 4), (x + 28, yb + 6)]); _wire(c, [(bx + 44, by + 4), (x + w - 33, by + 4), (x + w - 33, yb + 10)])
    _t(c, x + 4, y + 9, "1. Short the leads and null (record lead R)  2. Clean paint/rust at both points", 5.2)
    _t(c, x + 4, y + 1, "3. Read R between reference and tested point (IEEE 81-2012 Cl. 10)", 5.2, True, NAVY)

def loop(c, x, y, w, h, pfc=False):
    top, bot = y + h - 8, y + 13
    _box(c, x + 4, (top + bot) / 2 - 9, 30, 18, "Tx")
    c.setStrokeColor(NAVY); c.setLineWidth(0.9)
    c.line(x + 34, top - 4, x + w * 0.62, top - 4); _t(c, x + 40, top - 1, "L (phase)", 5.2, True, RED)
    c.line(x + 34, (top + bot) / 2, x + w * 0.62, (top + bot) / 2); _t(c, x + 40, (top + bot) / 2 + 2, "N", 5.2, True, NAVY)
    c.setStrokeColor(HexColor("#2E7D32")); c.line(x + 34, bot + 4, x + w * 0.62, bot + 4); _t(c, x + 40, bot + 6, "PE / earth", 5.2, True, HexColor("#2E7D32"))
    dx = x + w * 0.62; _box(c, dx, bot - 2, 30, top - bot, "Device")
    tx = x + w * 0.80; _box(c, tx, (top + bot) / 2 - 8, 42, 16, "4118A")
    _wire(c, [(dx + 30, top - 4), (tx + 6, top - 4), (tx + 6, (top + bot) / 2 + 8)], RED)
    if pfc: _wire(c, [(dx + 30, (top + bot) / 2), (tx, (top + bot) / 2)], NAVY)
    _wire(c, [(dx + 30, bot + 4), (tx + 36, bot + 4), (tx + 36, (top + bot) / 2 - 8)], HexColor("#2E7D32"))
    if pfc:
        _t(c, x + 4, y + 1, "L-PE → PFC · L-N → PSC · device breaking capacity (kA) must exceed the higher", 5.2, True, NAVY)
    else:
        c.setStrokeColor(RED); c.setDash(2, 1.5); c.setLineWidth(0.6)
        c.rect(x + 22, bot + 1, dx - x - 6, top - bot - 2); c.setDash()
        _t(c, x + 4, y + 1, "Zs = whole fault loop: phase → fault → PE → source (IS 3043:2018 Cl. 42)", 5.2, True, NAVY)

def riser(c, x, y, w, h):
    yg = y + h * 0.46; _ground(c, x + 4, x + w - 4, yg)
    c.setStrokeColor(NAVY); c.setLineWidth(1.2); c.line(x + 10, yg - 7, x + w - 10, yg - 7)
    _t(c, x + w - 12, yg - 14, "buried earth grid", 5.2, False, NAVY, "r")
    xa, xr = x + w * 0.2, x + w * 0.62
    for xx, lab in ((xa, "Reference riser"), (xr, "Riser under test")):
        c.setStrokeColor(NAVY); c.setLineWidth(1.2); c.line(xx, yg - 7, xx, y + h - 10); _t(c, xx, y + h - 7, lab, 5.2, True, NAVY, "c")
    _box(c, xr + 14, y + h - 24, 38, 12, "Equipment"); c.setLineWidth(1.2); c.line(xr, y + h - 18, xr + 14, y + h - 18)
    mp = (xr, yg + 8); c.setFillColor(RED); c.circle(*mp, 2, fill=1)
    _wire(c, [mp, (xr + 33, y + h - 24)], RED, (2, 1.5)); _t(c, xr + 38, yg + 8, "(1) towards equipment", 5.3, True, RED)
    _wire(c, [mp, (xa, yg + 8)], RED, (2, 1.5)); _t(c, (xa + xr) / 2, yg + 11, "(2) towards grid", 5.3, True, RED, "c")
    _t(c, x + 2, y + 1, "Null leads first · same reference riser for the area · compare risers (IEEE 81-2012 Cl. 10)", 5.2, True, NAVY)

def step_touch(c, x, y, w, h):
    yg = y + h * 0.56; _ground(c, x + 4, x + w - 4, yg)
    sx = x + 18; c.setStrokeColor(NAVY); c.setLineWidth(1.6); c.line(sx, yg, sx, y + h - 2)
    _t(c, sx - 3, y + h - 8, "structure", 5, True, NAVY, "r") if sx - 3 > x + 30 else _t(c, sx + 3, yg + 3, "structure", 5, True, NAVY)
    pt = sx + w * 0.15; _probe(c, pt, yg, 5, "", NAVY)
    _box(c, sx + 10, y + h - 13, 30, 11, "V touch")
    _wire(c, [(sx, y + h - 8), (sx + 10, y + h - 8)]); _wire(c, [(sx + 40, y + h - 8), (pt, y + h - 8), (pt, yg + 4)])
    _dim(c, sx, pt, yg - 10, "≤ 1 m", 5)
    p1, p2 = x + w * 0.42, x + w * 0.42 + w * 0.15
    _probe(c, p1, yg, 5, "", NAVY); _probe(c, p2, yg, 5, "", NAVY)
    _box(c, (p1 + p2) / 2 - 15, y + h - 13, 30, 11, "V step")
    _wire(c, [(p1, yg + 4), (p1, y + h - 8), ((p1 + p2) / 2 - 15, y + h - 8)]); _wire(c, [(p2, yg + 4), (p2, y + h - 8), ((p1 + p2) / 2 + 15, y + h - 8)])
    _dim(c, p1, p2, yg - 10, "1 m", 5)
    xc = x + w - 12; _probe(c, xc, yg, 7, "C", RED)
    _t(c, xc - 7, yg + 3, "remote current electrode", 5, False, RED, "r")
    _t(c, x + 2, y + 1, "Record It per reading · V actual = V measured × If / It → compare with tolerable limit (IEEE 81-2012 Cl. 9.3)", 5.1, True, NAVY)

def tower(c, x, y, w, h):
    cx, cy, s = x + w * 0.24, y + h * 0.55, min(h * 0.62, 34)
    c.setStrokeColor(NAVY); c.setLineWidth(0.8); c.rect(cx - s / 2, cy - s / 2, s, s)
    c.line(cx - s / 2, cy - s / 2, cx + s / 2, cy + s / 2); c.line(cx - s / 2, cy + s / 2, cx + s / 2, cy - s / 2)
    for i, (dx, dy) in enumerate(((-1, 1), (1, 1), (1, -1), (-1, -1)), 1):
        fx, fy = cx + dx * s / 2, cy + dy * s / 2
        c.setFillColor(NAVY); c.circle(fx, fy, 2.2, fill=1)
        c.setStrokeColor(RED); c.setLineWidth(0.9); c.circle(fx, fy, 5, fill=0)
        _t(c, fx + dx * 9, fy + (3 if dy > 0 else -7), f"F-{i}", 5.5, True, NAVY, "c")
    _t(c, cx, y + 2, "plan view – clamp (red) on each foot earth connection", 5, False, RED, "c")
    xp, xc = x + w * 0.62, x + w - 12
    c.setStrokeColor(GREY); c.setLineWidth(0.5); c.setDash(2, 1.5); c.line(cx + s / 2, cy, xc, cy); c.setDash()
    _probe(c, xp, cy, 0, "P", NAVY); _probe(c, xc, cy, 0, "C", RED)
    c.setFillColor(NAVY); c.circle(xp, cy, 2, fill=1); c.setFillColor(RED); c.circle(xc, cy, 2, fill=1)
    _t(c, (xp + xc) / 2, cy - 10, "auxiliary probes along a straight line", 5, False, black, "c")
    _t(c, x + w * 0.5, y + 10, "Earth wire stays connected · one reading per foot", 5.2)
    _t(c, x + w * 0.5, y + 2, "Zt = Σ 4 feet vs Zsat (acceptance list)", 5.2, True, NAVY)
