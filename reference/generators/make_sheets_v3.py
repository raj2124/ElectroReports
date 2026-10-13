"""ETPL ESHA measurement sheets v2 (MS-01..MS-09).
Built from ETPL's existing sheets, IS 3043:2018, IEEE 81-2012 and ETPL's instrument list."""
import io, qrcode
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor, black, white
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
import diagrams as DG

pdfmetrics.registerFont(TTFont("Arial", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"))
NAVY = HexColor("#0B2A67"); GREY = HexColor("#9AA3B2"); TINT = HexColor("#EEF1F6"); BOX = HexColor("#5F6878")
LOGO = ImageReader("/home/claude/esha/logo.png")
M = 26
PROJECT = dict(client="Gujarat Fluorochemicals Ltd", site="Dahej – FI Plant", project="024-D03_TS-26-GFL",
               wo="1002-1261201769")
INSTR = {"3290": "Metrel MI 3290 GX Earth Analyser", "3152": "Metrel MI 3152 EurotestXC",
         "4105": "Kyoritsu 4105A Digital Earth Tester", "4118": "Kyoritsu 4118A Loop / PSC Tester",
         "3295": "Metrel MI 3295 Step & Contact Voltage Meter*"}
CODE_LEGEND = ("C1 Corroded · C2 Strip / test link disconnected · C3 Single earth connection (optional) · "
               "C4 Buried / chamber not accessible or damaged · C5 Electrode damaged · C6 Loose connection · "
               "C7 Joint not cleaned / tinned / protected · C8 No tag / label · OTH Other – write remark")

class Sheet:
    def __init__(s, c, land):
        s.c = c; s.W, s.H = landscape(A4) if land else A4; s.CW = s.W - 2 * M
        c.setPageSize((s.W, s.H))

    def qr(s, data):
        img = qrcode.make(data, box_size=4, border=1); b = io.BytesIO(); img.save(b, "PNG"); b.seek(0)
        return ImageReader(b)

    def header(s, docno, title, test_ref, instruments, extra=()):
        c, W, H, CW = s.c, s.W, s.H, s.CW
        c.drawImage(LOGO, M, H - M - 17, 17, 17, mask="auto")
        c.setFont("Arial-Bold", 9); c.setFillColor(NAVY); c.drawString(M + 22, H - M - 12, "Elegrow Technology Pvt. Ltd.")
        c.setFont("Arial", 8); c.drawRightString(W - M, H - M - 12, f"{docno} | Rev 00")
        c.setStrokeColor(NAVY); c.setLineWidth(1.1); c.line(M, H - M - 21, W - M, H - M - 21)
        # title band
        y = H - M - 25
        fits = pdfmetrics.stringWidth(title, "Arial-Bold", 12) + pdfmetrics.stringWidth(test_ref, "Arial", 7.5) + 30 < CW
        bh = 22 if fits else 30
        c.setFillColor(NAVY); c.rect(M, y - bh, CW, bh, fill=1, stroke=0)
        c.setFillColor(white); c.setFont("Arial-Bold", 12); c.drawString(M + 8, y - 15, title)
        c.setFont("Arial", 7.5)
        if fits: c.drawRightString(W - M - 8, y - 15, test_ref)
        else: c.drawString(M + 8, y - 26, test_ref)
        y -= bh - 22
        # info grid (left) + QR (right)
        y -= 26; rh = 17
        rows = [[("Client", PROJECT["client"], .42), ("Project No.", PROJECT["project"], .33), ("Work Order No.", PROJECT["wo"], .25)],
                [("Site / Unit", PROJECT["site"], .42), ("Work date", "____ / ____ / 20____", .33), ("Sheet", "____ of ____", .25)],
                [("ETPL engineer(s)", "", .42), ("Client representative", "", .33), ("Ground condition", "☐ Dry  ☐ Damp  ☐ Wet", .25)]]
        rows += list(extra)
        qw = min(66, len(rows) * rh); gw = CW - qw - 6
        for r, row in enumerate(rows):
            x = M
            for label, val, frac in row:
                w = gw * frac; c.setStrokeColor(GREY); c.setLineWidth(0.5); c.rect(x, y - (r + 1) * rh, w, rh)
                c.setFillColor(NAVY); c.setFont("Arial-Bold", 6.3); c.drawString(x + 3, y - r * rh - 6.8, label)
                c.setFillColor(black); c.setFont("Arial", 8); c.drawString(x + 3, y - r * rh - 14.6, val)
                x += w
        gy = y - len(rows) * rh
        c.drawImage(s.qr(f"ER|{PROJECT['project']}|{docno}"), W - M - qw, y - qw + 1, qw, qw)
        # instrument row
        iy = gy - 4; ih = 16
        c.setStrokeColor(GREY); c.rect(M, iy - ih, CW, ih)
        c.setFillColor(NAVY); c.setFont("Arial-Bold", 6.5); c.drawString(M + 3, iy - 10.5, "Instrument used:")
        x = M + 66
        for k in instruments:
            c.setStrokeColor(BOX); c.rect(x, iy - 11.5, 7, 7)
            c.setFillColor(black); c.setFont("Arial", 7); c.drawString(x + 10, iy - 10.5, INSTR[k]); x += 14 + pdfmetrics.stringWidth(INSTR[k], "Arial", 7) + 8
        c.setFont("Arial", 7); c.drawString(x + 4, iy - 10.5, "Sr. No. __________   Cal. due ___/___/____")
        return iy - ih - 6

    def footer(s, docno, title, legend=True):
        c, W, CW = s.c, s.W, s.CW
        y = M + 44
        if legend:
            parts = CODE_LEGEND.split(" · "); c.setFont("Arial-Bold", 6); c.setFillColor(NAVY); c.drawString(M, y + 29, "Condition codes:")
            c.setFont("Arial", 6); c.setFillColor(black)
            c.drawString(M + 52, y + 29, " · ".join(parts[:5])); c.drawString(M + 52, y + 21.5, " · ".join(parts[5:]))
        for i, lab in enumerate(["ETPL engineer – name & signature", "Client representative – name & signature", "Time start / finish"]):
            w = CW * (.38 if i < 2 else .24); x = M + CW * (0, .38, .76)[i]
            c.setStrokeColor(GREY); c.setLineWidth(0.5); c.rect(x, y - 6, w, 26)
            c.setFillColor(NAVY); c.setFont("Arial-Bold", 6.3); c.drawString(x + 3, y + 13, lab)
        c.setStrokeColor(GREY); c.line(M, M + 14, s.W - M, M + 14)
        c.setFont("Arial-Bold", 7); c.setFillColor(NAVY); c.drawString(M, M + 4, "INTERNAL")
        c.setFont("Arial", 7); c.drawRightString(s.W - M, M + 4, f"{title} | {docno}")
        return y + (38 if legend else 26)

    def digits(s, x, y, w, h, pat):
        pass  # plain writing cell (no digit boxes)

    def ticks(s, x, y, w, h, opts, labels=True):
        c = s.c; sz = 6.6; fs = 6
        item = lambda o: sz + 2 + pdfmetrics.stringWidth(o, "Arial", fs)
        lines = [opts]
        if sum(item(o) for o in opts) + 4 * (len(opts) - 1) > w - 6 and len(opts) > 1:
            k = (len(opts) + 1) // 2; lines = [opts[:k], opts[k:]]
        lh = sz + 3.2; y_top = y + h / 2 + (len(lines) * lh) / 2 - lh
        for li, line in enumerate(lines):
            tot = sum(item(o) for o in line) + 4 * (len(line) - 1); xx = x + max(3, (w - tot) / 2); yy = y_top - li * lh + 1.6
            for o in line:
                c.setStrokeColor(BOX); c.setLineWidth(0.5); c.rect(xx, yy, sz, sz)
                c.setFont("Arial", fs); c.setFillColor(black); c.drawString(xx + sz + 2, yy + 1.2, o)
                xx += item(o) + 4

    def table(s, top, cols, n_rows, rh, group=1, prefill=None, bottom_limit=None):
        """cols: dict(h=header, w=fraction, k=kind, a=arg, merge=bool). group: rows per merged block."""
        c, CW = s.c, s.CW
        tot = sum(col["w"] for col in cols); ws = [col["w"] / tot * CW for col in cols]
        hh = 30; x = M
        for col, w in zip(cols, ws):
            c.setFillColor(TINT); c.rect(x, top - hh, w, hh, fill=1, stroke=0)
            c.setStrokeColor(GREY); c.setLineWidth(0.5); c.rect(x, top - hh, w, hh)
            c.setFillColor(NAVY); c.setFont("Arial-Bold", 6.5); lines = col["h"].split("\n")
            for i, ln in enumerate(lines):
                c.drawCentredString(x + w / 2, top - hh / 2 + (len(lines) - 1) * 3.9 - i * 7.8 - 2, ln)
            x += w
        y0 = top - hh
        for r in range(n_rows):
            y = y0 - (r + 1) * rh; x = M; first = r % group == 0
            pre = prefill[r] if prefill and r < len(prefill) else {}
            for col, w in zip(cols, ws):
                k, a = col["k"], col.get("a")
                if col.get("merge"):
                    if first:
                        c.setStrokeColor(GREY); c.rect(x, y0 - (r + group) * rh, w, group * rh)
                        if k == "sr": c.setFont("Arial", 7.5); c.setFillColor(black); c.drawCentredString(x + w / 2, y0 - (r + group / 2) * rh - 3, str(r // group + 1))
                        if k == "digits": s.digits(x, y0 - (r + group / 2) * rh - rh / 2, w, rh, a)
                        if k in ("ticks", "ticksnl"): s.ticks(x, y0 - (r + group) * rh, w, group * rh, a)
                        if k == "text" and col["h"] in pre:
                            c.setFont("Arial", 7); c.setFillColor(black); c.drawString(x + 3, y0 - (r + group / 2) * rh - 3, pre[col["h"]])
                    x += w; continue
                c.setStrokeColor(GREY); c.rect(x, y, w, rh)
                if k == "sr": c.setFont("Arial", 7.5); c.setFillColor(black); c.drawCentredString(x + w / 2, y + rh / 2 - 3, str(r + 1))
                elif k == "fixed": c.setFont("Arial", 7.5); c.setFillColor(black); c.drawCentredString(x + w / 2, y + rh / 2 - 3, a[r % len(a)])
                elif k == "text" and col["h"] in pre: c.setFont("Arial", 7); c.setFillColor(black); c.drawString(x + 3, y + rh / 2 - 3, pre[col["h"]])
                elif k == "digits": s.digits(x, y, w, rh, a)
                elif k in ("ticks", "ticksnl"): s.ticks(x, y, w, rh, a)
                x += w
        end = y0 - n_rows * rh
        if bottom_limit: assert end > bottom_limit + 12, f"overflow by {bottom_limit + 12 - end:.0f}pt"
        return end

    def band(s, y_top, fig, lines, bh):
        c = s.c; CW = s.CW; fw = CW * (0.50 if s.W < 700 else 0.40)
        c.setStrokeColor(GREY); c.setLineWidth(0.5); c.rect(M, y_top - bh, fw, bh)
        c.setFillColor(NAVY); c.setFont("Arial-Bold", 6); c.drawString(M + 3, y_top - 7, "SET-UP")
        fig(c, M + 4, y_top - bh + 2, fw - 8, bh - 10)
        tx = M + fw + 8; tw = CW - fw - 8; ty = y_top - 7
        c.setFont("Arial-Bold", 6.3); c.setFillColor(NAVY); c.drawString(tx, ty, "HOW TO RECORD"); ty -= 8.5
        c.setFont("Arial", 6.3); c.setFillColor(black)
        for ln in lines:
            words = ln.split(); cur = ""; first = True
            for wd in words:
                trial = (cur + " " + wd).strip()
                if pdfmetrics.stringWidth(("• " if first else "  ") + trial, "Arial", 6.3) > tw:
                    c.drawString(tx, ty, ("• " if first else "  ") + cur); ty -= 7.8; cur = wd; first = False
                else: cur = trial
            c.drawString(tx, ty, ("• " if first else "  ") + cur); ty -= 8.6
        assert ty > y_top - bh - 4, "instructions overflow band"

BAND = 74
ROWS = {}
def make(c, land, docno, title, ref, instr, cols, n_rows, rh, notes=(), group=1, prefill=None, extra=(), legend=True,
         fill=True, fig=None, min_rows=0):
    s = Sheet(c, land); top = s.header(docno, title, ref, instr, extra)
    fb = s.footer(docno, title.split(" – ")[0], legend)
    bh = BAND if not land else 61
    space = top - 30 - (fb + bh + 10)
    n_rows = max(n_rows, min_rows)
    fit = int(space // rh) // group * group
    if fill and fit > n_rows: n_rows = fit
    rh = min(space / n_rows, 34)
    assert rh >= 17.5, f"{title}: rows too small ({rh:.1f}pt)"
    ROWS[docno] = (n_rows, round(rh, 1))
    end = s.table(top, cols, n_rows, rh, group, prefill)
    s.band(end - 6, fig, notes, bh); c.showPage()

col = lambda h, w, k, a=None, merge=False: dict(h=h, w=w, k=k, a=a, merge=merge)
CODES = ["C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8"]
c = canvas.Canvas("/mnt/user-data/outputs/ESHA_Measurement_Sheets_v3_MS-01_to_MS-09.pdf")
c.setTitle("ETPL ESHA Measurement Sheets v3 (MS-01 to MS-09)"); c.setAuthor("Elegrow Technology Pvt. Ltd.")
SP = ["0.50", "1.00", "1.50", "2.00", "2.50", "3.00", "3.50", "4.00", "4.50", "5.00"]
V2 = {"001": 10, "002": 10, "003": 19, "004": 12, "005": 12, "006": 19, "007": 18, "008": 12, "009": 20}

make(c, False, "ET-D03-FRM-001", "MS-01  Soil Resistivity – Measurement Sheet", "Wenner 4-probe · IS 3043:2018 Cl. 40 · IEEE 81-2012 Cl. 7",
     ["3290", "4105"],
     [col("Sr.", 4, "sr"), col("Probe spacing\na (m)", 10, "fixed", SP),
      col("Direction 1\nreading", 14, "digits"), col("Direction 2\nreading", 14, "digits"),
      col("Direction 3\nreading", 14, "digits"), col("Direction 4\nreading", 14, "digits"), col("Remark", 18, "text")], 10, 30,
     notes=["Write the value shown on the instrument and tick its unit in the header: R (Ω) or ρ (Ω·m). The app calculates ρ = 2πaR, mean per spacing, average and IS 3043 Table 4 corrosivity.",
            "Four probes in a straight line at equal spacing a, driven 10–15 cm. If depth exceeds a/20 the app uses the full formula (Cl. 40.4.1).",
            "Substations / generating stations: at least 8 directions from the centre (Cl. 40.2.1) – use extra sheets. Keep the line clear of buried metal and earth conductors."],
     extra=[[("Location name", "", .42), ("Reading unit", "☐ R (Ω)   ☐ ρ (Ω·m)", .33), ("Probe depth (cm)", "", .25)],
            [("Direction 1 / 2 bearing", "D1 _____°   D2 _____°", .42), ("Direction 3 / 4 bearing", "D3 _____°   D4 _____°", .33), ("Weather", "☐ Sunny ☐ Cloudy ☐ After rain", .25)]],
     legend=False, fill=False, fig=DG.wenner)

TAGS = [("GFL-FI-EP-%02d" % i, l) for i, l in enumerate(["Transformer TR-1", "Transformer TR-1", "MCC Room", "MCC Room", "Reactor Bay R-201",
         "Reactor Bay R-202", "Cooling Tower", "Pump House", "Pipe Rack PR-3", "DG Area"], 1)]
make(c, True, "ET-D03-FRM-002", "MS-02  Earth Electrode (Pit) Resistance – Measurement Sheet",
     "Fall of potential 61.8% · IS 3043:2018 Cl. 14.2, 41 · IEEE 81-2012 Cl. 8", ["3290", "4105"],
     [col("Sr.", 2.3, "sr"), col("Pit TAG No.", 7, "text"), col("Location / name of pit", 9, "text"),
      col("Electrode\ntype", 8, "ticks", ["Plate", "Pipe", "Rod", "Strip"]),
      col("Material", 12.5, "ticks", ["Cu", "Cu-bonded", "GI", "MS", "CI", "SS"]),
      col("Length l\n(m)", 5, "digits"), col("Dia d / strip\nwidth (mm)", 6, "digits"),
      col("R without grid\nat 61.8% (Ω)", 7, "digits"), col("Check R\nat 50% (Ω)", 6, "digits"), col("Check R\nat 70% (Ω)", 6, "digits"),
      col("R with\ngrid (Ω)", 6, "digits"), col("Method", 8.5, "ticks", ["61.8%", "Clamp", "Slope"]),
      col("Condition codes", 11.5, "ticks", CODES), col("Photo\nref.", 3.6, "text"), col("Remark", 6, "text")], 10, 27,
     prefill=[{"Pit TAG No.": t, "Location / name of pit": l} for t, l in TAGS], min_rows=V2["002"],
     notes=["Open the test link for 'R without grid'; restore it after. Read at 61.8%, then move P to 50% and 70%: if the three readings agree the 61.8% value is valid (Cl. 41.1.6).",
            "Current electrode C at least 30 m from a rod, pipe or plate (Cl. 41.1.5); auxiliary rods 12.5 mm MS, up to 1 m deep (Cl. 41.1.4). Record C distance in the header.",
            "Plate: write plate size (m × m) in Remark. Pre-printed tags come from the asset register; new pits on a continuation sheet."],
     extra=[[("Current electrode distance C (m)", "", .42), ("Auxiliary rods", "☐ 12.5 mm MS, ≤ 1 m deep", .33), ("90°/180° check done", "☐ Yes  ☐ No", .25)]],
     fig=DG.fop)

make(c, False, "ET-D03-FRM-003", "MS-03  Earth Continuity & Bonding – Measurement Sheet", "Ground integrity · IEEE 81-2012 Cl. 10 · IS 3043:2018 Cl. 17.2.3",
     ["3152", "3290"],
     [col("Sr.", 3.5, "sr"), col("TAG No.", 8, "text"), col("Main location /\narea", 12, "text"),
      col("Reference point\n(earth bar / riser)", 13, "text"), col("Tested point\n(equipment / riser)", 13, "text"),
      col("Distance\n(m)", 7, "digits"), col("Test current\n(A)", 7.5, "digits"), col("Measured\nR (Ω)", 8.5, "digits"),
      col("Condition\ncodes", 10.5, "ticks", ["C1", "C2", "C6", "C7"]), col("Photo\nref.", 5, "text"), col("Remark", 8, "text")], 19, 26,
     min_rows=V2["003"],
     notes=["Short the leads and null the tester; write lead R in the header – it is subtracted from every reading (IEEE 81-2012 Cl. 10.2).",
            "Keep one reference point for the area so results can be compared point to point. Record distance between the two points where known.",
            "High-current integrity test: write test current and use Remark for the voltage drop (V)."],
     extra=[[("Test-lead resistance (Ω)", "", .42), ("Lead null done", "☐ Yes", .33), ("Method", "☐ Low-Ω   ☐ High-current", .25)]],
     fig=DG.continuity)

PH = ["R-E", "Y-E", "B-E"]
make(c, True, "ET-D03-FRM-004", "MS-04  Earth Fault Loop Impedance – Measurement Sheet", "Zs, L-PE · IS 3043:2018 Cl. 42 · IS/IEC 60364-6", ["4118", "3152"],
     [col("Sr.", 2.6, "sr", merge=True), col("Location of panel / equipment", 14, "text", merge=True), col("Feeder name & tag no.", 13, "text", merge=True),
      col("Device\ntype", 11, "ticks", ["ACB", "MCCB", "MCB", "SFU", "Fuse"], merge=True), col("Device rating\nIn (A)", 7, "digits", merge=True),
      col("Trip\ncurve", 7, "ticks", ["B", "C", "D", "Adj."], merge=True), col("Measured\npoint", 5, "fixed", PH),
      col("Loop impedance\nZs (Ω)", 9, "digits"), col("Voltage\nL-PE (V)", 7, "digits"), col("Photo\nref.", 4.5, "text", merge=True),
      col("Remark", 11, "text", merge=True)], 12, 20, group=3, min_rows=V2["004"],
     notes=["Live working: risk assessment, PPE and permit mandatory. Measure at the outgoing side of the device, one reading per phase to PE.",
            "Copy device type, rating In and trip curve from the nameplate – the app uses them for the device's disconnection limit.",
            "A panel is healthy only if all three phases meet the project's Zs criterion (Standards Acceptance List)."],
     fig=DG.loop)

make(c, True, "ET-D03-FRM-005", "MS-05  Prospective Fault & Short-circuit Current – Measurement Sheet", "PFC (L-PE) & PSC (L-N) · IS/IEC 60364-4-43 Reg. 434.5.1", ["4118", "3152"],
     [col("Sr.", 2.6, "sr", merge=True), col("Location of panel / equipment", 14, "text", merge=True), col("Feeder name & tag no.", 13, "text", merge=True),
      col("Device\ntype", 11, "ticks", ["ACB", "MCCB", "MCB", "SFU", "Fuse"], merge=True), col("Rating\nIn (A)", 6, "digits", merge=True),
      col("Breaking capacity\nIcu / Icn (kA)", 8, "digits", merge=True), col("Phase", 4.2, "fixed", ["R", "Y", "B"]),
      col("PFC L-PE\n(kA)", 7.5, "digits"), col("PSC L-N\n(kA)", 7.5, "digits"), col("Voltage\nL-N (V)", 6, "digits"),
      col("Photo\nref.", 4.5, "text", merge=True), col("Remark", 9, "text", merge=True)], 12, 20, group=3, min_rows=V2["005"],
     notes=["Breaking capacity (kA) from the nameplate is REQUIRED – without it the result is 'Not Assessable'.",
            "Record both PFC (L-PE) and PSC (L-N) per phase; the device must interrupt the higher of the two.",
            "Write values in kA as shown on the 4118A (convert A to kA if the display shows A)."],
     fig=lambda c_, x, y, w, h: DG.loop(c_, x, y, w, h, pfc=True))

make(c, False, "ET-D03-FRM-006", "MS-06  Riser Integrity – Measurement Sheet", "Two-direction continuity · IEEE 81-2012 Cl. 10", ["3290", "3152"],
     [col("Sr.", 4, "sr"), col("Main location", 15, "text"), col("Equipment / riser /\nmeasurement point", 20, "text"),
      col("R towards\nequipment (Ω)", 12.5, "digits"), col("R towards\ngrid (Ω)", 12.5, "digits"),
      col("Condition\ncodes", 12, "ticks", ["C1", "C2", "C6", "C7"]), col("Photo\nref.", 6, "text"), col("Remark", 11, "text")], 19, 26,
     min_rows=V2["006"],
     notes=["Null the leads first; write lead R in the header (IEEE 81-2012 Cl. 10.2).",
            "Reading ① from the measurement point to the equipment body; reading ② from the same point to the grid / reference riser.",
            "Use one reference riser for the whole area so risers can be compared with each other."],
     extra=[[("Test-lead resistance (Ω)", "", .42), ("Reference riser / point", "", .33), ("Test current", "☐ 200 mA ☐ Other ____", .25)]],
     fig=DG.riser)

make(c, False, "ET-D03-FRM-007", "MS-07  Ground (Grid) Resistance – Measurement Sheet",
     "61.8% / slope / clamp · IS 3043:2018 Cl. 41, Table 17 · IEEE 81-2012 Cl. 8", ["3290", "4105"],
     [col("Sr.", 3.6, "sr"), col("Point / pit TAG", 10.5, "text"), col("Location of measured ground", 16, "text"),
      col("Method", 11, "ticks", ["61.8%", "Slope", "Clamp"]), col("Current probe\ndistance C (m)", 9, "digits"),
      col("R at 50%\n(Ω)", 9.5, "digits"), col("R at 61.8%\n(Ω)", 9.5, "digits"), col("R at 70%\n(Ω)", 9.5, "digits"),
      col("Condition\ncodes", 9.5, "ticks", ["C1", "C2", "C6", "C7"]), col("Remark", 11.5, "text")], 18, 27, min_rows=V2["007"],
     notes=["Grid stays connected (in-service value). Choose C from Table 17 by the grid's maximum dimension L.",
            "The 61.8% value is valid when the 50% and 70% readings agree; if not, move C further away or use the slope method (Cl. 41.1.8).",
            "Clamp method: no auxiliary probes – write '—' in the distance and 50/70% columns."],
     extra=[[("Max. grid dimension L (m)", "", .42), ("Table 17 distance C used (m)", "", .33), ("90°/180° check", "☐ Yes  ☐ No", .25)]],
     fig=lambda c_, x, y, w, h: DG.fop(c_, x, y, w, h, grid=True))

make(c, True, "ET-D03-FRM-008", "MS-08  Step & Touch Voltage – Measurement Sheet",
     "Current injection · IEEE 81-2012 Cl. 9 · IS 3043:2018 Cl. 10.3 · IEEE 80-2013", ["3295", "3290"],
     [col("Sr.", 2.6, "sr"), col("Location / structure / measurement point", 16, "text"),
      col("Measurement type", 15, "ticks", ["Step", "Touch-structure", "Touch-mesh", "Touch-fence", "Touch-gate"]),
      col("Surface", 9, "ticks", ["Soil", "Gravel", "Concrete"]), col("Distance from\nstructure (m)", 6.5, "digits"),
      col("Test current\nIt (A)", 6.5, "digits"), col("V1\n(V)", 6, "digits"), col("V2\n(V)", 6, "digits"),
      col("V3\n(V)", 6, "digits"), col("Vmax\n(V)", 6, "digits"), col("Photo\nref.", 4, "text"), col("Remark", 9, "text")], 12, 24,
     min_rows=V2["008"],
     notes=["Step: two probes 1 m apart. Touch: probe within 1 m of the structure (arm reach). Probes ~150 mm into moist soil (IEEE 81-2012 Cl. 9.2–9.3).",
            "Write the test current It for every row – the app scales each voltage by If / It before comparing with the tolerable limit.",
            "*MI 3295 is not on the current instrument list – confirm availability."],
     extra=[[("Earth fault current If (kA) & source", "", .42), ("Fault clearing time ts (s) & source", "", .33), ("Remote current electrode distance (m)", "", .25)],
            [("Surface layer ρs (Ω·m) & thickness hs (m)", "", .42), ("Body weight basis", "☐ 50 kg   ☐ 70 kg   ☐ Both", .33), ("Injection point", "", .25)]],
     legend=False, fig=DG.step_touch)

make(c, False, "ET-D03-FRM-009", "MS-09  Tower Footing Impedance – Measurement Sheet",
     "Clamp method · IS 3043:2018 Cl. 41 · IEEE 81-2012 Cl. 8", ["3290"],
     [col("Sr.", 3.6, "sr", merge=True), col("Tower no. /\nlocation", 14, "text", merge=True), col("Foot", 5.5, "fixed", ["F-1", "F-2", "F-3", "F-4"]),
      col("Foot-to-earth\nconnection", 11, "ticks", ["Given", "Not given"]), col("Measured current\nI (mA)", 11, "digits"),
      col("Measured impedance\nZ (Ω)", 11, "digits"), col("Condition\ncodes", 10.5, "ticks", ["C1", "C2", "C5", "C6"]),
      col("Photo\nref.", 5.5, "text"), col("Remark", 11, "text")], 20, 21, group=4, min_rows=V2["009"],
     notes=["Four feet per tower are one unit. Clamp each foot's earth connection in turn; the overhead earth wire stays connected.",
            "Mark 'Not given' where a foot has no earth connection – that foot is reported as a defect.",
            "The app calculates Zt from the four feet and compares it with Zsat from the project's acceptance profile."],
     extra=[[("Line / circuit name", "", .42), ("Line voltage (kV) / tower type", "", .33), ("Zsat (Ω) from acceptance list", "", .25)]],
     fig=DG.tower)
c.save()
for k, (n, rh) in ROWS.items():
    print(f"{k}: v2 {V2[k[-3:]]} rows -> v3 {n} rows @ {rh}pt")
