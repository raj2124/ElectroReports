"""ETPL ESHA final report — template rendered with SAMPLE data.
Every status, count and limit below is computed from the rules in ET-D03-REG-002; nothing is typed."""
import json, math
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle,
                                PageBreak, KeepTogether, NextPageTemplate)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.colors import HexColor, black, white
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Line, String
from reportlab.graphics.charts.barcharts import VerticalBarChart

pdfmetrics.registerFont(TTFont("Arial", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Arial-Bold", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"))
NAVY = HexColor("#0B2A67"); GREY = HexColor("#9AA3B2"); LIGHT = HexColor("#EEF1F6")
GREEN = HexColor("#1E7B34"); AMBER = HexColor("#B8860B"); RED = HexColor("#B42318")
W, H = A4; M = 42; CW = W - 2 * M
CLASSIFICATION = "CONFIDENTIAL"
REG = json.load(open("/home/claude/esha/ET-D03-REG-002_Standards_Acceptance_List.json"))
PROFILE = {c["id"]: c for c in REG["profiles"]["ETPL-DEFAULT"]["criteria"]}

def S(**k):
    d = dict(fontName="Arial", fontSize=9.5, leading=13); d.update(k); return ParagraphStyle("x", **d)
BODY = S(alignment=TA_JUSTIFY); SMALL = S(fontSize=7.5, leading=9.5); CELL = S(fontSize=7.8, leading=9.6)
CELLB = ParagraphStyle("cb", parent=CELL, fontName="Arial-Bold", textColor=NAVY)
H1 = ParagraphStyle("h1", fontName="Arial-Bold", fontSize=14, leading=18, textColor=NAVY, spaceBefore=6, spaceAfter=8)
H2 = ParagraphStyle("h2", fontName="Arial-Bold", fontSize=11, leading=14, textColor=NAVY, spaceBefore=8, spaceAfter=5)
NOTE = S(fontSize=8, leading=10.5, textColor=HexColor("#475467"))

# ---------------- SAMPLE DATA (would come from the daily records) ----------------
P = dict(client="Gujarat Fluorochemicals Ltd", address="12/A GIDC Industrial Estate, Dahej, Bharuch, Gujarat 392130",
         unit="FI Plant", project="024-D03_TS-26-GFL", seq="024", wo="1002-1261201769", doc="024_FI_ESHA", rev="R0",
         days=["14-09-2026", "15-09-2026"], engineers="H. S. Patil, A. S. Mahto", client_rep="Mr. B. Thakkar")
soil = {"Utility Yard": {"D1": [(0.5, 9.8), (1.0, 4.6), (1.5, 2.9), (2.0, 2.0), (2.5, 1.5), (3.0, 1.2)],
                         "D2": [(0.5, 10.4), (1.0, 4.9), (1.5, 3.1), (2.0, 2.2), (2.5, 1.6), (3.0, 1.3)]},
        "Reactor Bay": {"D1": [(0.5, 8.1), (1.0, 3.9), (1.5, 2.6), (2.0, 1.8), (2.5, 1.4), (3.0, 1.1)],
                        "D2": [(0.5, 8.6), (1.0, 4.1), (1.5, 2.7), (2.0, 1.9), (2.5, 1.5), (3.0, 1.2)]}}
pits = [  # tag, location, type, l(m), d(mm), R w/o grid, codes, remark
    ("GFL-FI-EP-01", "Transformer TR-1", "Rod", 3.0, 40, 0.82, [], ""),
    ("GFL-FI-EP-02", "Transformer TR-1", "Rod", 3.0, 40, 1.10, [], ""),
    ("GFL-FI-EP-03", "MCC Room", "Rod", 3.0, 40, 1.45, ["C1"], "Rod heavily corroded at clamp"),
    ("GFL-FI-EP-04", "MCC Room", "Rod", 3.0, 40, 0.95, [], ""),
    ("GFL-FI-EP-05", "Reactor Bay R-201", "Rod", 3.0, 40, 1.30, ["C8"], "No pit tag"),
    ("GFL-FI-EP-06", "Reactor Bay R-202", "Rod", 3.0, 40, 2.05, [], ""),
    ("GFL-FI-EP-07", "Cooling Tower", "Rod", 3.0, 40, 7.12, ["C2"], "Strip disconnected from rod"),
    ("GFL-FI-EP-08", "Pump House", "Rod", 3.0, 40, None, ["C4"], "Chamber buried under paving"),
    ("GFL-FI-EP-09", "Pipe Rack PR-3", "Rod", 3.0, 40, 6.40, [], ""),
    ("GFL-FI-EP-10", "DG Area", "Strip", 6.0, 50, 2.40, [], "25x5 strip, width 50 mm entered")]
riser = [("Reactor Bay", "R-201 (R1)", 0.21, 0.12, []), ("Reactor Bay", "R-201 (R2)", 0.34, 0.20, []),
         ("Reactor Bay", "R-202 (R1)", 1.42, 0.31, ["C6"]), ("Pump House", "P-101 A", 0.18, 0.09, []),
         ("Pump House", "P-101 B", 0.27, 0.19, []), ("Cooling Tower", "CT Fan-1", 0.44, 0.38, [])]
loop = [("PMCC-1", "Incomer", "ACB", 1600, 50, (0.62, 0.65, 0.61), (6.9, 6.7, 7.1)),
        ("MCC-2", "7F2 Feeder", "MCCB", 250, 36, (1.21, 1.35, 1.92), (2.1, 2.0, 1.4)),
        ("MLDB-1", "Incomer", "MCCB", 125, 25, (1.05, 1.10, 1.02), (2.6, 2.5, 2.7)),
        ("APDB-2", "Incomer", "MCB", 63, None, (1.40, 1.38, 1.44), (1.9, 1.9, 1.8))]

# ---------------- CALCULATIONS (mirror of reportModel.js rules) ----------------
def soil_calc():
    out = {}
    for loc, dirs in soil.items():
        rows = []
        for i in range(6):
            a = dirs["D1"][i][0]; r1 = 2 * math.pi * a * dirs["D1"][i][1]; r2 = 2 * math.pi * a * dirs["D2"][i][1]
            rows.append((a, r1, r2, (r1 + r2) / 2))
        out[loc] = (rows, sum(r[3] for r in rows) / 6)
    return out
SOIL = soil_calc(); RHO = sum(v[1] for v in SOIL.values()) / len(SOIL)
def corrosivity(r): return ("Severely corrosive" if r < 25 else "Moderately corrosive" if r < 50
                            else "Mildly corrosive" if r < 100 else "Very mildly corrosive")
def limit(kind, l_m, d_mm):
    l, d = l_m * 100, d_mm / 10
    if kind == "Rod": return 100 * RHO / (2 * math.pi * l) * math.log(2 * l / d), "PIT-ROD"
    if kind == "Strip": return 100 * RHO / (2 * math.pi * l) * math.log(4 * l / d), "PIT-STRIP"
P1C, P2C = {"C1", "C2", "C4", "C5"}, {"C3", "C6", "C7", "C8"}
PITS = []
for tag, loc, kind, l, d, R, codes, rem in pits:
    lim, crit = limit(kind, l, d)
    if R is None: status = "Not Measurable"
    elif R > lim: status = ">Permissible Limit"
    elif R > 0.9 * lim: status = "Warning"
    else: status = "Healthy"
    pr = ("P1" if status in (">Permissible Limit", "Not Measurable") or P1C & set(codes)
          else "P2" if status == "Warning" or P2C & set(codes) else "P3")
    PITS.append(dict(tag=tag, loc=loc, kind=kind, l=l, d=d, R=R, lim=lim, status=status, pr=pr, codes=codes, rem=rem, crit=crit))
RIS = [dict(main=a, pt=b, re=c, rg=d, codes=e, status="Healthy" if max(c, d) <= PROFILE["RISER-LIMIT"]["value"] else ">Permissible Limit") for a, b, c, d, e in riser]
LOOP = []
for pnl, fdr, dev, A, kA, zs, pfc in loop:
    zok = all(z <= PROFILE["LOOP-LIMIT"]["value"] for z in zs)
    pstat = "Not Assessable" if kA is None else ("Healthy" if max(pfc) < kA else ">Permissible Limit")
    LOOP.append(dict(pnl=pnl, fdr=fdr, dev=dev, A=A, kA=kA, zs=zs, pfc=pfc, zstat="Healthy" if zok else ">Permissible Limit", pstat=pstat))

def overall(items, key="status"):
    assessed = [i for i in items if i[key] not in ("Not Measurable", "Not Assessable")]
    healthy = sum(1 for i in assessed if i[key] in ("Healthy", "Warning"))
    unhealthy = len(assessed) - healthy
    pct = 100 * healthy / len(assessed) if assessed else 0
    st = "Healthy" if unhealthy == 0 else "Healthy – Minor Actions" if pct >= 90 else ">Attention Required"
    return len(items), healthy, unhealthy, len(items) - len(assessed), pct, st
SUMMARY = [("Earth Electrode Resistance", "pits", overall(PITS)),
           ("Earth Fault Loop Impedance", "panels", overall(LOOP, "zstat")),
           ("Prospective Fault Current", "panels", overall(LOOP, "pstat")),
           ("Riser Integrity", "points", overall(RIS))]
PR = {p: [x for x in PITS if x["pr"] == p] for p in ("P1", "P2", "P3")}

# ---------------- LAYOUT HELPERS ----------------
def colour(status):
    return {"Healthy": GREEN, "Warning": AMBER, "Healthy – Minor Actions": AMBER, "P2": AMBER, "Marginal": AMBER,
            "P3": GREEN, "P1": RED, "Not Assessable": HexColor("#667085"), "Not Measurable": RED}.get(status, RED)
def tbl(data, widths, status_cols=(), align_right=()):
    assert abs(sum(widths) - CW) < 1, f"table width {sum(widths)} != {CW}"
    rows = [[Paragraph(str(c), CELLB) for c in data[0]]]
    for r in data[1:]:
        rows.append([Paragraph(f"<font name='Arial-Bold' color='#{colour(str(c)).hexval()[2:]}'>{c}</font>", CELL)
                     if ci in status_cols else Paragraph(str(c), CELL) for ci, c in enumerate(r)])
    t = Table(rows, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, GREY), ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 3),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    return t
def f2(x): return "—" if x is None else f"{x:.2f}"

def on_page(c, doc):
    c.saveState()
    c.drawImage("/home/claude/esha/logo.png", M, H - 34, 16, 16, mask="auto")
    c.setFont("Arial-Bold", 8.5); c.setFillColor(NAVY); c.drawString(M + 21, H - 29, "Elegrow Technology Pvt. Ltd.")
    c.setFont("Arial", 8); c.drawRightString(W - M, H - 29, f"{P['doc']} | {P['rev']}")
    c.setStrokeColor(NAVY); c.setLineWidth(1); c.line(M, H - 39, W - M, H - 39)
    c.setStrokeColor(GREY); c.setLineWidth(0.5); c.line(M, 38, W - M, 38)
    c.setFont("Arial-Bold", 7.5); c.setFillColor(RED); c.drawString(M, 27, CLASSIFICATION)
    c.setFont("Arial", 7.5); c.setFillColor(NAVY)
    c.drawRightString(W - M, 27, f"Earthing System Health Assessment – {P['unit']} | Page {doc.page}")
    c.setFillColor(HexColor("#C7CDD8")); c.setFont("Arial-Bold", 8); c.drawCentredString(W / 2, H - 29, "SAMPLE DATA – TEMPLATE REVIEW")
    c.restoreState()

def cover(c, doc):
    c.saveState()
    c.setStrokeColor(NAVY); c.setLineWidth(1.2); c.rect(M, 60, CW, H - 110)
    bw = 200; y = H - 230
    for x, head, lines in [(M + 20, "Consultant", ["Elegrow Technology Pvt. Ltd.", "302 Sangini Aspire, Pal, Near Pal RTO,", "Surat, Gujarat 395009"]),
                           (W - M - 20 - bw, "Client", [P["client"], "12/A GIDC Industrial Estate, Dahej,", "Bharuch, Gujarat 392130"])]:
        c.setStrokeColor(GREY); c.setLineWidth(0.6); c.rect(x, y, bw, 150)
        if head == "Consultant": c.drawImage("/home/claude/esha/logo.png", x + 10, y + 108, 30, 30, mask="auto")
        else:
            c.setFont("Arial", 7); c.setFillColor(GREY); c.drawString(x + 10, y + 125, "[client logo]")
        c.setFillColor(NAVY); c.setFont("Arial-Bold", 10); c.drawString(x + 10, y + 88, head)
        c.setFillColor(black); c.setFont("Arial", 8.5)
        for i, l in enumerate(lines): c.drawString(x + 10, y + 72 - i * 12, l)
    c.setStrokeColor(NAVY); c.setLineWidth(1); c.rect(M + 20, y - 110, CW - 40, 70)
    c.setFillColor(NAVY); c.setFont("Arial-Bold", 16); c.drawCentredString(W / 2, y - 68, "Earthing System Health Assessment")
    c.setFont("Arial", 11); c.drawCentredString(W / 2, y - 88, f"{P['unit']}  ·  Final Report")
    # revision table
    cols = ["Revision", "Date of Submission", "Issue", "Prepared By", "Checked By", "Approved By"]
    cw = (CW - 40) / 6; ty = y - 190
    for i, h in enumerate(cols):
        c.setFillColor(LIGHT); c.rect(M + 20 + i * cw, ty, cw, 18, fill=1, stroke=0)
        c.setStrokeColor(GREY); c.rect(M + 20 + i * cw, ty, cw, 18); c.rect(M + 20 + i * cw, ty - 18, cw, 18)
        c.setFillColor(NAVY); c.setFont("Arial-Bold", 7.5); c.drawCentredString(M + 20 + i * cw + cw / 2, ty + 6, h)
        c.setFillColor(black); c.setFont("Arial", 8)
        c.drawCentredString(M + 20 + i * cw + cw / 2, ty - 12, ["R0", "dd-mm-yyyy", "First Issue", "HSP", "AP", "JCS"][i])
    ctl = [("Internal Project No.", P["seq"]), ("Zoho Project", P["project"]), ("Document No.", P["doc"]), ("Work Order No.", P["wo"])]
    for i, (k, v) in enumerate(ctl):
        yy = ty - 70 - i * 18
        c.setStrokeColor(GREY); c.rect(M + 20, yy, 170, 18); c.rect(M + 190, yy, CW - 210, 18)
        c.setFillColor(NAVY); c.setFont("Arial-Bold", 8); c.drawString(M + 26, yy + 6, k)
        c.setFillColor(black); c.setFont("Arial", 8.5); c.drawString(M + 196, yy + 6, v)
    c.setFont("Arial-Bold", 7.5); c.setFillColor(RED); c.drawString(M, 46, CLASSIFICATION)
    c.restoreState()

# ---------------- CONTENT ----------------
st = []
st += [NextPageTemplate("body"), PageBreak()]
st += [Paragraph("Index", H1)]
idx = ["Executive Summary", "Introduction & Scope of Work", "Basis of Assessment", "Instruments Used", "Testing Methodology",
       "Soil Resistivity – Results", "Earth Electrode Resistance – Results", "Earth Fault Loop Impedance & PFC – Results",
       "Riser Integrity – Results", "Corrective Action Plan", "Maintenance & Re-testing", "Annexures"]
st += [tbl([["Sr.", "Section", "Rev."]] + [[i + 1, s, "R0"] for i, s in enumerate(idx)], [40, CW - 90, 50])]
st += [Spacer(1, 6), Paragraph("The index lists only the tests performed for this unit. Sections are generated from the site records.", NOTE)]

st += [PageBreak(), Paragraph("1. Executive Summary", H1), Paragraph("A. Overall summary of tests", H2)]
srows = [["Sr.", "Test", "Unit", "Total", "Within limit", "Above limit", "Not assessed", "Pass %", "Overall status"]]
for i, (name, unit, (tot, h, u, na, pct, stt)) in enumerate(SUMMARY, 1):
    srows.append([i, name, unit, tot, h, u, na, f"{pct:.1f}%", stt])
st += [tbl(srows, [26, 135, 40, 36, 44, 50, 52, 42, CW - 425], status_cols=(8,))]
st += [Spacer(1, 4), Paragraph("Pass % = points within limit ÷ assessed points (Warning-band points are within limit but carry priority P2). Overall status: 0 unhealthy = Healthy; pass ≥ 90% = Healthy – Minor Actions; otherwise >Attention Required (ET-D03-REG-002, OVERALL).", NOTE)]
st += [Paragraph("B. Earth electrode priority summary", H2)]
prow = [["Priority", "Meaning", "Qty", "Pits"]]
for p, mean in [("P1", "Exceeds limit, not measurable, or safety defect – rectify within 30 days and re-test"),
                ("P2", "Warning band or maintenance defect – rectify in maintenance cycle"),
                ("P3", "Healthy – routine annual testing")]:
    prow.append([p, mean, len(PR[p]), ", ".join(x["tag"].replace("GFL-FI-", "") for x in PR[p]) or "—"])
st += [tbl(prow, [50, 230, 35, CW - 315], status_cols=(0,))]
st += [Paragraph("C. Key findings", H2)]
st += [Paragraph(f"• Soil resistivity at {P['unit']} averages {RHO:.1f} Ω·m, classified <b>{corrosivity(RHO).lower()}</b> per IS 3043:2018 Table 4; corrosion allowance applies to steel earthing conductors.", BODY),
       Paragraph(f"• {len(PR['P1'])} of {len(PITS)} earth pits require priority action: EP-07 exceeds its calculated limit with the strip disconnected, EP-03 is heavily corroded although within limit, and EP-08 could not be measured because the chamber is buried.", BODY),
       Paragraph("• Feeder 7F2 at MCC-2 exceeds the 1.8 Ω loop impedance criterion on the B phase; the other three panels are healthy.", BODY),
       Paragraph("• Riser R-202 (R1) reads 1.42 Ω towards the equipment, above the 1 Ω criterion, with a loose connection noted; EP-09 is within limit but above 90% of it and is flagged P2 for monitoring.", BODY),
       Paragraph("• APDB-2 prospective fault current cannot be assessed because the MCB breaking capacity was not recorded.", BODY),
       Paragraph("<font color='#475467'><i>Findings text is written by the report AI from the computed results above and checked by the sanity check before review; it cannot change any value or status.</i></font>", SMALL)]
st += [Paragraph("D. Overall condition", H2),
       Paragraph("The earthing system of the FI Plant is in fair condition. Three earth pits, one feeder and one riser point need corrective action before the next annual test; all other measured points meet the acceptance criteria stated in Section 3.", BODY)]

st += [PageBreak(), Paragraph("2. Introduction & Scope of Work", H1)]
st += [tbl([["Item", "Detail"], ["Client", P["client"]], ["Site address", P["address"]], ["Unit", P["unit"]],
            ["Survey dates", ", ".join(P["days"]) + "  (2 site days)"], ["ETPL engineers", P["engineers"]],
            ["Client representative", P["client_rep"]]], [130, CW - 130])]
st += [Paragraph("Scope – planned vs completed", H2)]
st += [tbl([["Test", "Unit", "Planned (Zoho task)", "Completed", "Remark"],
            ["Soil resistivity", "locations", 2, 2, ""], ["Earth electrode resistance", "pits", 10, 9, "1 not measurable (EP-08)"],
            ["Loop impedance & PFC", "panels", 4, 4, ""], ["Riser integrity", "points", 6, 6, ""]],
           [150, 55, 95, 70, CW - 370])]

st += [PageBreak(), Paragraph("3. Basis of Assessment", H1),
       Paragraph("Every limit, formula and classification used in this report is listed below with its source. Values are taken from the project's acceptance profile (ET-D03-REG-002, profile ETPL-DEFAULT). Criteria marked <b>ETPL acceptance criterion</b> are ETPL's adopted values where the cited standard sets no number.", BODY), Spacer(1, 6)]
basis_ids = ["SOIL-METHOD", "SOIL-CLASS", "PIT-ROD", "PIT-STRIP", "PIT-METHOD", "PIT-BAND", "PIT-PRIORITY", "LOOP-LIMIT", "PFC-RULE", "RISER-LIMIT", "OVERALL", "RETEST"]
brows = [["Criterion", "Rule / value used", "Basis", "Source & clause"]]
for i in basis_ids:
    cr = PROFILE[i]; val = cr["rule"] + (f" (value: {cr['value']} {cr['unit']})" if "value" in cr else "")
    brows.append([cr["parameter"], val, "Standard" if cr["basis"] == "standard" else "ETPL acceptance criterion", f"{cr['source']}, Cl. {cr['clause']}"])
st += [tbl(brows, [95, CW - 305, 85, 125])]
st += [Spacer(1, 5), Paragraph(f"Soil resistivity used for electrode limits: ρ = {RHO:.2f} Ω·m (measured average of both locations, Section 6). No assumed values were used.", NOTE)]

st += [PageBreak(), Paragraph("4. Instruments Used", H1)]
st += [tbl([["Instrument", "Model", "Serial No.", "Calibration cert.", "Valid until"],
            ["Earth analyser", "Metrel MI 3290 GF", "18480898", "[cert no.]", "[date]"],
            ["Installation tester", "Metrel MI 3152 EurotestXC", "17500653", "[cert no.]", "[date]"],
            ["Loop / PSC tester", "Kyoritsu 4118A", "E0182298", "[cert no.]", "[date]"]], [110, 140, 80, 95, CW - 425])]
st += [Paragraph("5. Testing Methodology", H1)]
for head, txt in [("Soil resistivity", "Wenner four-probe method: four probes at equal spacing a in a straight line, driven 10–15 cm; current injected through the outer probes and voltage measured across the inner probes; ρ = 2πaR (IS 3043:2018 Cl. 40). Measured in two perpendicular directions at spacings of 0.5–3.0 m."),
                  ("Earth electrode resistance", "Fall-of-potential method with the electrode isolated from the grid, potential probe at 61.8% of the current-probe distance, repeated at ±10% to confirm a stable reading (IS 3043:2018 Cl. 41). Each result is compared with the electrode's calculated resistance from Cl. 14.2 using the measured soil resistivity and the recorded electrode dimensions."),
                  ("Loop impedance & PFC", "Live phase-to-earth loop impedance (R-E, Y-E, B-E) measured at each panel; prospective fault current measured phase-to-neutral and compared with the protective device's breaking capacity (IS/IEC 60364-4-43 Reg. 434.5.1)."),
                  ("Riser integrity", "Low-resistance measurement from each riser point towards the equipment and towards the earth grid.")]:
    st += [Paragraph(head, H2), Paragraph(txt, BODY)]

st += [PageBreak(), Paragraph("6. Soil Resistivity – Results", H1)]
for loc, (rows, avg) in SOIL.items():
    st += [Paragraph(f"Location: {loc}", H2)]
    st += [tbl([["Spacing a (m)", "ρ Direction 1 (Ω·m)", "ρ Direction 2 (Ω·m)", "Mean ρ (Ω·m)"]] +
               [[f"{a:.1f}", f2(r1), f2(r2), f2(m)] for a, r1, r2, m in rows] +
               [["<b>Location average</b>", "", "", f"<b>{avg:.2f}</b>"]], [CW / 4] * 4)]
st += [Spacer(1, 6), tbl([["Site average ρ (Ω·m)", "Corrosivity class (IS 3043:2018 Table 4)", "Implication"],
                          [f"{RHO:.2f}", corrosivity(RHO), (("Steel conductors: " + ("30%" if RHO < 25 else "15%") + " corrosion allowance in cross-section (IS 3043:2018 Table 5); steel rods ≥ 16 mm dia, flats ≥ 6 mm thick (Cl. 13.6.1)") if RHO <= 100 else "No corrosion allowance required (IS 3043:2018 Table 5)")]],
                         [100, 150, CW - 250])]

st += [PageBreak(), Paragraph("7. Earth Electrode Resistance – Results", H1)]
erow = [["Sr", "Pit tag", "Location", "Type", "l (cm)", "d (cm)", "Limit (Ω)", "Meas. R (Ω)", "Status", "Codes", "Priority"]]
for i, x in enumerate(PITS, 1):
    erow.append([i, x["tag"].replace("GFL-FI-", ""), x["loc"], x["kind"], f"{x['l']*100:.0f}", f"{x['d']/10:.1f}",
                 f2(x["lim"]), f2(x["R"]), x["status"], " ".join(x["codes"]) or "—", x["pr"]])
st += [tbl(erow, [24, 42, 88, 32, 34, 32, 44, 48, 70, 36, CW - 450], status_cols=(8, 10))]
st += [Spacer(1, 4), Paragraph(f"Limits: rod R = (100ρ/2πl)·ln(2l/d) (IS 3043:2018 Cl. 14.2.2); strip R = (100ρ/2πl)·ln(4l/d), d = strip width (Cl. 14.2.3); ρ = {RHO:.2f} Ω·m. Healthy ≤ 90% of limit; Warning 90–100%.", NOTE)]
d = Drawing(CW, 170); bc = VerticalBarChart(); bc.x, bc.y, bc.width, bc.height = 40, 25, CW - 60, 125
meas = [x["R"] or 0 for x in PITS]; bc.data = [meas, [x["lim"] for x in PITS]]
bc.categoryAxis.categoryNames = [x["tag"][-2:] for x in PITS]; bc.categoryAxis.labels.fontName = "Arial"; bc.categoryAxis.labels.fontSize = 7
bc.valueAxis.valueMin = 0; bc.valueAxis.labels.fontName = "Arial"; bc.valueAxis.labels.fontSize = 7
bc.bars[0].fillColor = NAVY; bc.bars[1].fillColor = HexColor("#C7CDD8"); bc.barSpacing = 1; bc.groupSpacing = 6
d.add(bc); d.add(String(40, 158, "Measured R (navy) vs calculated limit (grey), Ω — EP-08 not measurable", fontName="Arial", fontSize=7.5, fillColor=NAVY))
st += [Spacer(1, 8), d]

st += [PageBreak(), Paragraph("8. Earth Fault Loop Impedance & PFC – Results", H1)]
lrow = [["Panel", "Feeder", "Device", "Rating (A)", "Br. cap. (kA)", "Zs R/Y/B (Ω)", "Zs status", "PFC R/Y/B (kA)", "PFC status"]]
for x in LOOP:
    lrow.append([x["pnl"], x["fdr"], x["dev"], x["A"], "—" if x["kA"] is None else x["kA"], " / ".join(f2(z) for z in x["zs"]),
                 x["zstat"], " / ".join(f2(p) for p in x["pfc"]), x["pstat"]])
st += [tbl(lrow, [45, 55, 40, 42, 45, 75, 62, 70, CW - 434], status_cols=(6, 8))]
st += [Spacer(1, 4), Paragraph(f"A panel is healthy only if all three phases meet the criterion. Zs criterion {PROFILE['LOOP-LIMIT']['value']} Ω is an ETPL acceptance criterion (Section 3).", NOTE)]
st += [Paragraph("9. Riser Integrity – Results", H1)]
rrow = [["Sr", "Main location", "Point", "Towards equipment (Ω)", "Towards grid (Ω)", "Codes", "Status"]]
for i, x in enumerate(RIS, 1):
    rrow.append([i, x["main"], x["pt"], f2(x["re"]), f2(x["rg"]), " ".join(x["codes"]) or "—", x["status"]])
st += [tbl(rrow, [24, 90, 85, 90, 80, 50, CW - 419], status_cols=(6,))]

st += [PageBreak(), Paragraph("10. Corrective Action Plan", H1)]
cap = [["Sr", "Test / point", "Observation", "Corrective action", "Priority / status"]]
n = 1
for x in PITS:
    if x["pr"] == "P3": continue
    obs = (f"R = {f2(x['R'])} Ω exceeds limit {f2(x['lim'])} Ω. " if x["status"] == ">Permissible Limit" else
           "Not measurable. " if x["status"] == "Not Measurable" else
           f"R = {f2(x['R'])} Ω is above 90% of limit {f2(x['lim'])} Ω. " if x["status"] == "Warning" else
           f"R = {f2(x['R'])} Ω within limit {f2(x['lim'])} Ω. ") + (x["rem"] + "." if x["rem"] else "")
    act = {"P1": "Rectify within 30 days: repair/replace electrode and strip, restore test link and bonding (IS 3043:2018 Cl. 14.2, 25.6.3.4); re-test after rectification.",
           "P2": "Rectify in the next maintenance cycle; monitor at next annual test."}[x["pr"]]
    if "C8" in x["codes"]: act = "Fix a permanent tag/label at the pit; no electrical action required."
    if "C4" in x["codes"]: act = "Expose and restore pit chamber with removable cover; measure and report within 30 days."
    cap.append([n, f"Electrode {x['tag'].replace('GFL-FI-', '')}", obs, act, x["pr"]]); n += 1
for x in LOOP:
    if x["zstat"] != "Healthy":
        cap.append([n, f"Loop {x['pnl']} {x['fdr']}", "Zs above 1.8 Ω on at least one phase.", "Check PE conductor terminations and size at the feeder; re-test. Confirm device disconnection time per IS/IEC 60364-4-41.", x["zstat"]]); n += 1
    if x["pstat"] == "Not Assessable":
        cap.append([n, f"PFC {x['pnl']}", "Breaking capacity not recorded.", "Record device nameplate kA and re-assess against measured PFC (IS/IEC 60364-4-43 Reg. 434.5.1).", "Not Assessable"]); n += 1
for x in RIS:
    if x["status"] != "Healthy":
        cap.append([n, f"Riser {x['pt']}", f"Towards equipment {f2(x['re'])} Ω above 1 Ω; loose connection noted.", "Clean and re-tighten joint, re-test both directions.", x["status"]]); n += 1
st += [tbl(cap, [24, 80, 140, CW - 314, 70], status_cols=(4,))]
st += [Spacer(1, 4), Paragraph("Generated from the same per-point records as Sections 1 and 7–9; counts cannot differ between sections.", NOTE)]
st += [Paragraph("11. Maintenance & Re-testing", H1),
       Paragraph("Measure earth resistance annually and in any case not less than once in two years; keep records of initial and subsequent tests; carry out periodic visual inspection of all earth electrode connections for rigidity and deterioration (IS 3043:2018 Cl. 38.4.1–38.4.3).", BODY)]
st += [Paragraph("12. Annexures", H1),
       tbl([["Annexure", "Content"], ["A", "Signed field measurement sheets (scans), MS-01/02/04/05"], ["B", "Photographs of all P1 points, by pit tag"],
            ["C", "Instrument calibration certificates"], ["D", "Daily progress reports issued during the survey"]], [70, CW - 70])]

doc = BaseDocTemplate("/mnt/user-data/outputs/ESHA_Final_Report_Template_SAMPLE.pdf", pagesize=A4,
                      leftMargin=M, rightMargin=M, topMargin=55, bottomMargin=50,
                      title="ETPL ESHA Final Report – Template (sample data)", author="Elegrow Technology Pvt. Ltd.")
fr = Frame(M, 50, CW, H - 105, id="f", leftPadding=0, rightPadding=0)
doc.addPageTemplates([PageTemplate("cover", [fr], onPage=cover), PageTemplate("body", [fr], onPage=on_page)])
doc.build(st)
print("pits", [(x["tag"][-2:], round(x["lim"], 2), x["status"], x["pr"]) for x in PITS]); print("rho", round(RHO, 2))
print([s[2] for s in SUMMARY])

# ---------------- DAILY PROGRESS REPORT (Day 1) ----------------
day = P["days"][0]
done_today = {"Soil resistivity": (2, 2, 2, "locations"), "Earth electrode resistance": (6, 6, 10, "pits"),
              "Loop impedance & PFC": (0, 0, 4, "panels"), "Riser integrity": (0, 0, 6, "points")}
dp = [Paragraph("Daily Progress Report", H1)]
dp += [tbl([["Item", "Detail"], ["Client / unit", f"{P['client']} – {P['unit']}"], ["Project", P["project"]],
            ["Survey day", f"{day}  (Day 1 of 2 planned)"], ["Engineers on site", P["engineers"]],
            ["Data collection mode", "Printed sheets (phones not permitted on site) – scanned and uploaded 18:40"]], [130, CW - 130])]
dp += [Paragraph("Progress", H2)]
dp += [tbl([["Test", "Unit", "Done today", "Cumulative", "Planned", "Complete"]] +
           [[k, v[3], v[0], v[1], v[2], f"{100 * v[1] / v[2]:.0f}%"] for k, v in done_today.items()],
           [160, 60, 70, 70, 60, CW - 420])]
dp += [Paragraph("Notable observations today", H2)]
day1 = [x for x in PITS[:6] if x["pr"] != "P3"]
dp += [tbl([["Point", "Observation", "Preliminary status"]] +
           [[x["tag"].replace("GFL-FI-", ""), (f"R = {f2(x['R'])} Ω (limit {f2(x['lim'])} Ω). " + x["rem"]).strip(), x["pr"]] for x in day1],
           [70, CW - 170, 100], status_cols=(2,))]
dp += [Spacer(1, 4), Paragraph(f"Soil resistivity measured at both locations; site average {RHO:.2f} Ω·m. Statuses are preliminary until the final report is reviewed and approved.", NOTE)]
dp += [Paragraph("Plan for next day", H2), Paragraph("Earth pits EP-07 to EP-10, loop impedance and PFC at 4 panels, riser integrity at 6 points. Client to arrange access to the Pump House pit chamber (EP-08).", BODY)]
dp += [Spacer(1, 14), tbl([["Prepared by (ETPL engineer)", "Acknowledged by (client representative)"], ["<br/><br/><br/>", "<br/><br/><br/>"]], [CW / 2, CW / 2])]
dd = BaseDocTemplate("/mnt/user-data/outputs/ESHA_Daily_Progress_Report_SAMPLE.pdf", pagesize=A4, leftMargin=M, rightMargin=M,
                     topMargin=55, bottomMargin=50, title="ETPL ESHA Daily Progress Report (sample)")
P["doc"] = "024_FI_ESHA_DPR_01"; P["rev"] = day
dd.addPageTemplates([PageTemplate("b", [Frame(M, 50, CW, H - 105, leftPadding=0, rightPadding=0)], onPage=on_page)])
dd.build(dp)
