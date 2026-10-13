# ETPL Earthing System Health Assessment — Reporting & Field Standard

**Status:** v0.4 DRAFT for review by Shivam & Jignesh Sailor · **Date:** 7 Oct 2026
**v0.4 changes:** measurement sheets rebuilt as MS-01–MS-09 from ETPL's existing sheets, IEEE 81-2012 and ETPL instruments; riser and step & touch rules corrected from IEEE 81-2012.
**v0.3 changes:** Part 8 Standards Acceptance List added (per-client profiles); Part 7 open item closed; sheets and report template issued as drafts.
**v0.2 changes:** condition codes revised (C1–C8); soil uses IS 3043 Table 4; strip formula, 61.8% FOP and re-test interval added from IS 3043:2018; limits not in IS 3043 relabelled as ETPL acceptance criteria; sheet units approved.
**Replaces:** all previous ESHA report templates and measurement sheets (2023–2026), which become reference only.
**Basis:** ETPL Knowledge Hub › 02 Technical Knowledge (IS 3043:2018 full text and ETPL notes, standards register ET-D02-REG-001). Past site reports were used only to understand how work is done and what clients receive.
**Once approved:** issue as a controlled ETPL document (ET-D03-… numbering via the SOP builder) and implement in ElectroReports.

All IS 3043 citations below were checked against the full IS 3043:2018 text. Limits that IS 3043:2018 does **not** contain are marked **ETPL acceptance criterion** and must show their actual source (client specification, IEEE/IEC standard or ETPL practice) in the report's Basis of Assessment table.

---

## PART 1 — PRINCIPLES (apply to every report and sheet)

1. **One source of truth.** Every number in a report (scope quantity, table rows, summary counts, priority counts, CAP counts) is computed from the same per-point records. No count or summary sentence is ever typed by hand.
2. **Every limit is shown with its basis.** Each test section prints the limit value, the formula or rule, and the exact clause. No limit appears without a citation.
3. **Correct citations only.** Cite the edition actually used. IS 3043 references are to the **2018 (Second Revision)** clause numbers.
4. **No carry-over text.** A report may contain only text generated for that project, unit and date. The app blocks any summary sentence that names a different unit, client or ρ value than the record.
5. **Assumptions are declared.** If a value is assumed (e.g. soil ρ not measured), the assumption, its value and its source appear in the summary and in the calculation table.
6. **Calculations by code, wording by AI.** Values, limits, statuses and priorities come from `reportModel.js`. AI writes observation/recommendation sentences around those results and never changes them.
7. **Raw data is evidence.** The signed field sheet (scan) or the app's signed digital record is appended to every report.

---

## PART 2 — ASSESSMENT RULES (by test)

### 2.1 Soil Resistivity (Wenner 4-probe)
- **Method:** four probes in a straight line at equal spacing *a*; probe depth **10–15 cm** (IS 3043:2018 Cl. 40.4.1). ρ = 2πaR (Cl. 40.3, Eq. 8). Soil resistivity ranges from 1 to 10 000 Ω·m (Cl. 40.1.1); test soil uniformity per Cl. 40.5.
- **Directions:** 2 perpendicular directions minimum; 4 (N/S/E/W) where the client or site requires.
- **Spacings:** 0.5, 1.0, 1.5, 2.0, 2.5, 3.0 m standard set (extend if required).
- **Reported values:** per spacing per direction, mean per spacing, location average, site average.
- **Classification — CORRECTION:** past reports printed "Low < 100 / Medium 100–500 / High > 500 Ω·m — Category as per IS 3043 & IEEE 81". That classification does not appear in IS 3043:2018. The standard's soil classification is **corrosivity, Table 4 (Cl. 13.6.1)**:

| ρ (Ω·m) | Class (IS 3043:2018 Table 4) |
|---|---|
| < 25 | Severely corrosive |
| 25 – 50 | Moderately corrosive |
| 50 – 100 | Mildly corrosive |
| > 100 | Very mildly corrosive |

  The report prints the measured ρ and its Table 4 corrosivity class. Soil resistivity has **no pass/fail**. Low/Medium/High is **dropped** — it has no basis in IS 3043:2018.
- **Corrosivity drives recommendations:** in soil < 100 Ω·m, steel conductors need corrosion allowance (Table 5: 15% for 25–100, 30% below 25); steel flats ≥ 6 mm thick, steel rods ≥ 16 mm diameter (Cl. 13.6.1).

### 2.2 Earth Electrode (Pit) Resistance — the only test with P1/P2/P3
- **Method:** Fall-of-Potential, 3-terminal (IS 3043:2018 Cl. 41.1), R = V/I; electrode isolated (test link open); potential probe at **61.8%** of the current-probe distance (Cl. 41 a); repeat at ±10% to confirm stability. Alternative methods allowed by Cl. 41: clamp-on CT (without disconnection), slope test (large grids/substations), 90°/180° check. The method used is recorded per pit.
- **Inputs per pit (manual):** electrode type, material, length *l*, diameter *d* (or plate area *A*), measured R without grid, measured R with grid (where taken).
- **Standard limit, by electrode type:**
  - Rod / pipe: **R = (100ρ / 2πl) · ln(2l/d)**, *l* and *d* in **cm**, ρ in Ω·m — IS 3043:2018 **Cl. 14.2.2**.
  - Plate: **R = (ρ/4) · √(π/A)**, *A* = area of both sides in m² — IS 3043:2018 **Cl. 14.2.1**.
  - Strip / buried conductor: **R = (100ρ / 2πl) · ln(4l/d)**, *l* in cm, *d* = strip width, or twice the conductor diameter, in cm — IS 3043:2018 **Cl. 14.2.3**. Note **4l/d** for strips vs **2l/d** for rods.
- **ρ used:** this site's measured average. If soil was not measured, the engineer must enter an assumed ρ with its source; it is printed as an assumption.
- **Status:** Healthy if R ≤ 0.9 × limit; Warning if 0.9 × limit < R ≤ limit; >Permissible Limit if R > limit; Not Measurable if the reading couldn't be taken.
- **Priority:**

| Priority | Condition |
|---|---|
| **P1 — High** | R > limit, OR Not Measurable, OR condition code C1, C2, C4 or C5 |
| **P2 — Medium** | Warning band (above 90% of limit), OR condition code C3, C6, C7 or C8 |
| **P3 — Low** | Healthy and no condition code |

  A healthy pit with no condition code is always P3, never P1.
- **Physical checks from IS 3043:2018** (become checklist items on the sheet):
  - Equipment 250–650 V and HV neutrals earthed by **two separate and distinct connections** (Cl. 4.2) — recorded as optional code C3.
  - Pipe wall ≥ 1.6 mm (GI/Cu/SS); rod top ≥ 1.25 m below ground for driven rods (Cl. 14.2.2).
  - Connection points ≥ 150 mm above ground; painted surfaces cleaned; tinned at galvanised steel; protected with bitumastic paint where corrosion-prone (Cl. 25.6.3.4).

### 2.3 Earth Continuity
- Measured R between equipment and earth bar (4-wire DLRO).
- **Limit:** < 1 Ω — **ETPL acceptance criterion**. IS 3043:2018 has no numeric continuity limit; the value comes from IS 3043:1987 Cl. 22.2.3 (superseded). Cite it as such, or replace with the client's specification.
- Status: Healthy / >Permissible Limit.

### 2.4 Earth Fault Loop Impedance (Zs)
- Measured R-E, Y-E, B-E per feeder; voltage recorded.
- **Limit:** per-project setting, default **1.8 Ω** — **ETPL acceptance criterion**. IS 3043:2018 Cl. 42 describes the loop but sets no numeric limit (the true limit depends on the protective device and disconnection time, IS/IEC 60364-4-41). The report states the value used and its source.
- A **panel/feeder is Healthy only if all three phases are Healthy**. Summary counts panels and points separately.
- **Record protective device type, A rating and kA breaking capacity** for every feeder (missing in current practice).

### 2.5 Prospective Fault Current (PFC)
- Measured R-N, Y-N, B-N; PFC (kA).
- **Rule:** PFC < device breaking capacity (kA) → Healthy; ≥ → >Permissible Limit.
- If breaking capacity was not recorded → status **Not Assessable**, PFC recorded for reference, and a CAP item to obtain nameplate data. This status is shown separately, never counted as Healthy.

### 2.6 Riser Integrity
- Two readings: towards equipment, towards grid.
- **Limit:** both ≤ 1 Ω → Healthy — **ETPL acceptance criterion** (not in IS 3043:2018; past reports cite IEEE 81-2012 Cl. 10 — confirm against the hub copy of IEEE 81 before citing).

### 2.7 Ground Resistance
- Measured R per location. Compared with the same calculated electrode limit (Cl. 14.2) as the pit it belongs to, unless the client specifies a value; where a common earth is used, IS 3043:2018 notes 1 Ω or less as usual for power stations and large substations.

### 2.8 Step & Touch Potential (optional)
- Tolerable limits per IS 3043:2018 **Cl. 10.3** for 50 kg and 70 kg body weight:
  - E_step = (1000 + 6·Cs·ρs) × k / √ts
  - E_touch = (1000 + 1.5·Cs·ρs) × k / √ts
  - k = 0.116 (50 kg), 0.157 (70 kg); ts = fault clearing time (s); ρs = surface layer resistivity (= ρ if no surface layer).
- Inputs required on the sheet: ρs, Cs, ts (with source). Healthy if measured < tolerable.

### 2.9 Tower Footing
- 4 feet per tower; Zt = Σ foot impedances; Zsat default 10 Ω (project setting).
- Healthy ≤ Zsat; Marginal ≤ 1.2 × Zsat; else Not Acceptable — **ETPL acceptance criterion** (Zsat and the 1.2 band are not in IS 3043:2018; take Zsat from the utility/client specification and state it).

### 2.10 Overall status per test (summary table)
- 0 unhealthy → **Healthy**
- pass rate ≥ 90% → **Healthy – Minor Actions**
- otherwise → **>Attention Required**
- Not Measurable / Not Assessable points are listed separately and excluded from pass rate, with the count shown.

---

## PART 3 — CONDITION CODES (replace free-text comments)

Ticked on the sheet or in the app per point. They are the main input to priority and to the AI's observation text.

| Code | Condition | Applies to | Priority effect (electrode) |
|---|---|---|---|
| C1 | Electrode / strip corroded | Pit, riser | P1 |
| C2 | Strip or test link disconnected | Pit, riser | P1 |
| C3 | Only one earth connection (Cl. 4.2 requires two) — *optional, record where checked* | Pit, equipment | P2 |
| C4 | Electrode buried / chamber not accessible or damaged | Pit | P1 (Not Measurable if no reading) |
| C5 | Electrode physically damaged | Pit | P1 |
| C6 | Loose connection / bolt | All | P2 |
| C7 | Joint not cleaned, tinned or protected (Cl. 25.6.3.4) | All | P2 |
| C8 | No tag / label on pit | Pit | P2 |
| OTH | Other — one-line remark required | All | Reviewer assigns |

Free-text remarks remain allowed for anything unusual, and the AI is required to surface them in the observation.

---

## PART 4 — MEASUREMENT SHEETS (print version for no-phone sites)

### 4.1 Common header (every sheet, every page)
- ETPL logo · document ID · sheet version
- **QR code** encoding project no. / unit / test type / sheet ID / page no. (OCR auto-routes the scan to the right daily record)
- Client · Site / unit · Project no. · Zoho task · Work date
- Engineer(s) · Instrument model + serial no. + calibration due date
- Weather / soil condition (Dry / Damp / Wet) — affects resistance readings
- Page __ of __

### 4.2 Common footer
ETPL engineer name + signature · Client representative name + signature · Time started / finished.

### 4.3 Layout rules for OCR
- Numbers in **boxed digit cells** with a printed decimal point.
- Units printed in **every column header**.
- Condition codes as **tick-boxes**, plus one short remarks line per row.
- Tag and location **pre-printed** from the asset register / last year's report; a blank "new pit" row block at the end.
- Maximum 12 rows per page so handwriting stays legible.

### 4.4 Sheet set (v2)

| Sheet | Doc no. | Orientation | Rows | Instruments | Basis |
|---|---|---|---|---|---|
| MS-01 Soil Resistivity | ET-D03-FRM-001 | Portrait | 10 spacings 0.5–5.0 m × 4 directions | MI 3290 GX, Kyoritsu 4105A | IS 3043 Cl. 40; IEEE 81 Cl. 7 |
| MS-02 Earth Electrode (Pit) Resistance | ET-D03-FRM-002 | Landscape | 10 pits, tags pre-printed | MI 3290 GX, Kyoritsu 4105A | IS 3043 Cl. 14.2, 41; IEEE 81 Cl. 8 |
| MS-03 Earth Continuity & Bonding | ET-D03-FRM-003 | Portrait | 19 points | MI 3152, MI 3290 GX | IS 3043 Cl. 17.2.3, 21 |
| MS-04 Earth Fault Loop Impedance | ET-D03-FRM-004 | Landscape | 4 panels × R/Y/B-E | Kyoritsu 4118A, MI 3152 | IS 3043 Cl. 42; IS/IEC 60364-6 |
| MS-05 Prospective Fault & Short-circuit Current | ET-D03-FRM-005 | Landscape | 4 panels × R/Y/B | Kyoritsu 4118A, MI 3152 | IS/IEC 60364-4-43 Reg. 434.5.1 |
| MS-06 Riser Integrity | ET-D03-FRM-006 | Portrait | 19 points | MI 3290 GX, MI 3152 | IEEE 81 Cl. 10 |
| MS-07 Ground (Grid) Resistance — **new** | ET-D03-FRM-007 | Portrait | 18 points, readings at 51.8/61.8/71.8% | MI 3290 GX, Kyoritsu 4105A | IS 3043 Cl. 41; IEEE 81 Cl. 8 |
| MS-08 Step & Touch Voltage — **new** | ET-D03-FRM-008 | Landscape | 12 points | MI 3295 (confirm), MI 3290 GX | IEEE 81 Cl. 9; IS 3043 Cl. 10.3 |
| MS-09 Tower Footing Impedance — **new** | ET-D03-FRM-009 | Portrait | 5 towers × 4 feet | MI 3290 GX + clamps | IEEE 81 Cl. 8; IS 3043 Cl. 14 |

Changes from ETPL's previous sheets: ETPL navy theme and fixed header/footer; QR per sheet; work date per sheet (old sheets printed "Start Date" twice on loop and PFC); boxed digits and labelled tick-boxes; condition codes and photo reference on every row (requested on the old continuity sheet); instrument tick and serial no.; loop and PFC kept as separate per-phase sheets; PFC sheet records both PFC (L-PE) and PSC (L-N); "Earth Continuity" and "Continuity Test" merged into MS-03; test-lead resistance recorded (IEEE 81 Cl. 10.2); spelling fixed ("Riser Intrigrity"). Calculated columns (mean ρ, category, limits) removed from print sheets — the app calculates them.

**Step & touch — important:** IEEE 81-2012 Cl. 9.3(g) requires measured voltages to be scaled by If/It (fault current ÷ test current) before comparison with the tolerable limits. MS-08 records the test current per reading and the fault current in the header so the app can do this. Past reports that compared unscaled measured volts directly with tolerable limits must be re-checked.

## PART 5 — REPORT FORMATS

### 5.1 Final report — section order
1. **Cover** — Consultant / Client boxes (with client logo), title, unit/scope, revision table, Internal Project No. (SEQ only), Document No., Work Order No., Zoho project no.
2. **Document control & distribution**
3. **Index** — generated from the sections actually present.
4. **Executive Summary**
   - A. Overall test summary table (Part 2.10 rule)
   - B. Electrode priority summary (P1/P2/P3 with counts and pit tags)
   - C. Key findings (AI-written from the data, 3–6 bullets)
   - D. Overall condition statement
5. **Introduction & Scope** — client, address, scope with quantities **planned vs completed**, survey dates (all days), representatives.
6. **Basis of Assessment** (NEW) — one table: test · method · limit/rule · value used · clause. Includes declared assumptions.
7. **Instruments** — model, serial, calibration certificate no. and due date.
8. **Methodology** — per test performed (standard text).
9. **Results** — one section per test: summary line, results table, chart where useful (no single-slice 100% pie charts).
10. **Corrective Action Plan** — one table, all tests: point/tag · observation · action · priority (electrode) or status (others) · clause.
11. **Recommendations for maintenance & re-testing** — annual measurement, never less than once in two years (IS 3043:2018 Cl. 38.4.2); periodic visual inspection of connections (Cl. 38.4.3); keep test records (Cl. 38.4.1). CEA safety-regulation reference added only after confirming the current clause number.
12. **Annexures** — signed field sheets (scans), photos linked by tag, calibration certificates, MOM.

### 5.2 Daily Progress Report (to client, per site day)
1 page where possible: date · engineers · tests and quantities done today · cumulative progress vs planned · notable observations (P1 / unhealthy found today) · plan for next day. No CAP, no final conclusions.

### 5.3 Formatting
ETPL brand (navy #0B2A67, Arial). Header: logo + company, document no. + revision. Footer: classification, title, page x of y. Numbers right-aligned with fixed decimals (R: 2 dp, ρ: 2 dp, PFC: 2 dp). Dates dd-mm-yyyy. Red / amber / green status cells with text labels (not colour alone).

---

## PART 6 — FIELD & REVIEW GUIDELINES

### 6.1 Engineer — before site
- Confirm in the app: phones allowed? connectivity? → app mode or printed sheets.
- Print sheets from the app only (never reuse old printed sheets).
- Check instrument calibration validity; record serials.

### 6.2 Engineer — on site
- Fill the header before the first reading.
- Isolate electrodes for "without grid" readings; restore test links after.
- Tick condition codes for every point inspected — a blank code column means "inspected, no defect".
- Record which fall-of-potential method was used for each pit.
- Photograph every P1 condition with the tag visible.
- Client representative signs every page daily.

### 6.3 Engineer — end of day
- Scan all pages (phone scanner app, flat, full page, QR visible) and upload, or submit the day in-app.
- Resolve every OCR item flagged as low-confidence before submitting.

### 6.4 Reviewer checklist (before approval)
- Counts match across scope, tables, summary and CAP (the app checks this; reviewer confirms).
- Every P1 has a photo and an action.
- Assumptions declared; limits and clauses shown.
- Readings differing > 50% from last year's value for the same tag are explained.
- No text referring to another unit, client or project.

---

## PART 7 — STATUS OF DECISIONS
| # | Decision | Status |
|---|---|---|
| 1 | Soil classification | **Closed** — IS 3043:2018 Table 4 corrosivity; Low/Medium/High dropped |
| 2 | Sheet units (m / mm) | **Approved** |
| 3 | Condition codes | **Approved with changes** — C1–C8 as in Part 3 |
| 4 | IS 3043:2018 Amendments 1 & 2 | **Confirmed** — no change to Cl. 14.2 |
| 5 | Open limits | **Closed** — strip formula, FOP 61.8%, re-test interval from IS 3043:2018; loop 1.8 Ω, continuity < 1 Ω, riser ≤ 1 Ω, Zsat labelled ETPL acceptance criteria |

| 6 | Source printed for ETPL acceptance criteria | **Closed** — taken from the project's profile in the Standards Acceptance List (Part 8) |

---

## PART 8 — STANDARDS ACCEPTANCE LIST (ET-D03-REG-002)

Limits and their sources differ between clients, so they are **not** fixed in this standard or in the app's code. They live in one controlled register, `ET-D03-REG-002_Standards_Acceptance_List.json`:

- **ETPL-DEFAULT** — profile for Indian industrial and utility sites. Every criterion carries its rule/value, basis (`standard`, `etpl_acceptance` or `client_spec`), source and clause.
- **Client profiles** — created from CLIENT-TEMPLATE when a client's specification sets its own limits. A client profile lists only the criteria that differ; everything else falls back to ETPL-DEFAULT. The hub's `03 Client Specific Standards` folder (ADNOC, Shell, KOC, OETC, Bechtel, Dashiell) is the source for those clients.
- **The project picks its profile** (Zoho project field or app setting). The report's Basis of Assessment table and every calculation read the resolved profile, so the printed source always matches the value used.
- **Controlled changes:** each edit to the register gets a revision, date and approver.

## PART 9 — DRAFT DELIVERABLES (for review)

| File | Content |
|---|---|
| `ESHA_Measurement_Sheets_v2_MS-01_to_MS-09.pdf` | Print sheets ET-D03-FRM-001 to 009 (supersedes v1) |
| `ESHA_Final_Report_Template_SAMPLE.pdf` | Final report template rendered with sample data; all values computed from ET-D03-REG-002 |
| `ESHA_Daily_Progress_Report_SAMPLE.pdf` | One-page daily report to client |
| `ET-D03-REG-002_Standards_Acceptance_List.json` | Acceptance register, machine-readable |
| `generators/` | Python reference implementations of the layouts and rules for Claude Code to port into the app |
