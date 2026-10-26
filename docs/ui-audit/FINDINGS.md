# Session 6J.2 UI audit: phone entry sizes and gradient contrast

Run: `node scripts/ui-audit.js --widths=360 --role=engineer` (28 runs: 17 shared screens, all sheet steps, observation drawer, scan review). Sheet steps now cover **all nine sheets MS-01..MS-09**; MS-07 Ground resistance and MS-08 Step & touch voltage were not reachable before (the seeded day does not enable them), and the retired Earth Continuity step is still included. No calculation, validation or permission code changed; only `public/styles.css` and the dev harness.

| Measure (engineer, 360px) | Before | After |
|---|---|---|
| Sheet-entry controls under 44px (soil 100, tower 17, suggest-wording 12, others 0) | 129 | 0 |
| Entry inputs that clip a 6-character value ("123.45") | not measured; soil 36px and tower 26px wide | 0 of 100 soil / 17 tower inputs (narrowest input on any sheet: 181px) |
| axe violations | 0 | 0 |
| Page overflow, console errors, failed requests | 0 | 0 |
| Gradient / image text nodes measured by hand (axe "needs review") | 2,765 sampled of 4,274 listed (26 runs) | 3,084 sampled (28 runs, incl. MS-07 and MS-08) |
| Of those, below 4.5:1 (3:1 for large text) | 217 samples, 30 distinct styles, worst 1.02:1 | 1 sample, a false positive (below) |
| Nodes not measured because something else covers them (dialog, drawer) | n/a | 398 |

Still under 44px at 360px on the engineer screens: only the two Schedule timeline bars (34x30 and 72x30), unchanged from 6J.

## 1. Entry sizes
- Soil matrix and Tower footing are real tables on a phone (not cards), so their columns were squeezed to 36px and 26px. They now have `width: max-content`, a minimum of 88px per input (120px for tower number/location) and scroll sideways inside their own container. Page overflow stays 0.
- "Suggest remark wording" is 44x44 on a phone (was 24x24). Every other sheet-entry input, select, tick cell, add/remove button already measured 44px or more.
- Rule used on phone for the sheet tables: `.sheet-table .table-input, select, .button, .icon-button { min-height: 44px }`.

## 2. Contrast over gradients, measured by hand
Method (in the harness, `readBackdropContrast`): for every node axe lists as "needs review", read the real text colour (with opacity), hide all text, take a screenshot of that element and compute the contrast of the text colour against **every backdrop pixel under the text's line boxes**; the minimum is reported. At most 12 nodes per element/colour/size are sampled, spread over the page, because each needs its own screenshot. Large text (24px, or 18.66px bold) needs 3:1; everything else 4.5:1.

Found and fixed (worst measured ratio, before → after):
| Where | Before | Change |
|---|---|---|
| Scan review, "Discard" (grey ghost button on the dark hero) | 1.02 | light text and border inside `.detail-hero` |
| Dashboard report cards, test chips ("Soil" etc., pale text on white) | 1.08 | dark chip colours, as on other surfaces |
| Schedule, "You" under the sticky name column (translucent cell let bars show through) | 1.85 | opaque cell background |
| Builder and detail heroes, paragraph text | 2.17 | opaque gradient, darker light end |
| Section kickers "STEP 1", "PROJECTS", etc. (#1991ff on pale) | 2.43 to 3.1 | #0b5cad (6.3:1) |
| Primary buttons (white on #34a3ff to #2767ff) and step numbers | 2.97 to 3.42 | gradient #1566c4 to #1747d6 (5.6:1 or more) |
| Success buttons (white on #1d9dd8 start) | 3.76 | #15739f start (5.3:1) |
| Selected tick label and "Loop" (#0059ff on pale blue) | 3.53 | #0047cc |
| Zoho status text Planning (#9c6b11) and Execution (#0e76b6) | 4.23, 4.34 | #7a5600 and #0a5f94 |
| Dashboard hero chip text, report-card date, summary-orb labels, muted text on scan cards | 4.0 to 4.49 | white; #4a5a70; `--muted` now #4a5a70 |
| Decorative particles drifting behind hero text | 2.82 | particle opacity capped at 0.3 (a dot behind a glyph now leaves at least 4.5:1) |

Colour meaning is unchanged (blue stays blue, amber amber). Primary and success buttons and the hero banners are slightly darker than before.

Remaining sample below threshold, judged a false positive: Scan sheets upload page, 3.94:1 on the file input. The pixels measured are the browser's own "Choose Files" button (grey with dark border); the label text is black on light grey. Native control, not styled by the app.

Not measured: the 398 nodes covered by the print dialog or observation drawer (behind an overlay they are not readable anyway), SVG diagram labels, text that appears only on hover, focus or error states.

## Not covered
Other roles and widths were not re-run this session (6J's 308-run table below is still the latest for them; the new CSS applies to all widths, and the phone-only size rules apply at 720px and below). Dark mode, print CSS, 320px, text zoom, screen-reader testing.

---

# Session 6J UI audit: findings

## Status after the 6J fixes (re-run, all four roles)

Everything below the "Before" heading is the original audit, kept as written. After the fixes the harness was extended: it now also runs as engineer, reviewer and approver (308 runs: 29 admin screens + 3 extra sheet steps at 4 widths, plus the role-appropriate screens for the other three) and includes the Tower footing, Continuity and Earth continuity sheet steps. No calculation or permission code was changed. The 6I/6J test suites pass (`test:schedule-ui` needed its heading selector `h2` → `h1`).

| Measure | Before (116 runs, admin) | After (128 admin runs) | After (308 runs, 4 roles) |
|---|---|---|---|
| axe: no h1 (page-has-heading-one), nodes | 108 | 0 | 0 |
| axe: select-name | 11 | 0 | 0 |
| axe: label | 40 | 0 | 0 |
| axe: color-contrast | 114 | 0 | 0 |
| axe: scrollable-region-focusable | 13 | 0 | 0 |
| axe: empty-table-header | 12 | 0 | 0 |
| axe: target-size | 7 | 0 | 0 |
| axe: aria-allowed-role | 4 | 0 | 0 |
| Runs with page / element overflow | 9 / 22 | 0 / 0 | 0 / 0 |
| Elements under 24px | 1221 | 0 | 0 |
| Elements under 44px at 360px | 504 | 135 | n/a |
| Runs with text under 11px | 32 | 0 | 0 |
| Runs with unlabeled controls | 27 | 0 | 0 |
| Console errors, failed requests, focus-indicator gaps | 0 | 0 | 0 |
| Screens without exactly one h1 | 27 of 27 had none | 0 | 0 |

Notes on the numbers: the overflow checker was tightened between runs (it now reports only the innermost overflowing element and ignores content of a closed `<details>`), so the "before" overflow count is the old, looser one. The "under 44px" count is advisory and counts every control, not only the ones you listed.

### What changed
1. **Phone layout.** Root cause of the squashed, covered Print button and the page overflow: a later `.builder-layout-form` rule forced the two-column grid (340–380px rail) back on after the single-column rule, so on phones the form was squeezed to a sliver. It is now single-column up to 920px with the rail below the form. Admin tabs wrap (`.admin-tabs`). The schedule card's grid track had no `min-width: 0`, so the settings table stretched the whole card; fixed (`.sched-card`), and `fieldset { min-width: 0 }`. The builder stepper already scrolled inside its own container; the 29px page overflow was the same two-column bug. Checked by hand at 768px: all step pills fit, nothing is cut.
2. **Contrast.** New tokens `--warning-text` (#7a5600, 6.0:1) and `--neutral-text` (#4a5a70, 6.2:1) for the amber and grey pills; green, red and blue pills already passed (5.1–5.7:1). Colour meaning is unchanged. Schedule Holiday/Leave text uses the darker amber and is larger (0.7rem instead of 0.6rem). Schedule bars: orange `#c4631a` → `#a8510f` and green `#1f8a5b` → `#17724a` so the white text passes. The "assumed ρ" panel heading is darker. Sheet status pills 10.9px → 12px.
3. **Tap targets.** Checkboxes and radios are 24×24, and the whole label is the target: 28px high on desktop, 44×44 on a phone (admin role and Active boxes, schedule days, scan-review ticks, print dialog). Fill-down buttons are 24×24 on desktop and 44×44 on a phone. Also: `summary` rows, icon buttons, the "Why this status" and "Suggest wording" buttons are at least 24px, and the OCR mode chips and ghost buttons are 44px high on a phone.
4. **Labels and structure.** Names for the sign-in user picker, Zoho project picker, dashboard search box, every schedule rate field, and the per-user Active checkbox. Each view has exactly one `<h1>` (the former `h2` titles; the Admin screen got an "Administration" title). Deeper headings carry `aria-level` so the outline has no skipped levels, with no visual change. Empty table headers now have visually hidden text. Every table wider than its box is wrapped, made focusable and named as a scroll region (re-evaluated on resize). The observation drawer is a `div role="dialog"` instead of an `aside`. `workspace-save-planned` has the text "Save planned quantities"; the harness flagged it only because it sits in a closed `<details>`, so it was a false positive and the harness now skips closed details.

### Still open (not fixed, by choice or by size)
- Schedule bars are 30px high (72px wide for one-day tasks) on a phone. They are timeline cells; enlarging them changes the chart.
- ~~Soil resistivity matrix inputs are 36×44px and the tower location inputs 26×44px at 360px~~ Fixed in 6J.2.
- ~~"Suggest remark wording" buttons are 24px on a phone~~ Fixed in 6J.2 (44px).
- ~~About 130 controls on the sheet steps are still under 44px at 360px~~ Fixed in 6J.2: 0 on all nine sheets.
- ~~Contrast over gradients and images was not checked by hand~~ Done in 6J.2. SVG diagram labels render below 11px by scaling.
- Not covered: dark mode, print CSS, 320px width, browser text zoom, screen-reader testing, error and loading states.

---

# Before (original audit, unchanged)

Run: `node scripts/ui-audit.js` (mocks, seeded project with a daily and a final report, one scanned sheet).
29 screens × 4 widths (360 / 768 / 1280 / 1920) = 115 runs. Each run: axe-core (WCAG 2.0/2.1/2.2 A+AA + best-practice) plus the harness's own checks.
Raw data: `reference/ui-audit/results.json`; screenshots: `reference/screenshots/audit/` (both git-ignored).
Nothing has been fixed. No code in `src/` or `public/` was touched.

**Not covered:** the Tower footing / Continuity / Earth continuity sheet steps (the seeded day has only 5 tests), non-admin roles (everything ran as admin), toasts and error/empty states beyond the search-empty case, loading states, dark mode / print CSS, 320px width, browser text zoom, screen-reader testing, contrast *ratios* (axe gives pass/fail here; I did not record the numbers).

## Clean results
- No console errors, page errors or failed requests (4xx/5xx) on any screen.
- Every Tab stop on every screen at 1280 has a visible focus change (0 of ~480 stops without one).
- No duplicate ids, broken images, clipped text, `lang`/viewport/`<main>` all present, no clickable non-focusable `div`s.

## Findings, most severe first

### A. Real layout bugs on phones (360px)
| # | Finding | Where |
|---|---|---|
| A1 | **"Print sheets for site" cannot be tapped.** The sidebar panel (`.builder-aside .aside-panel`) covers it, and the button itself is squashed to 94×96px. Playwright's click is intercepted; even a forced click does not open the dialog. | Builder → Project step |
| A2 | **Admin tab row overflows the page**: 233px (Users), 330px (Audit), 416px (Knowledge base). Tabs and page heading run off screen. | Admin, all tabs |
| A3 | **Schedule header/controls overflow**: 267px at 360 and 9px at 768, only when the settings panel is open. | Schedule → settings |
| A4 | **Builder stepper overflows** by 29px at 360 (page scrolls sideways); at 768 the step pills extend past the viewport (to 1040px) without page overflow, so the later steps are probably reachable only by scrolling the stepper. Needs a manual look. | Builder, every step |

### B. Accessibility, confirmed by axe
| # | Finding | Rule / impact | Where |
|---|---|---|---|
| B1 | No `<h1>` on any screen. | page-has-heading-one, moderate | all 27 screens |
| B2 | Selects with no accessible name: sign-in user picker, `#zohoProjectPicker`. | select-name, **critical** | Sign-in, Builder project step, Print dialog |
| B3 | Inputs with no label: the nine per-test rate fields, and the Active checkbox per user. | label, **critical** | Schedule settings; Admin → Users |
| B4 | Low contrast on status pills (`.pill-neutral`, `.pill-warning`), the "Holiday/Leave" `em` text in the schedule, and the `.assumed-rho-panel` kicker. Up to 7 nodes on the report detail page. | color-contrast, serious | Detail, Workspace, Review, Schedule, Scan review, all sheet steps, Project step, Print dialog |
| B5 | A horizontally scrollable table can't be reached by keyboard. | scrollable-region-focusable, serious | Report detail (3rd sub-section), up to 1280 |
| B6 | Role checkboxes are 13×13 with no spacing. | target-size (WCAG 2.2), serious | Admin → Users at 360/768 |
| B7 | Empty `<th>` (workspace days table, schedule settings). | empty-table-header, minor | Workspace, Schedule settings |
| B8 | `role` not allowed on the observation drawer element. | aria-allowed-role, minor | Observation drawer |
| B9 | One button has no visible text or label: `workspace-save-planned`. (Harness finding; axe did not flag it, see disagreements.) | name | Workspace |

### C. Touch targets and text size (harness only)
- **13×13px checkboxes**: Admin → Users (5), Calendar sync (2), Schedule settings (14 at 360), **Scan review: 208 tick boxes**. This is the field-use screen on a tablet or phone.
- **Fill-down buttons are 16×17px**: 29 on the electrode sheet, 14 each on Loop, PFC and Riser. They sit in table headers.
- `<summary>` rows 18px high: "Planned quantities", the schedule clash lists.
- Print dialog Close button is 18×38px.
- At 360, sheet steps have 16–103 controls under 44px (mode chips are 34px high, "Edit job details" 36px). Schedule bars are 30px high and as narrow as 72px.
- Text under 11px: schedule "Holiday" 9.6px, "Leave" 9.9px. Text under 12px: 45 elements on report detail, 18 on the dashboard, 14 on the schedule.
- Dashboard search box has only a placeholder, no label. Axe accepts this; WCAG 3.3.2 practice says add one.

### D. Where axe and the harness disagree
| Item | Harness | axe | Read |
|---|---|---|---|
| Small targets | ~330 elements under 24px across screens | Only the 13px role checkboxes at 360/768 | axe's spacing exception passes elements with enough free space around them, and ignores inline/summary. Harness is stricter than WCAG 2.2 AA; treat C as advisory except the checkboxes. |
| Dashboard search label | flagged | not flagged | axe counts placeholder as a name. |
| `workspace-save-planned` no name | flagged | not flagged | Probably named from `title` or a hidden span; check by hand. |
| Colour contrast, empty table headers, scrollable region, heading-one, aria role | not covered | flagged | axe-only. |
| Overflow, covered/squashed button, tiny text, stepper | flagged | not covered | harness-only (axe has no layout checks). |
| Select and label problems | flagged | flagged | agree. |
| Axe "incomplete" (needs manual review) | | `color-contrast`, `label-content-name-mismatch` | Not resolved; the second means visible label text differs from the accessible name on some control. |

## Suggested order if you choose to fix
1. A1–A4 (phones: print button, admin tabs, schedule header, stepper). One shared CSS pass.
2. B2, B3 (add labels; small markup change, critical in axe).
3. B1 (an `<h1>` per view), B4 (pill and `em` colours, via CSS tokens), B5 (`tabindex=0` plus label on the scroll wrapper), B7, B8.
4. C: raise checkbox and fill-down hit areas (padding or label wrapping), especially scan review.

## Housekeeping
Uncommitted: `scripts/ui-audit.js`, `package.json` / `package-lock.json` (`@axe-core/playwright` devDependency), and this file. As agreed, the harness is not wired into `npm test`; I'd commit it as `chore: UI audit harness`.
