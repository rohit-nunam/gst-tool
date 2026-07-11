# GST Scrutiny Tool — README (Single Source of Truth)
**(paste this whole document into a new chat and Claude will have complete context — this supersedes and consolidates `GST_Tool_Full_Project_Context.md`, `CHANGELOG_v2.md`, `GST_Tool_v2_Session_Context.md`, `BUGFIX_ROUND_1.md`, and `OCR_LIMITATION.md` into one current-state document)**

**Taxpayer used to build/verify this tool:** M R HEALTHCARE PRIVATE LIMITED, GSTIN `05AAECM6380J1ZA`, pharma manufacturer, Ramnagar, Uttarakhand. **The tool itself is generic** — see §8 for exactly what is/isn't taxpayer-specific.
**Who's building this:** Darshan (Co-Founder/Director at Nunam Technologies — unrelated to his usual work), supporting a GST scrutiny/audit workflow, with a verification team independently testing output against real files.
**What it is:** a GST reconciliation tool that cross-checks a taxpayer's own filings against each other and against department-side data, surfacing every mismatch/anomaly an officer or consultant would raise in scrutiny — always with the exact arithmetic behind the flag, never a bare assertion, and never fabricating a number it doesn't have.

---

## 1. How to run it

```bash
pip install openpyxl pypdf pdfplumber pytesseract --break-system-packages
cd <folder with all data files + 20 .py files, exact script names below>
python3 master_build.py .
```
Output: one workbook, `GST_MASTER_<GSTIN>_FY<range>.xlsx`, written to the same folder.

Data filenames never matter (everything is detected by file content — sheet names, header text, marker text, PDF first-page signature). Only the 20 `.py` filenames matter, exactly as listed in §4.

**For a NEW taxpayer:** replace all data files with the new taxpayer's own files (any names), and either fill in `bs_pl_input.py`'s `BS_PL_DATA` dict with their real Balance Sheet/P&L figures (tagged with their GSTIN — see §8) or leave it empty to skip the BS/PL checks cleanly. Nothing else needs editing for a different taxpayer, financial year, or set of supplied documents — see §9 for the one exception (`HSN_RATE_MASTER`).

---

## 2. Architecture — five layers (was four; the forensic layer is new this round)

1. **Per-month engines** (run once per month that has data): raw side-by-side comparison → interpretive 14-check layer → 27-check E-Way-Bill matrix.
2. **Cross-month layer**: links an error reported in one month to its correction in a later month (Rectification Pairs); checks invoice-series integrity against Table 13's declared document counts (Doc-Series Integrity), now also cross-referenced against cancelled e-invoices.
3. **Annual/FY-wide layer**: reconciles ledgers, TPST, the portal's own comparison report, and the BO/360° Profile against the monthly totals.
4. **HSN-code-wise + fraud-pattern layer**: 36 checks (HSN classification, POS/state-code logic, named fraud patterns), run once across the whole FY.
5. **NEW — Forensic layer**: GSTR-9/9C/Table-8A four-way ITC reconciliation, turnover-gap rule, a generic Balance-Sheet/P&L rule engine, per-month filing-compliance/late-fee, and Cancelled E-Invoices — all run once across the whole FY, alongside layer 4.

**Input files are MERGED, whole-FY workbooks** — one Excel file per document type (GSTR-1, GSTR-3B, E-Invoice, GSTR-2B), each containing every month's data stacked in the same sheets, separated by a period-marker row:
```
Financial Year: 2022-23 | Tax Period: January | ARN: AA0501230730120 | ARN Date: DD-MM-YYYY | ...
```
`merged_period_utils.py` finds these markers by content and slices each sheet into `{month: [rows]}` blocks. GSTR-3B is the exception — it merges as one *sheet* per month, and even there the sheet's *name* is never trusted; the month comes from an in-sheet `Year`/`Tax Period` field. **Multiple merged files of the SAME document type are now supported** (e.g. 5 separate GSTR-1 workbooks, one per FY) — `folder_classifier.py` resolves, per month, which specific file covers it (see §10 for multi-year status).

Every parser is **content-based, never filename- or sheet-name-based.**

---

## 3. Complete input file catalog (16 document types, 14 optional + 2 mandatory)

### 3.1 Original 12 (all optional except #1–2)

| # | File | Frequency / format | Mandatory? | Contents |
|---|---|---|---|---|
| 1 | GSTR-1 (merged) | 1+ files, whole FY, Excel | **Yes** | B2B/SEZ/DE invoices, B2C Large, B2C Small, exports, CN/DN, HSN summary, amendments (b2ba/cdnra), Table 13 |
| 2 | GSTR-3B (merged) | 1+ files, whole FY, Excel (1 sheet/month) | **Yes** | Outward liability, RCM, ITC availed/reversed/ineligible, ARN + filing date |
| 3 | E-Invoice (merged) | 1+ files, whole FY, Excel | No | IRN, IRN date, invoice date/value/rate/tax, B2B/SEZ/DE, **IRN/e-invoice status (Valid/Cancelled) if present** |
| 4 | GSTR-2B (merged) | 1+ files, whole FY (quarterly blocks), Excel | No | Table 3 summary, B2B purchase invoices + eligibility flag, B2B-CDNR |
| 5 | Electronic Cash Ledger | 1, whole FY, CSV | No | Every cash transaction by tax head |
| 6 | Electronic Credit Ledger | 1, whole FY, CSV | No | Every credit transaction by tax head |
| 7 | Electronic Liability Register | 1, whole FY, CSV | No | Every liability transaction |
| 8 | Outward E-Way-Bill | 1, whole FY, Excel | No | EWB no./date/time, invoice no./date, assessable & tax value, HSN, vehicle |
| 9 | Inward E-Way-Bill | 1, whole FY, Excel | No | Same fields, inward |
| 10 | GST-Prime TPST | 1, whole FY, Excel | No | 12-month self-filing summary (taxable, tax by head, seller/invoice counts) |
| 11 | Portal "Tax liability & ITC comparison" | 1, whole FY, Excel | No | Monthly GSTR-1-vs-3B and 3B-vs-2B comparison, cumulative-shortfall columns |
| 12 | BO / 360° Profile | 1, whole FY (multi-year data), PDF | No | Financial summary, BIFA figures, ITC passed/received, Top-10 lists, related/cancelled-party ITC, DRC payments, Appeal/Case/Transfer sections |

### 3.2 New this round (4 more, all optional)

| # | File | Frequency / format | Contents |
|---|---|---|---|
| 13 | GSTR-9 (Annual Return) | 1, whole FY, PDF (text-based, not scanned) | Table 4 (outward liability), Table 5 (nil/exempt/non-GST), Table 6A (ITC via 3B), Table 9 (tax payable/paid, late fee, interest) |
| 14 | GSTR-9C (Reconciliation Statement) | 1, whole FY, PDF (text-based) | Turnover reconciliation (Table 5/7), ITC reconciliation (Table 12), ARN + ARN Date, tax paid reconciliation (Table 9) |
| 15 | Table 8A | 1, whole FY, Excel (government-standard export — same layout for every taxpayer) | Invoice-level inward supplies with "ITC available = Yes/No" flag + reason, B2BA/CDNR/CDNRA sheets |
| 16 | Balance Sheet + P&L | 1, whole FY, PDF **or hand-typed via `bs_pl_input.py`** | Total Assets/Liabilities, Revenue, Trade Payables/Receivables, Inventories, Fixed Assets, Investments, Provisions, Expenses, Finance Costs, Reserves & Surplus — see §11 for why this is NOT auto-parsed from a scanned PDF |

**Only #1 and #2 (GSTR-1, GSTR-3B) are mandatory.** Every other document degrades gracefully to an explicit SKIP/INFO state if absent — see §7 for the full missing-document behaviour matrix.

---

## 4. Complete Python file catalog (20 files, exact filenames required)

| Script | Role | Status this round |
|---|---|---|
| `merged_period_utils.py` | Marker-row detection — foundation for everything else | unchanged |
| `folder_classifier.py` | Identifies every file by content signature; discovers month coverage; now supports multiple files per document type | **modified** |
| `gst_scrutiny_tool.py` | Raw LEFT\|RIGHT\|DIFF\|MATCH comparison engine (one month at a time); E-Invoice parser incl. cancellation detection | **modified** |
| `gst_analysis_checks.py` | The "14 checks" interpretive layer | unchanged (works correctly now because upstream CONFIG is finally populated — see §6.2) |
| `gst_eway_recon.py` | The 27-check EWB matrix; now degrades honestly when a whole EWB direction is absent | **modified** |
| `gstr2b_parser.py` | Shared GSTR-2B reader; now degrades gracefully instead of crashing when 2B is absent | **modified** |
| `amendments.py` | Amendment sheets (b2ba/cdnra) + Table-13 doc-series integrity | unchanged |
| `run_monthly_pipeline.py` | Runs the three per-month engines together; now also wires filing-compliance and cancelled-e-invoice pass-through | **modified** |
| `annual_sources.py` | Ledgers, TPST, portal-comparison parsers | unchanged |
| `bo_profile_parser.py` | BO/360° Profile PDF parser; section-boundary detection hardened, new `--diagnose` mode | **modified** |
| `build_annual_workbook.py` | FY-wide annual-reconciliation sheets; BIFA lookup de-hardcoded, false-REVIEW-on-missing-data fixed | **modified** |
| `ewb_annual_parser.py` | Whole-FY EWB parser (incl. EWB generation time) | unchanged |
| `gst_unified_scrutiny.py` | Shared Excel sheet-writers, reused by `master_build.py` | unchanged |
| `hsn_fraud_checks.py` | HSN-code-wise + fraud-pattern engine (36 checks) | unchanged |
| `master_build.py` | **Single entry point** — orchestrates everything, writes the output workbook | **heavily modified** (dynamic month-ordering, per-month crash isolation, 5 new safe-parse wrappers, new sheet writers, forensic-layer wiring) |
| `annual_return_parser.py` | **NEW** — parses GSTR-9, GSTR-9C, Table 8A | **new** |
| `filing_compliance.py` | **NEW** — ARN-date extraction, statutory due dates, Sec 47 late fee, Sec 50 interest | **new** |
| `forensic_checks.py` | **NEW** — R13 (turnover-gap), R14 (four-way ITC), R0–R12 (BS/PL rule engine), Cancelled-E-Invoice cross-checks, Doc-Series cancelled-e-invoice enrichment | **new** |
| `bs_pl_input.py` | **NEW** — hand-typed, GSTIN-tagged structured input for the BS/PL rule engine (not OCR — see §11) | **new** |
| `gst_file_finder.py` | Retired — superseded by `folder_classifier.py`; not needed | n/a |

Total: **~9,075 lines** across the 20 active files (was ~8,000 before this round's additions).

---

## 5. Complete check catalog — every comparison this tool performs

### 5.1 Raw Comparison engine (`gst_scrutiny_tool.py`) — per month

| Section | Compares |
|---|---|
| A. Outward Liability | GSTR-1 (net of CN) vs GSTR-3B 3.1(a) — taxable, IGST, CGST, SGST, CESS |
| A (gross ref) | GSTR-1 gross invoice total vs GSTR-1's own HSN summary |
| A2. GSTR-1 internal | Invoice-level (net of CN) vs HSN summary — taxable, IGST, CGST, SGST |
| B. E-Invoice vs GSTR-1 | B2B taxable/IGST/CGST/SGST, unique-invoice count, invoices missing IRN (**cancelled e-invoices now excluded from both sides — see §6.1**) |
| B2. Line-level | Every invoice+rate combination, GSTR-1 vs E-Invoice; flags blank-invoice-number lines |
| C. RCM | GSTR-3B 3.1(d) & 4(A)(3) vs GSTR-2B RCM figures (SKIPPED, not faked, if GSTR-2B absent) |
| D. ITC (All other) | GSTR-3B 4(A)(5) vs GSTR-2B, gross and net-of-CN, by head (SKIPPED if 2B absent) |
| D2. ITC Reversal | GSTR-3B 4(B)(2) vs GSTR-2B's own credit-note figures, by head (SKIPPED if 2B absent) |

### 5.2 Analysis — 14 checks (`gst_analysis_checks.py`) — per month

0 GSTR-1 B2B totals · 1 Nil/exempt/non-GST (GSTR-1 vs 3.1c/3.1e) · 2 Credit-note effect on liability · 3 ITC arithmetic 4C=4A5+4A3−4B1−4B2 · 4 Effective tax-rate · 5 Dropped invoice numbers · 6 Duplicate invoice numbers · 7 E-invoice error flags · **8 IRN-vs-filing lag (now actually populated — was always INFO before this round)** · 9 Rate-wise e-invoice vs HSN · **10 Filing-gap/late fee (now actually populated)** · 11 POS vs GSTIN tax-head · 12 RCM routing · 13 HSN-summary vs named-invoice IGST · 14 ITC/Liability ratio.

### 5.3 E-Way-Bill — 27 checks (`gst_eway_recon.py`) — per month

1 EWB-Out in GSTR-1 · 2 EWB-Out value vs GSTR-1 · 3 EWB-Out unmatched · 4 Inter-state >₹50K no EWB · 5-6 EWB-Out vs E-Invoice · 7 EWB-date vs doc-date gap · 8 EWB-Out aggregate vs 3.1(a) · 9 Tax-type EWB vs GSTR-1 · 10-14 EWB-In vs GSTR-2B (matched/unmatched/threshold/unaccounted) · 15 EWB-Out vs EWB-In · 16 EWB-gen vs doc-date gap · 17 Triangulation · 18 HSN rate-wise · 19 3B ITC vs EWB-In · 20 E-Invoice purchase vs EWB-In · 21 EWB vs filing date · 22 Validity vs supply date · 23 Multiple EWBs/invoice · 24 EWB cancelled after filing · 25-26 Assessable-value ratios · 27 Same vehicle repeated trips.

**Checks #1,3,4,8,9,10,12,13,25,26 now show an honest SKIPPED (not a misleading PASS/REVIEW/FLAG) when the relevant EWB direction was never supplied at all for this taxpayer** — see §6.1.

### 5.4 Cross-month layer

- **Doc-Series Integrity** (`amendments.py` + `forensic_checks.py`): Table-13 declared serial ranges vs actual invoice numbers, per month. Three-tier matching (exact prefix → punctuation-normalized → fuzzy substring), cross-referenced against Table 13's own declared "Cancelled" figure, **and now also against the Cancelled-E-Invoices list** — a "missing" serial that turns out to be a cancelled e-invoice is marked explained, not left as an unexplained red flag.
- **Rectification Pairs**: every b2ba/cdnra amendment row traced back to the earlier month that first reported the original document, plus a DRC-payment cross-reference. Now works across any number of FYs (was hardcoded to one 12-month window).

### 5.5 Annual reconciliation (`build_annual_workbook.py`)

- Cash/Credit/Liability ledgers vs TPST vs portal comparison, monthly, PASS/REVIEW/N-A flagged.
- **FY Total vs BIFA**: full-year totals vs department's own BIFA figures — Liability (GSTR-3B), ITC (both BIFA columns shown side by side), and EWB liability. **FY lookup is now dynamic (was hardcoded to "2022-23"); an unmatched/empty BIFA no longer produces a false REVIEW — the sheet says explicitly why the figures are blank instead.**
- **Related-Party Alerts** / **Top Counterparties**: BO Profile's related/cancelled-party and Top-10 lists. **Now explicitly distinguish "genuinely zero rows" from "this section's parser found nothing" instead of an ambiguous empty sheet.**
- **Departmental proceedings** (Annual Cover sheet): BO Profile's Appeal/Case/Transfer sections.

### 5.6 HSN-code-wise + Fraud-Pattern checks (`hsn_fraud_checks.py`) — FY-wide, 36 checks

Categories A (HSN-only: wrong rate, exempt-charged, cess-missing, multi-rate, invalid HSN), B (POS/state: wrong tax head, B2C-Large no-EWB, SEZ misclassification, GSTR-2B "ITC Availability=No"), C (combined: branch-transfer, invoice-vs-EWB interstate mismatch), plus a 25+-item named fraud-pattern list (round-number invoices, reciprocal trading, CN timing, HSN drift, year-end dumping, ghost-supplier PAN clusters, IRN delay, midnight EWB, EWB bursts, credit hoarding, cash-timing patterns, rate outliers, etc.) — unchanged this round.

### 5.7 NEW — Forensic layer (`forensic_checks.py`, `filing_compliance.py`, `annual_return_parser.py`)

| Ref | Check | Needs |
|---|---|---|
| **R13** | Turnover-gap rule: GSTR-9C's exempt/nil/non-GST turnover adjustment cross-checked against GSTR-1 Table 8's own actual data rows — flags an "unsupported exempt turnover" if the adjustment has no supporting declaration anywhere | GSTR-9C + GSTR-1 (zero *extra* documents beyond what's already in the 12-file set) |
| **R14** | Four-way ITC reconciliation: GSTR-3B (via GSTR-9 Table 6A) vs GSTR-2B (FY total) vs Table 8A (invoice-level, "Yes" only) vs Books (GSTR-9C Table 12A) — flags a material, one-directional books-vs-returns gap distinctly from a return-internal mismatch | Any 2 of {GSTR-9, GSTR-2B, Table 8A, GSTR-9C} |
| **R0** | Balance Sheet self-balances (pre-flight gate; halts R1-R12 if it fails) | BS/PL structured input |
| **R1** | Revenue from Operations vs GSTR-9C Table 5 | BS/PL + GSTR-9C |
| **R2–R11** | Other Income, Trade Payables (Rule 37), Trade Receivables (bad-debt), Inventories (17(5)(h)), Fixed Assets (18(6)/16(3)), Investments (Schedule I), Provisions (undisclosed tax exposure, cross-checked against BO Profile DRC), Other Expenses (17(5) blocked credits), Finance Costs, Share Capital — each degrades to a specific "not tested — X not supplied" INFO if its BS/PL line item is missing | BS/PL structured input (some also use BO Profile) |
| **R12** | Reserves & Surplus roll-forward (Opening + Net Profit = Closing) | BS/PL |
| **Filing Compliance** | Per-month ARN date (GSTR-1 + GSTR-3B), statutory due date, Section 47 late fee (with Notification 07/2023-CT turnover caps), Section 50 interest | GSTR-1 + GSTR-3B (their own marker/sheet data — no extra document) |
| **Cancelled E-Invoices** | Aggregated list + two cross-checks: (a) cancelled invoice still reported in GSTR-1 B2B, (b) cancelled invoice with a still-live outward EWB | E-Invoice file with a recognisable status column |

---

## 6. What changed and why — the headline fixes

### 6.1 Cancelled E-Invoices (built this round, then bug-fixed after QA)
Detects a cancellation-status column in the E-Invoice export (case-insensitive match against several real-world header-name variants), excludes cancelled rows from every E-Invoice-vs-GSTR-1 comparison (a cancelled invoice is correctly absent from GSTR-1, so including it manufactured false gaps — confirmed and fixed: 6 real cancelled invoices totalling ₹44,12,291 were producing 6 false "LINE-LEVEL GAP" mismatches before the fix), and cross-references cancelled invoices against Doc-Series "missing" serials.

### 6.2 ARN Date / Filing-Gap / Late Fee (built this round, then bug-fixed after QA)
Root cause of the original v1 gap: the ARN-date extraction function lived inside a legacy code path `master_build.py` never called, so `GSTR1_FILING_DATE`/`GSTR3B_FILING_DATE` were always `None` and checks #8/#10 always showed INFO. Rebuilt as `filing_compliance.py`, properly wired per month. A follow-up bug (the regex only recognised "Date ... Filing/ARN" word order, not the real file's "ARN Date" word order) was found and fixed against the exact real marker text.

### 6.3 Graceful degradation for missing documents
- GSTR-2B entirely absent for a month: previously crashed the whole run; now shows SKIPPED for the affected comparison sections only.
- An entire EWB direction never supplied: previously produced misleading PASS/REVIEW/FLAG (false "0/0 matched" clean results, or false floods of "missing EWB" on every large invoice); now shows honest SKIPPED.
- Annual-level sources (ledgers, TPST, portal-comparison, BO-Profile) individually absent: previously crashed the whole run; now each degrades independently.
- One month's unexpected parsing error: previously could crash a multi-month run; now that month is skipped and logged, every other month still processes.

### 6.4 5-year / multi-taxpayer architecture
`folder_classifier.py` now collects every matching file per document type (was: silently kept only the last match, discarding earlier years). `master_build.py`'s month-ordering is now computed dynamically from the actual data instead of a hardcoded 12-month list. See §10 for exactly how much of this is proven vs still needs a real multi-year test.

### 6.5 Balance-Sheet/P&L rule engine actually wired in
Built and unit-tested in isolation, but never actually called from `master_build.py`'s main pipeline until this round — fixed, plus a new safety check (§8) so a stale/wrong-taxpayer's BS/PL data can never be silently applied.

---

## 7. Missing-document behaviour matrix (verified this round)

| Document missing | Result |
|---|---|
| E-Invoice | Comparison Sections B/B2 skipped; Analysis check #8 INFO; EWB checks #5-7 INFO; Cancelled-E-Invoices sheet shows "not supplied". Everything else runs normally. |
| GSTR-2B (whole month) | Comparison Sections C/D/D2 show one SKIPPED row each (not a crash, not a fake zero-compare); EWB checks #10-13,#26 show SKIPPED/INFO depending on whether 2B was PDF-only or fully absent. |
| Outward EWB (whole FY) | EWB checks #1,3,4,8,9,25 show SKIPPED with an explicit reason. |
| Inward EWB (whole FY) | EWB checks #10,12,13 show SKIPPED. |
| GSTR-9 | R14 loses its "3B" source; runs with the remaining sources if ≥2 available, else INFO. |
| GSTR-9C | R13 shows INFO immediately; R14 loses its "Books" source; BS/PL R1 shows INFO. |
| Table 8A | R14 loses its "8A" source. |
| Balance Sheet / P&L | R0-R12 show one INFO row: "bs_pl_input.py not found/empty". |
| Cash/Credit/Liability Ledger, TPST, Portal-Comparison, BO-Profile (any one) | That specific annual-reconciliation sheet shows N/A for the affected figures instead of crashing the run. |

**Only GSTR-1 and GSTR-3B are mandatory.** Every other document is optional and independently gracefully-degrading — confirmed by isolated unit tests for each (see the session transcript / prior turns for exact test code and output).

---

## 8. Any-taxpayer readiness — what's generic vs what needs manual updating

**Fully generic, no editing needed for a new taxpayer:** self-GSTIN/company-name detection, FY/month detection, output filename, all 12 original document parsers, ARN/filing-compliance, cancelled-e-invoice detection, all 27+14+8-section monthly checks, Doc-Series/Rectification-Pairs, annual reconciliation, the 36 HSN/fraud checks' *structure* (though see below), R13/R14.

**Needs manual updating per taxpayer:**
- **`bs_pl_input.py`** — hand-typed Balance Sheet/P&L figures (see §11 for why). **Safety-tagged with a `_gstin` field**: `master_build.py` checks this against the taxpayer actually being processed and refuses to use the data (with a printed warning) if it doesn't match — this was a real risk caught and fixed this round (a stale taxpayer's BS/PL figures could otherwise have been silently reused for a different taxpayer's run).
- **`hsn_fraud_checks.py`'s `HSN_RATE_MASTER`** — curated from this taxpayer's actual HSN codes (pharma). A different industry's HSN codes aren't in it; check A1 (wrong-rate-vs-master) simply won't have anything to compare an unlisted HSN code against (degrades to not-flagged, not a crash or false flag) until the master is extended.

---

## 9. Known limitations (honest, consolidated)

- **BO Profile sheets (BIFA, Related-Party Alerts, Top Counterparties)**: hardened against the most likely causes (case/whitespace-tolerant section-marker matching, a joined-line fallback for split headers, relaxed numeric-token thresholds), and now never show a *misleading* result if parsing still fails (explicit "why this might be incomplete" notes replace what used to be a silent 0 or an ambiguous empty sheet) — but the underlying section-detection was not re-verified against a real BO Profile PDF this round (none was supplied). A new `bo_profile_parser.py --diagnose <path>` command exists specifically to close this out precisely next round.
- **QRMP due-date category (Category X/Y state list)** in `filing_compliance.py` was written from general knowledge, not cross-verified against the exact current CBIC notification — low-risk for this taxpayer (confirmed monthly filer, not QRMP, so this code path isn't exercised), but flagged for any future QRMP taxpayer.
- **Multi-year (5-FY) support**: file-discovery and month-ordering are generalized and reasoned through carefully, but not run end-to-end against a real multi-year folder (none was supplied). Three sub-layers still resolve to the FIRST matched file per document type only: the 36-check HSN/Fraud layer, the annual-reconciliation layer (ledgers/TPST/portal/BO-profile), and the GSTR-9/9C/Table-8A forensic layer — each would need per-FY looping to be true multi-year (contained follow-up, not a rebuild). **This entire limitation is irrelevant for single-FY use**, which is the current stated use case.
- **Balance Sheet/P&L**: never auto-OCR'd from a scanned PDF (tested and found to silently misread digits — a ₹6,00,000 error on Finance Costs alone). Structured hand-typed input only (`bs_pl_input.py`), by design — see §11.

---

## 10. Multi-year status in one line
Architecturally built and reasoned through; not run end-to-end against real multi-year data this round (only single-FY data was available to test against). Safe to use for single-FY today; treat 5-year output as unverified until tested against a real 2+-FY folder.

---

## 11. Balance Sheet / P&L — why it's hand-typed, not OCR'd

The real BS/P&L PDF for this taxpayer is a **scanned image** (confirmed: zero extractable text, "Scanned with CamScanner" watermark). OCR (`pytesseract`) was tested and found to silently misread real digits (Finance Costs ₹49,73,007.06 → OCR'd as ₹43,73,007.05, a ₹6 lakh error that would itself trigger a false forensic FLAG). A scrutiny tool cannot silently trust a guessed digit. `bs_pl_input.py` is therefore a plain Python dict, hand-transcribed once by a human reading the real document, cross-validated because the figures independently reconcile exactly (R0 Balance Sheet self-balances, R1 revenue-vs-GSTR-9C, R12 reserves-roll-forward all PASS with zero variance against numbers computed a completely different way). If a text-based (non-scanned) BS/P&L export becomes available for any taxpayer, the same watermark-tolerant PDF-text approach already proven on GSTR-9/GSTR-9C could be extended to it directly.

---

## 12. Output workbook structure

| Sheet(s) | Content |
|---|---|
| Master Dashboard | Every FLAG/MISMATCH/REVIEW across ALL layers (monthly + HSN/Fraud + **Forensic R13/R14 + Cancelled E-Invoices, newly wired in**), ranked together; header shows real FY-span and within-span gaps |
| `<Month> Comparison/Analysis14/EWB/EWB Detail` ×N months | Per-month engine output |
| Doc-Series Integrity | Table-13 vs actual, 3-tier matched, fuzzy-match + declared-cancellation + **cancelled-e-invoice explanations** |
| Rectification Pairs | Amendments linked to original month + DRC payments |
| HSN & Fraud Pattern Checks | All 36 checks' findings |
| **Filing Compliance & Late Fee** *(new)* | Per-month ARN, filed date, due date, Rs late fee + interest for GSTR-1 and GSTR-3B |
| **Forensic Checks (R13-R14)** *(new)* | Turnover-gap rule + four-way ITC reconciliation |
| **Cancelled E-Invoices** *(new)* | Cross-check findings + full list of cancelled invoices with month/rate/taxable/IGST/IRN/cancel-date |
| Annual Cover & Caveats | Sources, known limitations, departmental proceedings |
| Annual Ledger Walkthrough | Monthly ledger/TPST/portal reconciliation |
| FY Total vs BIFA | FY totals vs department's BIFA figures (**dynamic FY, honest blank-vs-mismatch distinction**) |
| Related-Party Alerts / Top Counterparties | BO Profile lists (**honest zero-rows note**) |

---

## 13. Verified results (exact numbers reproduced from real files, this round)

- **Table 8A**: 2,534 total B2B rows, 2,408 "ITC available = Yes". CGST ₹97,11,690.70, SGST ₹97,11,690.70, IGST ₹2,81,96,772.63, Total ₹4,76,20,154.03.
- **GSTR-9**: Table 6A (ITC via 3B) total ₹4,78,88,427.54. Table 4B (B2B outward) taxable ₹29,98,31,494.95. Table 9 liability IGST ₹2,90,78,901.00.
- **GSTR-9C**: Turnover ₹46,56,39,087.14. Exempt/nil/non-GST adjustment ₹16,70,77,200.00. ITC per books ₹4,68,38,857.26. ARN `AA0503234208977`, ARN Date `30-12-2023`.
- **R14 (four-way ITC)**: 3B−Books = ₹10,49,570.28; 2B−Books = ₹9,39,250.40; 8A−Books = ₹7,81,296.77 — all exact matches to the independently-produced Forensic Framework document.
- **R13 (turnover-gap)**: correctly FLAGs the ₹16.7 Cr unsupported-exempt-turnover scenario.
- **R0/R1/R12 (BS/PL)**: all PASS with exact-match figures (Balance Sheet self-balances both FYs; Revenue exactly ties to GSTR-9C; Reserves & Surplus rolls forward to the paisa).
- **Bug-fix verifications**: Cancelled-E-Invoice header detection + total-exclusion (synthetic test using the real reported header name); ARN-date extraction (synthetic test using the real reported marker text); GSTR-2B/EWB graceful-degradation (9+ isolated scenario tests, all passing); BIFA empty-dict no-longer-produces-false-REVIEW (direct test).

---

## 14. Session history (consolidated)

1. **Original build** (prior sessions) — single-month tool (6 scripts) → scaled to full-FY multi-month master pipeline → merged-file migration for whole-FY workbooks.
2. **QA review + rectification** (prior session) — 5 confirmed bugs fixed (GSTR-3B Table-4 label collision, doc-series punctuation handling, Dashboard label split, BIFA dual-column transparency, EWB-vs-BIFA annual check).
3. **HSN & fraud-pattern layer** (prior session) — 36 checks built, curated rate master, POS/state-code bug caught and fixed pre-ship.
4. **Forensic Comparison Framework** (this project's session, turn 1) — independent forensic analysis produced from GSTR-9/9C/8A/BS-PL, identifying the four-way ITC gap, turnover-gap, and the generic BS/PL rule framework (R0-R14).
5. **v2 build** (this session, turn 2) — implemented graceful degradation for all optional documents, 5-year architecture foundation, Cancelled E-Invoices, Filing Compliance/Late Fee, and the Forensic layer (R13/R14/R0-R12) — all built and, where real files existed, verified to exact-match figures.
6. **Dashboard wiring fix** (this session, turn 3) — Forensic and Cancelled-E-Invoice findings wired into the Master Dashboard's ranked list (previously only visible on their own sheets).
7. **QA bug-fix round** (this session, turn 4) — 5 bugs reported by the validation team: Cancelled-E-Invoice header case-sensitivity + total-pollution (fixed, verified), ARN-date word-order (fixed, verified), BIFA/Related-Party/Top-Counterparty parsing (hardened, not fully verified — no real BO Profile PDF available).
8. **Genericity audit** (this session, turn 5) — found and fixed: BS/PL rule engine was never actually wired into the pipeline; a stale-taxpayer BS/PL-data risk if `bs_pl_input.py` isn't updated per taxpayer (now GSTIN-tagged and checked).

---

## 15. What the validation team should focus on next

In priority order, to move the remaining "hardened but not verified" items to "verified":
1. **Real BO Profile PDF** (or `python3 bo_profile_parser.py --diagnose <path>` output) — closes out the BIFA/Related-Party/Top-Counterparty parsing with the same certainty as the two bugs already fixed this round.
2. **A second FY's worth of merged files** (any FY) — the direct way to prove the 5-year architecture end-to-end rather than by code review alone.
3. **Real Cash/Credit/Liability Ledger CSVs + TPST + Portal Comparison** — re-verify the annual-reconciliation layer against the new safe-parse wrappers under real data, not just synthetic/absent-file tests.
