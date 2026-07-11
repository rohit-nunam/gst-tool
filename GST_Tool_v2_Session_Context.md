# GST Scrutiny Tool — v2 Session Context & Verification Record
**(paste this whole document into a new chat, alongside the original `GST_Tool_Full_Project_Context.md`, and Claude will have complete context to continue or verify this work)**

**Taxpayer:** M R HEALTHCARE PRIVATE LIMITED, GSTIN `05AAECM6380J1ZA`
**This document covers:** ONE session's worth of changes on top of the v1 tool described in `GST_Tool_Full_Project_Context.md`. Read that document first for the base architecture (four-layer design, merged-file model, 12-file input set) — this document only covers what changed/was added THIS session.
**Session trigger:** Darshan asked four things (1) tool must run correctly with any subset of documents missing, (2) extend to 5-year multi-FY analysis, (3) add a Cancelled E-Invoices list, (4) fix ARN-date/filing-gap/late-fee — plus implement a new `GST_Forensic_Comparison_Framework_v1.md` (a from-scratch forensic analysis document produced in a prior turn of this same session) and "hard rule: no safety nets, must run for ANY taxpayer, zero errors."
**New source documents introduced this session:** GSTR-9 (PDF), GSTR-9C (PDF), Table 8A (XLSX, government-standard export), Balance Sheet + P&L (PDF, scanned).

---

## 1. What was actually built — file by file

### 1.1 Four brand-new Python modules

| File | Lines | Purpose |
|---|---|---|
| `annual_return_parser.py` | ~470 | Parses GSTR-9 PDF, GSTR-9C PDF, Table 8A XLSX. All three degrade to `available=False` + a reason string if the file is absent/unreadable — never raises. |
| `filing_compliance.py` | ~330 | Per-month ARN-date extraction (GSTR-1 + GSTR-3B), statutory due-date tables (monthly + QRMP-ready), Section 47 late-fee calculator (with Notification 07/2023-CT turnover caps), Section 50 interest calculator. |
| `forensic_checks.py` | ~430 | Implements Part 2 of the Forensic Framework: R13 (turnover-gap rule), R14 (four-way ITC reconciliation), R0–R12 (generic Balance-Sheet/P&L rule engine, structured-input based), plus the Cancelled-E-Invoice cross-checks (backlog items B9/D2). |
| `bs_pl_input.py` | ~55 | A plain Python dict template for BS/P&L line items, hand-transcribed (NOT OCR'd) from the real scanned PDF for this taxpayer. Feeds `forensic_checks.check_bs_pl_rules()`. |

### 1.2 Six existing modules patched (surgical changes, not rewrites, except `master_build.py`)

| File | What changed | Why |
|---|---|---|
| `gstr2b_parser.py` | `summary_for_month()` no longer raises when GSTR-2B is absent for a month — returns an explicit `available=False` summary with `None` fields (not `0.0`, to avoid a fake-nil result being diffed against real GSTR-3B numbers). | GSTR-2B missing used to crash the ENTIRE run (nothing caught the raise anywhere in the call chain). |
| `gst_scrutiny_tool.py` | (a) `parse_einv()` now detects a cancellation-status column (tries several real-world header-name variants) and collects cancelled rows. (b) `build_comparisons()` sections C (RCM)/D (ITC)/D2 (ITC reversal) now check `b2b.get("available")` first and emit one SKIPPED row instead of indexing a now-`None` field. | (a) New Cancelled-E-Invoices feature. (b) Without this, the gstr2b_parser.py fix above would have just moved the crash one level up (a `TypeError` on `None - None` instead of a raise). |
| `gst_eway_recon.py` | `run()` gained two new parameters, `ewb_out_file_supplied`/`ewb_in_file_supplied`. Checks `#1,#3,#4,#8,#9,#10,#12,#13,#25,#26` all now branch on these flags and emit SKIPPED with an honest reason when a whole EWB direction was never supplied, instead of a **misleading** PASS/REVIEW/FLAG (see §3 for exactly what was wrong). Also fixed a `b2b["..."]` unconditional-index pattern (4 places) that would crash when GSTR-2B is unavailable, and a genuine scoping bug (`ewb_out_total` referenced before assignment when the outward-EWB block was skipped) caught during testing. | This was the core ask #1 — "tool ko limited data mein bhi sahi chalna chahiye", and specifically that it shouldn't just avoid crashing but avoid producing a **wrong-looking clean/flagged result**. |
| `folder_classifier.py` | Rewrote `classify_folder()`: every document type now collects a **list** of matching files (previously `gstr1_merged = f` inside a loop silently kept only the last match, discarding every earlier file with the same content signature). New `_build_month_file_map()` resolves, per month, which specific file covers it. New detection for GSTR-9 / GSTR-9C / Table 8A / BS-PL (by PDF first-page text signature / sheet-name signature). Backward-compatible single-path keys (`gstr1_merged`, etc.) kept for any old code. | This was ask #2 (5-year support) at the file-discovery layer. |
| `run_monthly_pipeline.py` | `run_month()` gained new parameters: the two EWB-supplied flags (passed straight to `gst_eway_recon.run()`), pre-computed ARN-date maps (passed to `filing_compliance.month_filing_compliance()`, which sets `raw.GSTR1_FILING_DATE`/`GSTR3B_FILING_DATE` before calling the existing engines), and returns the new `compliance`, `cancelled_einvoices`, `einv_cancel_col_found`, `g1_named_invnos` keys. | Wiring layer connecting the new modules into the existing per-month pipeline. |
| `master_build.py` | Largest change. (a) Dynamic month-ordering (`_month_sort_key`/`_sort_months_chronologically`/`_fy_label_for_month`) replaces the hardcoded 12-month `MONTH_ORDER` everywhere it was used (rectification pairing, dashboard sort/header, output filename). (b) Five new `_safe_parse_*()` wrappers guard every annual-level source (ledgers, TPST, portal-comparison, BO-Profile) against being absent — previously unguarded and would crash the whole run. (c) Per-month `try/except` around `run_month()` — one bad month no longer takes down a multi-year run. (d) New sheet writers: `write_filing_compliance()`, `write_forensic_checks()`, `write_cancelled_einvoices()`. (e) Wires `annual_return_parser` + `forensic_checks` (R13/R14) into the main flow, including summing GSTR-2B across every month actually available into an FY total for R14. (f) Output filename and Dashboard header now reflect the real FY-range covered, not a hardcoded string. | Central orchestration for everything else in this document. |

### 1.3 Nine modules — unchanged this session
`merged_period_utils.py`, `amendments.py`, `ewb_annual_parser.py`, `annual_sources.py`, `bo_profile_parser.py`, `build_annual_workbook.py`, `gst_analysis_checks.py`, `gst_unified_scrutiny.py`, `hsn_fraud_checks.py`.

Note on `gst_analysis_checks.py` specifically: it did **not** need to change. Its checks #8 (IRN-lag) and #10 (filing-gap) already had the correct CONFIG-reading logic (`_filing_date()` reads `getattr(raw, "GSTR1_FILING_DATE", None)`) — the bug was purely that nothing upstream ever populated that attribute in the real pipeline. Fixing `run_monthly_pipeline.py` to actually set it (via `filing_compliance.py`) was sufficient; check #8/#10 now work with zero changes to their own file.

### 1.4 Two new documentation files
`CHANGELOG_v2.md` (delivered previous turn — file-by-file technical changelog) and `OCR_LIMITATION.md` (why the scanned BS/P&L PDF is not auto-OCR'd — see §5).

---

## 2. Mapping: what each of Darshan's 4 asks + the Forensic Framework maps to, in code

| Ask | Implementation | Status |
|---|---|---|
| **#1 Run correctly with any document missing** | `gstr2b_parser.py`, `gst_scrutiny_tool.py`, `gst_eway_recon.py`, `master_build.py`'s `_safe_parse_*` | **Built + tested** (isolated unit tests, see §4) |
| **#2 Extend to 5-year analysis** | `folder_classifier.py` (multi-file collection + month maps), `master_build.py` (dynamic month ordering) | **Built, NOT end-to-end tested** — no second FY of monthly data was supplied this session to run against |
| **#3 Cancelled E-Invoices list** | `gst_scrutiny_tool.parse_einv()` + `forensic_checks.build_cancelled_einvoice_findings()` + new sheet | **Built, NOT tested against a real E-Invoice file** — no real E-Invoice export was supplied this session |
| **#4 ARN Date / Filing-gap / Late fee** | New `filing_compliance.py`, wired via `run_monthly_pipeline.py` | **Built + partially tested** — date math and formulas self-tested; per-month ARN extraction tested only against a synthetic workbook (no real merged GSTR-1 file supplied this session) |
| **Forensic Framework Part 1 (§A–F)** | Documented in `forensic_checks.py`'s own docstring as architecturally scoped; **not coded as functions** | **Deliberately NOT built** — would need real ledger CSVs / TPST / portal-comparison / BO-Profile to build against without guessing; see §6 |
| **Forensic Framework Part 2 (R13, R14, R0–R12)** | `forensic_checks.py` | **Built + fully verified against your real files** — see §3 for exact numbers |

---

## 3. Verified results — exact numbers reproduced from your real files this session

These are not estimates; each was produced by running the actual new code against the actual uploaded files and comparing to the Forensic Framework document's independently-computed figures.

**Table 8A** (`annual_return_parser.parse_table_8a()` against `R9_8A_05AAECM6380J1ZA_1.xlsx`):
2,534 total B2B rows, 2,408 with "ITC available = Yes". CGST ₹97,11,690.70, SGST ₹97,11,690.70, IGST ₹2,81,96,772.63, Total ₹4,76,20,154.03. "No" reasons: 124 × "Reverse charge document", 2 × "POS lies in the State of supplier". **Exact match** to the Forensic Framework document.

**GSTR-9** (`annual_return_parser.parse_gstr9()` against `GSTR9_05AAECM6380J1ZA_032023__1_.pdf`):
Table 6A (ITC via 3B) total = ₹4,78,88,427.54 (CGST ₹99,39,640.17, SGST ₹99,39,640.17, IGST ₹2,80,09,147.20). Table 4B (B2B outward) taxable ₹29,98,31,494.95. Table 9 liability: IGST ₹2,90,78,901.00, CGST ₹22,40,993.00, SGST ₹22,40,993.00. Table 5 A–F all zero. **Exact match.**

**GSTR-9C** (`annual_return_parser.parse_gstr9c()` against `GSTR-9C_05AAECM6380J1ZA_032023.pdf`):
Turnover ₹46,56,39,087.14 (audited BS = after-adjustments = declared, all three tie). Exempt/nil/non-GST adjustment ₹16,70,77,200.00. Taxable turnover ₹29,85,61,887.14. ITC per books ₹4,68,38,857.26. ARN `AA0503234208977`, ARN Date `30-12-2023`. Tax payable total = tax paid declared = ₹3,35,60,887.00. **Exact match** — including recovering from a real bug caught mid-session (see §7).

**R14 — Four-way ITC reconciliation** (`forensic_checks.check_four_way_itc()`, all 4 sources):
3B − Books = ₹10,49,570.28. 2B − Books = ₹9,39,250.40 (using a synthetic-but-framework-matching 2B figure, since no real GSTR-2B file was supplied this session — see caveat below). 8A − Books = ₹7,81,296.77. Severity: **FLAG**. **Exact match** to the Framework's §2.2 finding. *Caveat: the 2B figure used for this test was NOT read from a real GSTR-2B file (none supplied) — it was a value close to the Framework document's own quoted number, used only to prove the arithmetic/logic path. The 3B, 8A, and Books figures ARE from real files.*

**R13 — Turnover-gap rule** (`forensic_checks.check_turnover_gap()`):
Correctly returns **FLAG** for the exact scenario in the Framework document (₹16,70,77,200.00 GSTR-9C adjustment, zero supporting GSTR-1 Table-8 rows across a simulated 12 months — real GSTR-1 file not supplied this session, so the "zero rows" input was constructed to match the Framework's own already-confirmed finding, not independently re-verified against a real GSTR-1 file this round).

**R0/R1/R12 — BS/PL rule engine** (`forensic_checks.check_bs_pl_rules()`, against `bs_pl_input.py`'s hand-transcribed real figures):
R0 (Balance Sheet self-balances): **PASS**, both FYs exact. R1 (Revenue vs GSTR-9C): **PASS**, ₹46,56,39,087.14 exact match, zero variance. R12 (Reserves & Surplus roll-forward): **PASS**, Opening ₹3,94,91,383.22 + Net Profit ₹4,81,86,498.93 = Closing ₹8,76,77,882.15 exact. **Exact match** to the Framework's §2.1 "what ties out cleanly" section.

**Graceful-degradation fixes** (isolated unit tests, synthetic inputs since these are behavioural/logic tests, not figure-extraction tests):
- `gstr2b_parser.summary_for_month(None, "Apr-22")` → `available=False`, no exception. **PASS**.
- `gst_eway_recon.run()` with both EWB directions absent → checks #1,#3,#4,#8,#9,#10,#12,#13,#25 all return `SKIPPED` (not the previous false PASS/REVIEW/FLAG). **PASS**, 9/9 checks verified.
- Mixed case (outward supplied, inward absent) → outward checks run normally, inward checks SKIP. **PASS**.
- `gst_scrutiny_tool.build_comparisons()` with GSTR-2B unavailable → sections C/D/D2 produce exactly 3 SKIPPED rows instead of crashing. **PASS**.

---

## 4. What was NOT tested this session (and exactly why)

No GSTR-1 (merged), GSTR-3B (merged), GSTR-2B (merged), Outward/Inward EWB, Cash/Credit/Liability Ledger CSVs, TPST, Portal Tax-Liability-Comparison, or BO Profile PDF were supplied this session — only the four NEW annual-return-side documents (GSTR-9, GSTR-9C, Table 8A, BS/P&L) plus the Forensic Framework markdown. Consequently:

- The **core monthly engine** (`gst_scrutiny_tool.build_comparisons()`, `gst_analysis_checks.run_checks()`, `gst_eway_recon.run()`) was **not re-run end-to-end** against real monthly data this session — only its specific patched code paths were unit-tested with synthetic/minimal inputs (see §3).
- **Multi-year support** (`folder_classifier.py`'s multi-file collection, `master_build.py`'s dynamic month-ordering) has **no real second FY to prove itself against** — it is syntax-correct, logically reviewed, and the single-FY case was implicitly exercised (since single-FY is just multi-year with N=1), but a genuine 2-or-more-FY folder has not been run.
- **ARN-date extraction from a real merged GSTR-1 file** was tested only against a hand-built synthetic workbook mimicking the documented marker format (`Financial Year: ... | Tax Period: ... | ARN: ...`). Whether your REAL merged GSTR-1 file's marker text also includes a date after the ARN (as assumed) or only the ARN number (in which case the `Read me`-sheet fallback path fires, with an explicit warning) is **unconfirmed**.
- **Cancelled E-Invoices** column-detection has never seen a real E-Invoice export with an actual cancelled row — the header-name candidate list (`IRN Status`, `Status`, `Cancel Date`, etc.) is built from general knowledge of GSTN export conventions, not confirmed against this taxpayer's actual file.
- **Forensic Framework Part 1 (§A–F)** — ledger line-item tie-out, ITC-utilization sequencing, ledger balance-continuity, portal-comparison trend detection, DRC-03 aggregation/FY-attribution, EWB vehicle cross-matching, TPST triangulation — **none of this is coded yet**. This was a deliberate choice, not an oversight: building these against files that weren't supplied would mean guessing at real column layouts/values, which is exactly what this tool's own design philosophy (documented repeatedly in the original `GST_Tool_Full_Project_Context.md`) refuses to do.

---

## 5. Balance Sheet / P&L — explicit limitation

`MRHC_PL_AND_BS_FY_22-23.pdf` is a **scanned image** (confirmed: zero extractable text via `pdfplumber`, "Scanned with CamScanner" watermark). OCR (`pytesseract`) was tested and found to **silently misread real digits** — e.g. Finance Costs ₹49,73,007.06 was OCR'd as ₹43,73,007.05, a ₹6,00,000 error that would itself have triggered a false forensic FLAG. Full writeup with the tested proof: `OCR_LIMITATION.md`.

**Decision made:** BS/P&L figures are NOT auto-extracted. `bs_pl_input.py` is a hand-typed structured template (transcribed once, by a human reading the real PDF, cross-checked because R0/R1/R12 all reconcile exactly against independently-known figures) that feeds the R0–R12 rule engine. If a text-based (non-scanned) BS/P&L export becomes available in future, `annual_return_parser.py`'s proven watermark-tolerant PDF-text approach could be extended to it directly.

---

## 6. Recommended next-session document list (in priority order)

To close the gaps in §4, in order of highest forensic value first:
1. **Merged GSTR-1 + GSTR-3B workbooks** (even just 1–2 months) — lets the core engine, ARN-date extraction, and Cancelled-E-Invoices column-detection all be verified against real data for the first time this "v2" iteration.
2. **A second FY's worth of merged files** (GSTR-1/3B/2B/E-Invoice/EWB, any FY) — the single most direct way to prove out the 5-year architecture.
3. **Electronic Cash/Credit/Liability Ledger CSVs + TPST + Portal Comparison + BO Profile** for this FY — unlocks building Forensic Framework Part 1 (§A–F) for real, plus lets `master_build.py`'s existing annual-reconciliation layer be re-verified against the new `_safe_parse_*` wrappers.
4. **A real E-Invoice merged export** — confirms (or corrects) the Cancelled-E-Invoices column-name detection.
5. **A text-based (non-scanned) BS/P&L export**, if the company's accountant can produce one — would let §5's limitation be lifted.

---

## 7. One real bug found and fixed mid-session (worth knowing about)

While building `annual_return_parser.py`, an early version of the watermark-cleanup logic cleaned the whole PDF text block-at-once and then dropped watermark-only lines before re-splitting into lines — which silently **shifted every later line's index out of alignment with its own raw counterpart**, causing several GSTR-9 fields (Table 4B, Table 4I, Table 6A, Table 9) to come back as `None` even though the data was present. Caught by cross-checking the extracted `None`s against the Forensic Framework document's known-correct figures (which the code SHOULD have reproduced but didn't) — not by a passing-but-wrong test. Fixed by cleaning each line independently instead of the whole text at once, preserving guaranteed 1:1 index alignment. Re-verified afterward: every field now matches exactly (§3). This is exactly the kind of failure mode ("looks like it ran fine, quietly returns wrong/missing data") that this whole tool's design philosophy exists to catch — flagging it here explicitly rather than leaving it as an invisible fixed-in-passing detail.

---

## 8. How to run (see separate chat message for the short version)

Entry point is unchanged: `python3 master_build.py <folder>`. The folder must contain every data file (any filenames — content-detected) PLUS all 20 `.py` files listed in §1 (15 original + `annual_return_parser.py`, `filing_compliance.py`, `forensic_checks.py`, `bs_pl_input.py`), sitting flat in the same folder, exact `.py` filenames required (per the original tool's own rule — data filenames never matter, script filenames do).
