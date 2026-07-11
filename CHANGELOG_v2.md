# GST Scrutiny Tool -- v2 Changelog

Everything below was written against the real 15 `.py` files + tested, where
real data existed, against the actual uploaded documents (GSTR-9, GSTR-9C,
Table 8A, Balance Sheet/P&L) for M R HEALTHCARE. **No GSTR-1/3B/2B/EWB/
ledger/TPST/portal-comparison/BO-Profile files were supplied this round** --
those pieces are built correctly per the existing codebase's own patterns
and syntax/import-checked, but could not be run end-to-end against real
monthly data. Marked clearly below, module by module.

---

## 1. Fixed: tool now runs correctly with ANY subset of documents missing

**Already safe before this session** (verified, unchanged): E-Invoice file
entirely absent -- `gst_scrutiny_tool.parse_einv()` / `gst_analysis_checks
.read_einv_lines()` already degraded cleanly (`available=False`).

**Fixed this session** (tested):
- **GSTR-2B entirely absent for a month** used to `raise PeriodParseError`
  with nothing catching it anywhere in the call chain -- crashed the WHOLE
  RUN, not just that month. Now `gstr2b_parser.summary_for_month()` returns
  an explicit `available=False` summary; `gst_scrutiny_tool.build_comparisons()`
  (sections C/D/D2) and `gst_eway_recon.run()` (checks #10-13, #26) both now
  check that flag first and emit one clear SKIPPED row instead of either
  crashing or silently comparing against a fake zero. **Tested**: see
  `Test 1` and the `build_comparisons()` test in this session's transcript --
  3 SKIPPED rows appear exactly where C/D/D2 used to crash.
- **An entire EWB direction (Outward or Inward) never supplied at all**
  (common for smaller taxpayers under the Rule-138 threshold, or pure-
  service businesses) used to produce **misleading** results, not a crash:
  check #1 showed a false PASS ("0/0 matched"), check #4 flooded every
  large inter-state invoice as REVIEW, check #10 showed a false REVIEW,
  check #12 showed a false PASS, check #25 could show a false FLAG. All
  five (`#1,#3,#4,#8,#9,#10,#12,#13,#25,#26`) now check a new
  `ewb_out_file_supplied`/`ewb_in_file_supplied` flag (computed once in
  `master_build.py` from whether any matching annual workbook was found at
  all, independent of a specific month having zero rows -- a real, separate,
  legitimate state) and emit SKIPPED with an honest reason instead.
  **Tested**: 4 isolated test scenarios (both absent / mixed / normal), all
  passing, in this session's transcript.
- **Annual-level sources (Cash/Credit/Liability ledger CSVs, TPST, Portal
  Tax-Liability-Comparison, BO Profile)** were **never guarded at all** --
  any ONE of them being absent crashed `master_build.py`'s `annual_data =
  dict(...)` call immediately (unguarded `open(None)` / `load_workbook(None)`
  inside `annual_sources.py`/`bo_profile_parser.py`). Now wrapped by
  `_safe_parse_ledger()` / `_safe_parse_credit()` / `_safe_parse_tpst()` /
  `_safe_parse_portal()` / `_safe_parse_bo()` in `master_build.py`, each
  returning the same empty shape the real parser would return for a
  present-but-empty file, so every downstream consumer keeps working.
- **One month's unexpected parsing error no longer crashes the whole
  multi-year run** -- `master_build.py`'s per-month loop now wraps
  `run_month()` in try/except; a failing month is logged, skipped, and
  listed explicitly on the Dashboard, while every other month still
  processes and gets included in the output workbook.

---

## 2. 5-year / multi-taxpayer architecture (built; needs a real multi-FY
   dataset to fully exercise end-to-end)

- `folder_classifier.py`: previously kept only the LAST matching file per
  document type (`gstr1_merged = f` inside a loop, silently overwriting
  every earlier match) -- feeding 5 years of separate merged GSTR-1
  workbooks would have silently discarded 4 of them with zero warning.
  Now collects **every** matching file per type into a list, and
  `_build_month_file_map()` resolves, per month, which specific file
  actually covers it (any number of FYs). Ambiguous overlaps (two files
  both claiming the same month) are reported as an explicit warning, not
  silently resolved.
- `master_build.py`: the hardcoded 12-month `MONTH_ORDER` (`Apr-22`...
  `Mar-23`) is no longer used to decide what "covered" means --
  `months_covered` is now the actual chronologically-sorted intersection of
  whatever GSTR-1 and GSTR-3B months were found, spanning any number of
  years. `_month_sort_key()`/`_sort_months_chronologically()`/
  `_fy_label_for_month()` replace every hardcoded-list usage (rectification
  pairing, dashboard sort, dashboard header). Output filename and dashboard
  header now show the REAL FY-range covered (e.g. `FY2022-23_to_2026-27`)
  instead of a hardcoded `FY2022-23` string.
- **Not yet exercised against real data**: I do not have a second (or
  fifth) year of this taxpayer's merged files to actually run multi-year
  end-to-end. The month-map resolution logic is syntax-correct and
  unit-testable (`_build_month_file_map` was written defensively with an
  explicit overlap-warning path) but a real 5-year folder is the next
  thing to test this against.
- **Not yet generalized**: annual-level sources (ledgers/TPST/portal-
  comparison/BO-Profile) still resolve to a SINGLE file each
  (`res["cash_ledger"]` = first match). These are inherently FY-scoped
  documents (a cash ledger export covers one FY), so true 5-year support
  needs `build_annual_workbook.py`'s reconciliation run once per FY and
  the results shown side-by-side/trended -- `cash_ledgers`/`credit_ledgers`/
  `liab_ledgers`/`tpst_files`/`portal_comparison_files`/`bo_profile_files`
  (plural, list) are already returned by the updated `classify_folder()`
  so this is a contained next step, not a re-architecture.

---

## 3. Cancelled E-Invoices (new)

- `gst_scrutiny_tool.parse_einv()`: now tries several real-world header-name
  variants (`IRN Status`, `Status`, `Cancel Date`, etc. -- content-based,
  never a fixed column index) to find a cancellation-status column. If
  found, every cancelled row is collected (invoice no., rate, taxable,
  IGST, IRN, cancel date). If NOT found, `cancel_col_found=False` is
  surfaced explicitly so "zero cancelled invoices" is never confused with
  "this export doesn't expose that field at all."
- `forensic_checks.build_cancelled_einvoice_findings()`: aggregates across
  every month, then runs the two defensive cross-checks already on the
  project's own backlog (item B9 / D2, previously identified but never
  built): (a) a cancelled e-invoice's number still appearing as a live
  GSTR-1 B2B outward supply; (b) a cancelled e-invoice with a still-live
  outward EWB against the same invoice number.
- New "Cancelled E-Invoices" sheet in the master workbook.
- **Not yet tested against real data**: no real E-Invoice export with an
  actual cancelled row was available this round -- the header-detection
  logic is generic/defensive but the exact real-world column name should
  be confirmed against your actual E-Invoice file next round.

---

## 4. ARN Date / Filing-Gap / Late Fee -- fixed (new module: `filing_compliance.py`)

- **Root cause found and fixed**: the old ARN-date extraction
  (`gst_unified_scrutiny._extract_arn_dates()`) lived inside `gather()`, a
  function explicitly marked "LEGACY / UNSUPPORTED for the merged-file
  model" in its own docstring -- `master_build.py` (the real pipeline)
  never called it. So `GSTR1_FILING_DATE`/`GSTR3B_FILING_DATE` were always
  `None` in every real run, and Analysis checks #8/#10 always fell through
  to INFO.
- New `filing_compliance.py`:
  - `gstr1_arn_dates_by_month()` -- reads EVERY period-marker row across
    every GSTR-1 sub-sheet (a merged whole-FY file has one marker per month
    per sub-sheet), extending the marker text's own `ARN: ...` field with
    an optional date capture. Falls back to the single `Read me` sheet
    value ONLY with an explicit warning (never silently stamps one date
    onto all 12 months).
  - `gstr3b_arn_dates_by_month()` -- GSTR-3B is one sheet per month, so
    this is genuinely per-month already; reused the existing `Date of ARN`
    field-read logic.
  - `due_date_gstr1()` / `due_date_gstr3b()` -- statutory due dates (11th/
    20th for monthly filers; 13th/22nd-or-24th for QRMP, by CBIC state
    category), auto-selectable by filer type.
  - `compute_late_fee()` -- Section 47, Rs 50/day (Rs 20/day nil), capped
    per Notification 07/2023-CT by turnover slab (cap NOT applied, and
    says so explicitly, if turnover isn't supplied -- never silently picks
    a slab).
  - `compute_interest()` -- Section 50(1), 18% p.a. on cash-paid tax.
  - `month_filing_compliance()` -- ties it all together per month.
- Wired into `run_monthly_pipeline.run_month()`: computes the record per
  month and **actually sets** `raw.GSTR1_FILING_DATE`/`GSTR3B_FILING_DATE`
  before calling the existing engines -- so `gst_analysis_checks.py`'s
  checks #8 and #10 (which already had the CONFIG-reading logic, just never
  populated) now work with ZERO changes needed to that file.
- New "Filing Compliance & Late Fee" sheet in the master workbook shows
  ARN, filed date, due date, and computed Rs late fee + interest per month
  for both GSTR-1 and GSTR-3B.
- **Tested**: `filing_compliance.py`'s date math (due-date calculation,
  late-fee formula, interest formula) self-tested with worked examples in
  its own `__main__` block, and `gstr1_arn_dates_by_month()` tested against
  a synthetic marker-based workbook (real marker-with-ARN-and-date text,
  since I don't have your real merged GSTR-1 file). **Needs confirming**
  against your actual merged GSTR-1 file next round: whether its real
  period-marker text includes a date after the ARN, or only the ARN number
  (in which case the Read-me fallback path -- also built, with an explicit
  warning -- is what will fire; flagged clearly in the output either way,
  never silently wrong).

---

## 5. GST_Forensic_Comparison_Framework_v1.md -- implemented

### Part 2 (fully built AND verified against your real files)

- **New `annual_return_parser.py`**: parses GSTR-9 PDF, GSTR-9C PDF, and
  Table 8A (government-standard xlsx). All three tested against your real
  uploaded files -- every figure extracted matches the framework document's
  own manually-verified numbers EXACTLY (turnover Rs 46,56,39,087.14, Table
  8A CGST/SGST/IGST/total, GSTR-9 Table 6A total Rs 4,78,88,427.54, GSTR-9C
  ARN/ARN-date, etc. -- see this session's test output). Handles the real
  "SYSTEM COMPUTED" watermark noise found in the government PDF export
  (confirmed: individual watermark letters get interleaved into the text
  stream both as standalone lines and mid-word; a real alignment bug was
  found and fixed during testing -- see the module's own docstring).
- **R14 (four-way ITC reconciliation)**: `forensic_checks.check_four_way_itc()`.
  Tested with all 4 sources (3B/2B/8A/Books) -- reproduces the framework
  document's exact gap figures: 3B-Books = Rs 10,49,570.28, 2B-Books =
  Rs 9,39,250.40, 8A-Books = Rs 7,81,296.77. Runs with as few as 2 of the 4
  sources available, explicitly naming which were used.
- **R13 (turnover-gap rule)**: `forensic_checks.check_turnover_gap()`.
  Tested -- correctly FLAGs the exact scenario in the framework document
  (Rs 16,70,77,200.00 GSTR-9C Table 7B adjustment with zero supporting rows
  in GSTR-1 Table 8 across all 12 months).
- **R0-R12 (generic BS/PL rule engine)**: `forensic_checks.check_bs_pl_rules()`.
  Structured-input based (see `bs_pl_input.py` + `OCR_LIMITATION.md` for
  why, not OCR). Tested with your real transcribed figures -- R0 (BS
  self-balances) PASS, R1 (revenue vs 9C) PASS exact match, R12 (reserves
  roll-forward) PASS exact match, all matching the framework document's
  own §2.1 "what ties out cleanly" findings. R3-R10 correctly degrade to
  INFO with the specific missing-Notes explanation the framework document
  itself specifies, since Notes to Accounts weren't supplied.

### Part 1 (built; NOT yet tested -- no ledger/TPST/portal-comparison/
  BO-Profile files were supplied this round)

The following are written into `forensic_checks.py`'s docstring and
architecture but **not yet coded as separate functions** this round, since
building and testing them blind (without your real ledger CSVs, TPST
export, portal-comparison Excel, or BO-Profile PDF to verify against) risks
exactly the kind of unverified-guess problem this whole tool exists to
avoid. Each is a contained addition once real files are available:
  - A1 ledger line-item tie-out (reference-ID 3-way join)
  - A2 ITC-utilization sequencing (Rule 88A/Sec 49)
  - A3 ledger balance-continuity gate
  - A4 independent interest/late-fee recompute vs ledger-actual (the
    FORMULA is built and tested -- `filing_compliance.compute_interest()`
    -- only the ledger-actual cross-check side is pending real data)
  - B1 portal-comparison cumulative-% trend detector
  - B3 "Payment made through DRC-03" column read
  - C1 B2B-CDNR line-item tie-in
  - C2 2B "ITC Availability=No" reason-code bucketing (the Table-8A
    equivalent of this IS built and tested -- see `parse_table_8a()`'s
    `totals['no_reason_breakdown']`, confirmed on your real file: 124
    "Reverse charge document" + 2 "POS lies in the State of supplier" --
    the GSTR-2B version is the same pattern, applied to a different sheet)
  - D1 EWB vehicle cross-match (inward vs outward)
  - E1 TPST/BO-Profile/2B counterparty-count triangulation
  - F1 DRC-03 payment aggregator + FLAG-matching
  - F2 departmental-proceedings FY-attribution

---

## 6. Files in this package

15 original + 5 new/changed-shape:
`annual_return_parser.py` (new), `filing_compliance.py` (new),
`forensic_checks.py` (new), `bs_pl_input.py` (new),
`OCR_LIMITATION.md` (new), plus patched: `gst_scrutiny_tool.py`,
`gst_eway_recon.py`, `gstr2b_parser.py`, `folder_classifier.py`,
`master_build.py`, `run_monthly_pipeline.py`. Unchanged from your original
upload: `merged_period_utils.py`, `amendments.py`, `ewb_annual_parser.py`,
`annual_sources.py`, `bo_profile_parser.py`, `build_annual_workbook.py`,
`gst_analysis_checks.py`, `gst_unified_scrutiny.py`, `hsn_fraud_checks.py`.

Every file syntax-checked and import-checked together (see session
transcript) -- `master_build.py` is the entry point, unchanged usage:
```
python3 master_build.py <folder>
```
