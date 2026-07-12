# GST Scrutiny Tool — Complete Reference (v3)
**This is the single, authoritative document for this tool. Paste it whole into any new chat and Claude will have full working context to continue development, fix bugs, or explain any part of the system to a validation team.**

**Taxpayer used to build and verify every check in this tool:** M R HEALTHCARE PRIVATE LIMITED, GSTIN `05AAECM6380J1ZA`, a pharmaceutical manufacturer in Ramnagar, Uttarakhand. **The tool itself is generic** (built to run for any GSTIN) — §9 lists exactly what is/isn't taxpayer-specific and what to edit for a new taxpayer.

**This document supersedes** every earlier README/changelog produced during this project (`GST_Tool_Full_Project_Context.md`, `CHANGELOG_v2.md`, `GST_Tool_v2_Session_Context.md`, `BUGFIX_ROUND_1.md`, `OCR_LIMITATION.md`). Those are kept in the package for historical detail, but this document is current and self-contained — you do not need to read them first.

---

## 1. What this tool is, in one paragraph

A GST reconciliation and forensic-scrutiny tool for one taxpayer's one (or more) financial year. It takes the taxpayer's own filed returns (GSTR-1, GSTR-3B, GSTR-2B, E-Invoice, E-Way-Bills), department-side data (ledgers, TPST, portal comparison, BO Profile, annual return GSTR-9/9C, Table 8A), and optionally a Balance Sheet/P&L, and cross-checks every one of them against every other one — surfacing every mismatch, gap, or anomaly a GST officer or consultant would raise in scrutiny, always with the exact underlying arithmetic shown, never a bare assertion. It is built on one non-negotiable rule, repeated everywhere in the code as comments: **never fabricate or silently guess a number; when something is missing or uncertain, say so explicitly (SKIP/INFO/REVIEW), never present an assumption as a verified fact.**

---

## 2. How to run it

```bash
pip install openpyxl pypdf pdfplumber pytesseract --break-system-packages
cd <folder containing all data files + all 21 tool files below>
python3 master_build.py .
```
Output: one workbook, `GST_MASTER_<GSTIN>_FY<range>.xlsx`, written to the same folder.

Data filenames never matter — every file is identified by its own content (sheet names, header text, marker text, PDF first-page signature), never by filename. **Only the tool's own 21 filenames matter, exactly as listed in §5.** All files (data + tool) sit flat in one folder; no subfolders needed.

**For a new taxpayer:** just replace the data files. Optionally fill in `bs_pl_input.py`'s `BS_PL_DATA` dict with the new taxpayer's real Balance Sheet/P&L figures (see §7.6) — everything else needs zero editing except the one exception in §9 (`HSN_RATE_HISTORY`'s curated codes, which only cover this taxpayer's own real HSN codes so far).

---

## 3. Architecture — six layers

1. **Per-month engines** — for every month with both GSTR-1 and GSTR-3B data: a raw side-by-side comparison (`gst_scrutiny_tool.py`) → a 14-check interpretive layer (`gst_analysis_checks.py`) → a 27-check E-Way-Bill matrix (`gst_eway_recon.py`) → (new) a per-HSN rate-review table appended into the same Comparison sheet.
2. **Cross-month layer** — links an error reported in one month to its correction in a later month (Rectification Pairs); checks invoice-series integrity against Table 13's declared document counts, cross-referenced against declared cancellations AND cancelled e-invoices (Doc-Series Integrity).
3. **Annual/FY-wide layer** — reconciles ledgers, TPST, the portal's own comparison report, and the BO/360° Profile against the monthly totals (`build_annual_workbook.py`).
4. **HSN-code-wise + fraud-pattern layer** — 36+ checks (HSN classification, POS/state-code logic, named fraud patterns), run once across the whole FY (`hsn_fraud_checks.py`).
5. **Forensic layer** — GSTR-9/9C/Table-8A four-way ITC reconciliation, turnover-gap rule, a generic Balance-Sheet/P&L rule engine, per-month filing-compliance/late-fee, Cancelled E-Invoices (`forensic_checks.py`, `filing_compliance.py`, `annual_return_parser.py`).
6. **HSN rate-reference layer** — three independent, differently-trusted sources feeding the checks above and a dedicated per-month review table; see §7 for the full deep-dive, since this is the most-discussed and most-revised part of the tool this session.

**Input files are MERGED, whole-FY workbooks** — one Excel file per document type, containing every month's data stacked in the same sheets, separated by a period-marker row (`Financial Year: 2022-23 | Tax Period: January | ARN: ... | ARN Date: ...`). `merged_period_utils.py` finds these markers by content and slices each sheet into `{month: [rows]}` blocks. GSTR-3B is the exception — it merges as one *sheet* per month, and even there the sheet's *name* is never trusted; the month comes from an in-sheet `Year`/`Tax Period` field. **Multiple merged files of the same document type across different FYs are supported** — `folder_classifier.py` resolves, per month, which specific file covers it (architecturally built; not yet run end-to-end against a real 2+-FY folder — see §10).

Every parser is content-based, never filename- or sheet-name-based.

---

## 4. Complete input file catalog

### 4.1 Core filing data (2 mandatory, rest optional)

| # | File | Format | Mandatory? | Contents |
|---|---|---|---|---|
| 1 | GSTR-1 (merged) | Excel, whole FY | **Yes** | B2B/SEZ/DE invoices, B2C Large/Small, exports, CN/DN, HSN summary, amendments (b2ba/cdnra), Table 13 |
| 2 | GSTR-3B (merged) | Excel, whole FY (1 sheet/month) | **Yes** | Outward liability, RCM, ITC availed/reversed/ineligible, ARN + filing date |
| 3 | E-Invoice (merged) | Excel, whole FY | No | IRN, IRN date, invoice date/value/rate/tax, IRN/e-invoice status (Valid/Cancelled) if the export carries that column |
| 4 | GSTR-2B (merged) | Excel, whole FY (quarterly blocks) | No | Table 3 summary, B2B purchase invoices + eligibility flag, B2B-CDNR |
| 5 | Electronic Cash Ledger | CSV, whole FY | No | Every cash transaction by tax head |
| 6 | Electronic Credit Ledger | CSV, whole FY | No | Every credit transaction by tax head |
| 7 | Electronic Liability Register | CSV, whole FY | No | Every liability transaction |
| 8 | Outward E-Way-Bill | Excel, whole FY | No | EWB no./date/time, invoice no./date, assessable & tax value, HSN, vehicle |
| 9 | Inward E-Way-Bill | Excel, whole FY | No | Same fields, inward |
| 10 | GST-Prime TPST | Excel, whole FY | No | 12-month self-filing summary |
| 11 | Portal "Tax liability & ITC comparison" | Excel, whole FY | No | Monthly GSTR-1-vs-3B and 3B-vs-2B comparison |
| 12 | BO / 360° Profile | PDF, whole FY (can hold multi-year data) | No | Financial summary, BIFA figures, ITC passed/received, Top-10 lists, related/cancelled-party ITC, DRC payments, Appeal/Case/Transfer sections |
| 13 | GSTR-9 (Annual Return) | PDF, text-based (not scanned) | No | Table 4 (outward liability), Table 5 (nil/exempt), Table 6A (ITC via 3B), Table 9 (tax paid/late fee/interest) |
| 14 | GSTR-9C (Reconciliation Statement) | PDF, text-based | No | Turnover reconciliation, ITC reconciliation, ARN + ARN Date |
| 15 | Table 8A | Excel, government-standard export | No | Invoice-level inward supplies with "ITC available Yes/No" flag + reason |
| 16 | Balance Sheet + P&L | PDF **or** `bs_pl_input.py` hand-typed dict | No | Total Assets/Liabilities, Revenue, Payables/Receivables, Inventories, Fixed Assets, Investments, Provisions, Expenses, Finance Costs, Reserves & Surplus — see §7.6 for why a scanned PDF is never auto-parsed |

### 4.2 HSN rate-reference sources (new this session — see §7 for the full picture)

| # | File | Format | Mandatory? | What it actually provides |
|---|---|---|---|---|
| 17 | HSN/SAC code-and-description master | Excel (sheets `HSN_MSTR`+`SAC_MSTR`) | No — bundled default ships with the tool | Code EXISTENCE + official description only. No rate column at all. |

**Only #1 and #2 are mandatory.** Everything else degrades gracefully and independently if absent — see §8's matrix.

---

## 5. Complete file catalog — 20 `.py` files + 1 bundled data file

| File | Role |
|---|---|
| `merged_period_utils.py` | Marker-row detection — foundation for every other parser |
| `folder_classifier.py` | Identifies every file by content signature; discovers month coverage; supports multiple files per document type (multi-year) |
| `gst_scrutiny_tool.py` | Raw comparison engine (one month at a time); E-Invoice parser incl. cancellation detection |
| `gst_analysis_checks.py` | The "14 checks" interpretive layer |
| `gst_eway_recon.py` | The 27-check EWB matrix; degrades honestly when a whole EWB direction is absent |
| `gstr2b_parser.py` | Shared GSTR-2B reader; degrades gracefully instead of crashing when 2B is absent for a month |
| `amendments.py` | Amendment sheets (b2ba/cdnra) + Table-13 doc-series integrity |
| `run_monthly_pipeline.py` | Runs the three per-month engines together; wires filing-compliance and cancelled-e-invoice pass-through |
| `annual_sources.py` | Ledgers, TPST, portal-comparison parsers |
| `bo_profile_parser.py` | BO/360° Profile PDF parser; hardened section-boundary detection; `--diagnose` mode |
| `build_annual_workbook.py` | FY-wide annual-reconciliation sheets; BIFA lookup is FY-dynamic (not hardcoded); false-REVIEW-on-missing-data fixed |
| `ewb_annual_parser.py` | Whole-FY EWB parser |
| `gst_unified_scrutiny.py` | Shared Excel sheet-writers (Comparison, Analysis14, EWB), reused by `master_build.py` |
| `hsn_fraud_checks.py` | HSN-code-wise + fraud-pattern engine (36+ checks); date-versioned HSN rate reference; mcp-india-stack integration; HSN/SAC master-validity check |
| `master_build.py` | Single entry point. Orchestrates every layer, writes the output workbook, including the new per-month HSN Rate Review table |
| `annual_return_parser.py` | Parses GSTR-9, GSTR-9C, Table 8A |
| `filing_compliance.py` | ARN-date extraction, statutory due dates, Sec 47 late fee, Sec 50 interest |
| `forensic_checks.py` | R13 (turnover-gap), R14 (four-way ITC), R0-R12 (BS/PL rule engine), Cancelled-E-Invoice cross-checks, Doc-Series cancelled-e-invoice enrichment |
| `bs_pl_input.py` | Hand-typed, GSTIN-tagged structured input for the BS/PL rule engine (never OCR — see §7.6) |
| `HSN_SAC_default.xlsx` | Bundled data file (not `.py`) — the default HSN/SAC code-and-description master, used when no run-specific override is supplied (see §7.3) |

Total: approximately 9,700 lines of Python across 20 files.

---

## 6. Complete check catalog

### 6.1 Raw Comparison engine (`gst_scrutiny_tool.py`) — per month
A. Outward Liability (GSTR-1 net-of-CN vs 3.1(a)) · A2 GSTR-1-internal cross-check · B. E-Invoice vs GSTR-1 (taxable/IGST/CGST/SGST, unique-invoice count, missing-IRN — cancelled e-invoices excluded from both sides) · B2. Line-level invoice+rate match · C. RCM (3.1(d) & 4A3 vs 2B — SKIPPED if 2B absent) · D. ITC gross+net-of-CN (SKIPPED if 2B absent) · D2. ITC Reversal (SKIPPED if 2B absent).

### 6.2 Analysis — 14 checks (`gst_analysis_checks.py`) — per month
0 B2B totals · 1 Nil/exempt/non-GST · 2 Credit-note effect · 3 ITC arithmetic 4C formula · 4 Effective tax-rate · 5 Dropped invoice numbers · 6 Duplicate invoice numbers · 7 E-invoice error flags · 8 IRN-vs-filing lag · 9 Rate-wise e-invoice vs HSN · 10 Filing-gap/late fee · 11 POS vs GSTIN tax-head · 12 RCM routing · 13 HSN-summary vs named-invoice IGST · 14 ITC/Liability ratio. (Checks #8 and #10 now actually populate — see §8's bug-fix list, item 2.)

### 6.3 E-Way-Bill — 27 checks (`gst_eway_recon.py`) — per month
1 EWB-Out in GSTR-1 · 2 Value vs GSTR-1 · 3 Unmatched EWB-Out · 4 Inter-state >Rs 50K no EWB · 5-6 vs E-Invoice · 7 Date gap · 8 Aggregate vs 3.1(a) · 9 Tax-type vs GSTR-1 · 10-14 EWB-In vs GSTR-2B · 15 Out vs In · 16 Gen-vs-doc-date gap · 17 Triangulation · 18 HSN rate-wise · 19 3B ITC vs EWB-In · 20 purchase E-Invoice vs EWB-In · 21 vs filing date · 22 validity vs supply date · 23 multiple EWBs/invoice · 24 EWB cancelled after filing · 25-26 assessable-value ratios · 27 same-vehicle repeated trips. (Checks #1,3,4,8,9,10,12,13,25,26 now show honest SKIPPED, not misleading PASS/REVIEW/FLAG, when a whole EWB direction was never supplied — bug-fix list item 1.)

### 6.4 Cross-month layer
- **Doc-Series Integrity** — Table-13 declared ranges vs actual invoice numbers, 3-tier matching, cross-referenced against declared cancellations AND (new) the Cancelled E-Invoices list.
- **Rectification Pairs** — amendment rows traced to their original month, cross-referenced against DRC payments.

### 6.5 Annual reconciliation (`build_annual_workbook.py`)
Cash/Credit/Liability ledgers vs TPST vs portal comparison, monthly. FY Total vs BIFA (FY-dynamic lookup; honest blank-vs-mismatch distinction). Related-Party Alerts / Top Counterparties (honest zero-rows note distinguishing "genuinely none" from "parser found nothing"). Departmental proceedings (Appeals/Cases/Transfers) on the Annual Cover sheet.

### 6.6 HSN-code-wise + Fraud-Pattern checks (`hsn_fraud_checks.py`) — FY-wide, 36+ checks
- **A1** wrong rate vs curated `HSN_RATE_HISTORY` (date-aware — see §7.1)
- **A1-EXT** wrong rate vs mcp-india-stack reference (REVIEW only, date-gated to post-22-Sep-2025 months only — see §7.2)
- **A7** (new) HSN/SAC code-existence check against the official code master (see §7.3)
- A2 exempt-charged · A3 cess-missing · A4 multi-rate same code · A5 blocked-ITC-by-HSN · A6 invalid/short HSN
- B1/B3 wrong tax head vs POS · B2 B2C-Large no-EWB · B3-2B GSTR-2B "ITC Availability=No" reason codes · B4 SEZ misclassification
- C1 RCM-HSN-not-declared · C2 branch-transfer · C3-C5 various combined checks
- Approximately 25 named fraud-pattern checks: round-number invoices, reciprocal trading, CN timing, HSN drift, year-end dumping, ghost-supplier PAN clusters, IRN delay, midnight EWB, EWB bursts, credit hoarding, cash-timing patterns, rate outliers, etc. (checks #1 through #57 per `run_all()`'s own registration — unchanged this session)

### 6.7 Forensic layer (`forensic_checks.py`, `filing_compliance.py`, `annual_return_parser.py`)
- **R13** turnover-gap rule (GSTR-9C exempt/nil/non-GST adjustment vs GSTR-1 Table 8's actual rows)
- **R14** four-way ITC reconciliation (3B via GSTR-9 / 2B FY-total / Table 8A / Books via GSTR-9C) — runs with any 2+ of the 4 sources
- **R0-R12** generic Balance-Sheet/P&L rule engine (self-balance gate, revenue-vs-9C, payables/receivables/inventory/fixed-assets/investments/provisions/expenses/finance-costs screens, reserves roll-forward) — structured-input based, wired into the main pipeline (was built standalone earlier and only actually connected to `master_build.py`'s output in this final round)
- **Filing Compliance** — per-month ARN date, statutory due date, Sec 47 late fee (with Notification 07/2023-CT turnover caps), Sec 50 interest
- **Cancelled E-Invoices** — aggregated list + cross-checks: (a) cancelled invoice still in GSTR-1 B2B, (b) cancelled invoice with a still-live outward EWB

### 6.8 NEW — Per-month HSN Rate Review table (inside each month's Comparison sheet)
Appended below the existing raw-comparison table in every `Comparison <Month>` sheet: every HSN code used that month, its taxable value / rate charged / tax amount (straight from GSTR-1's own HSN summary), plus two reference-rate columns (curated `HSN_RATE_HISTORY` and mcp-india-stack), and a Status column that ALWAYS reads "VERIFY" (elevated to "VERIFY -- reference rate differs", red-highlighted, when either reference disagrees with the charged rate). This is intentionally a raw side-by-side worksheet for human review, not an automated verdict — the automated verdict is checks A1/A1-EXT/A7 above.

---

## 7. The HSN rate-reference system — full deep-dive

This was the single most-revised part of the tool this session, across several rounds of real-data discovery. Three genuinely different sources exist, each with a different trust level, and none of them alone is sufficient:

### 7.1 `HSN_RATE_HISTORY` (in `hsn_fraud_checks.py`) — the primary, curated source
A hand-built, date-versioned Python dict: `{HSN_code: [(from_date, to_date_or_None, rate_or_None, description, confidence, source_note), ...]}`. Longest-prefix code matching (`_hsn_prefix_lookup`), then period matching by the invoice's own month (`_hsn_rate_for_date`).

**Why date-versioned at all:** confirmed via research this session that GST rates were massively overhauled on 22-Sep-2025 (Notification 9/2025-CT(Rate), superseding Notification 01/2017-CT(Rate) — the 12% slab merged into 18%, a new 40% peak rate was added, most medicaments moved to 5%/Nil). A single "current" rate snapshot would be actively wrong for scrutinizing a past FY like this taxpayer's FY22-23. Every code's pre-22-Sep-2025 rate is "high confidence" (taken from the taxpayer's own real, filed FY22-23 GSTR-1). Post-22-Sep-2025 rates are researched per-code; where research was inconclusive, the rate is explicitly `None` with confidence `"unconfirmed"` — the check then shows INFO ("rate reference unconfirmed for this period"), never a guessed FLAG/REVIEW.

Current coverage (9 codes, all from this taxpayer's real data): `3003, 3004, 3808, 3915, 4707, 7204, 7606, 8402, 998843`. Post-2.0 confidence: `3808` and `7204` confirmed unchanged (high), `4707` confirmed unchanged (medium), `3003`/`3004` genuinely product-dependent (unconfirmed), `3915`/`7606`/`8402`/`998843` not researched (unconfirmed). One real discrepancy flagged, not resolved: `3915` sits at 5% in this taxpayer's own filed data, but several generic web sources claim plastic scrap is 18% — kept at the taxpayer-verified figure with the conflict noted in the source comment, not silently overwritten either way.

### 7.2 `mcp-india-stack` (PyPI package) — supplementary, REVIEW-only, now known to be nearly empty
Real, non-hallucinated PyPI package (`rehan1020/MCP-India-Stack` on GitHub) bundling a `data/hsn/hsn_master.csv`. Read directly via the CSV file (`_load_mcp_india_stack_hsn_table()`), bypassing the package's own heavy MCP-server/fastmcp/starlette dependency chain entirely (confirmed via real package inspection this session: `lookup_hsn_code()` is not a plain importable function the way its own marketing example showed — it's wrapped for MCP-protocol access only).

**Critical finding, confirmed on a real freshly-refreshed export this session:** 22,471 of 22,500 rows (99.9%) carry an all-zero rate (CGST=SGST=IGST=CESS=0.0) — this is not "nil-rated," it's "no rate was ever populated for this entry." Only 29 rows (0.1%) carry a real rate, and those correspond exactly to a handful of codes the maintainer patched by hand per the package's own changelog (`8517, 9401, 2523, 3004, 8708`). Of this taxpayer's own 9 curated codes, 8 have ZERO usable rate data in this source, and the 9th (`3004`) has a disagreeing duplicate row that gets dropped by the loader's own ambiguity guard. After all defensive filtering, this source yields 9 usable codes total, none of them this taxpayer's own real codes.

Given this, the loader (`_load_mcp_india_stack_hsn_table()`) applies three independent defensive filters before trusting any row:
1. Drop rows where CGST+SGST does not equal IGST (internal inconsistency)
2. Drop any code with multiple rows that DISAGREE on rate (never guess which applies — stricter than the package's own `lookup_hsn_code()`, which silently takes `rows[0]`)
3. Drop any lone row whose rate is exactly 0.0 (added this session after the finding above — a single zero-rate row is now known to overwhelmingly mean "no data," not "genuinely nil")

The resulting check (A1-EXT) is always REVIEW severity, never FLAG, carries the package's own disclaimer text verbatim, and is date-gated to only apply for months on/after 22-Sep-2025 (the package's undated "latest" data most likely reflects post-GST-2.0 rates; applying it against a pre-2.0 invoice would recreate the exact systematic-mismatch risk this whole design avoids).

**Bottom line on this source:** it is now understood to add close to zero real coverage for this taxpayer. It remains wired in (harmlessly — it will essentially never fire) in case a future taxpayer's HSN codes happen to fall in the ~30 codes the package maintainer has manually patched, or in case the upstream package substantially improves its data quality later.

### 7.3 HSN/SAC code-and-description master (new — `HSN_SAC_default.xlsx` / user-supplied override) — code validity, NOT rates
A genuinely official-looking master (very likely sourced from the NIC e-Invoice system's own downloadable HSN/SAC list, per its exact `HSN_MSTR`/`SAC_MSTR` sheet structure) — 21,935 HSN codes + 681 SAC codes, zero duplicates, full 2/4/6/8-digit hierarchy, verified this session to contain all 9 of this taxpayer's curated codes. Critical limitation: this file has CODE + DESCRIPTION columns only — no rate column exists at all. It therefore cannot answer "what's the rate," only "does this code exist" and "what's its official description."

Loader: `_load_hsn_sac_master(override_path=None)` — tries a run-supplied file first (any file in the run folder with sheets `HSN_MSTR`+`SAC_MSTR`, auto-detected by `folder_classifier.py`), falls back to the bundled `HSN_SAC_default.xlsx` shipped alongside the `.py` files. "Hardcoding" a new default: replace `HSN_SAC_default.xlsx` with a new upload in a future session — this is a manual, session-triggered action (no background auto-update exists — see §7.4), consistent with how `HSN_RATE_HISTORY` itself is maintained.

New check A7: flags (REVIEW, not FLAG — the master is a point-in-time snapshot) any reported HSN/SAC code that does NOT appear in this master at all, and shows the official description alongside the taxpayer's own for a quick human eyeball-check (no automated fuzzy-text matching is attempted — too noisy/fragile to be worth it).

### 7.4 Why none of this is "automatic," and what internet access does/doesn't change
Extensively discussed and tested this session. Summary:
- Claude (this tool's author) only runs when invoked in a chat session — there is no background process that watches CBIC/GST-Council announcements and updates the code between sessions.
- `mcp-india-stack`'s own "auto-update" (`--refresh-all`, pulling from a jsDelivr CDN mirror of its GitHub repo) is real and was tested this session — but even a fresh pull confirmed the exact same 99.9%-empty problem, since the underlying GitHub data itself is what's incomplete, not a stale local cache.
- The official government HSN/SAC search (`services.gst.gov.in/services/searchhsnsac`) is a JavaScript-rendered single-page app with no documented public bulk API — not realistically automatable from this tool, and even if it were, it only exposes the CURRENT rate, which is exactly the wrong thing to check a PAST FY against (see §7.1's whole reason for existing).
- Recommended practical workflow, agreed with the person building this tool: run with the internet OFF for actual scrutiny runs (reproducibility — a forensic finding should not change just because a third-party reference silently updated underneath it). Whenever genuinely new rate information is needed, explicitly ask Claude in a session to research and add a new dated entry to `HSN_RATE_HISTORY` — a manual, deliberate, verifiable action, not a silent background one.

### 7.5 What "efficient" means for this layer, honestly
For a taxpayer whose HSN codes are already in `HSN_RATE_HISTORY` (i.e., this taxpayer, today), the check is fast, precise, and high-confidence. For a new taxpayer/industry, expect: A1 shows nothing (no curated codes), A1-EXT shows almost nothing (mcp-india-stack's near-total coverage gap), A7 will at least confirm code validity, and the new per-month "HSN Rate Review" table (§6.8) will show every code with a "VERIFY" flag and mostly "not found"/"not in curated list" reference columns — i.e., the tool correctly tells you it doesn't know, rather than guessing. The single highest-leverage improvement for a new taxpayer is spending 10-15 minutes extending `HSN_RATE_HISTORY` with that taxpayer's own real codes, the same way this taxpayer's 9 codes were curated.

### 7.6 Balance Sheet / P&L — same "don't guess" discipline, different reason
`bs_pl_input.py`'s `BS_PL_DATA` dict is hand-typed, never OCR'd. Tested: OCR (`pytesseract`) on this taxpayer's real scanned BS/PL PDF silently misread a real digit (Rs 49,73,007.06 read as Rs 43,73,007.05, a Rs 6 lakh error) — exactly the kind of error this tool exists to prevent, not commit. `BS_PL_DATA` carries a `_gstin` tag; `master_build.py` refuses to use it if the tag doesn't match the taxpayer actually being processed (prevents a stale/wrong-taxpayer's figures being silently reused).

---

## 8. Bug-fix history this session (chronological, for the validation-team record)

1. GSTR-2B entirely absent for a month used to crash the whole run (nothing caught the raise). Fixed: `gstr2b_parser.summary_for_month()` now returns `available=False` explicitly; consumers show one clean SKIPPED row.
2. An entire EWB direction never supplied used to show misleading PASS/REVIEW/FLAG (false "0/0 matched" clean results, or floods of false "missing EWB" flags). Fixed: `gst_eway_recon.run()` takes explicit `ewb_out_file_supplied`/`ewb_in_file_supplied` flags; affected checks now show honest SKIPPED.
3. Annual-level sources individually absent (ledgers/TPST/portal/BO-Profile) used to crash the whole run via an unguarded `open(None)`. Fixed: five `_safe_parse_*()` wrappers in `master_build.py`.
4. `folder_classifier.py` kept only the LAST matching file per document type, silently discarding earlier years' files for multi-year runs. Fixed: collects every match into a list; `_build_month_file_map()` resolves per-month file ownership.
5. ARN-date extraction was completely unwired — the old function lived in a legacy code path `master_build.py` never called, so filing-gap/late-fee checks always showed INFO. Rebuilt as `filing_compliance.py`, properly wired.
6. Cancelled-E-Invoice header detection was case-sensitive — the real export's exact header (`E-invoice status`, lowercase) never matched the candidate list (`E-Invoice Status`, different case) even though the right name was "in the list." Fixed: case-insensitive matching.
7. Cancelled invoices were still polluting the main E-Invoice totals even once detected — a cancelled invoice (correctly absent from GSTR-1) was still counted in the E-Invoice side of every comparison, manufacturing false gaps (confirmed: 6 real cancelled invoices, Rs 44,12,291 taxable, were producing 6 false "LINE-LEVEL GAP" mismatches). Fixed: cancelled rows are now excluded from every aggregate/comparison, recorded separately.
8. ARN-date regex only recognised "Date [of] Filing/ARN" word order, but the real marker text uses "ARN Date" (ARN comes first) — `ARN: AA050422057237G | ARN Date: 10-05-2022` never matched. Fixed and re-verified against the exact real marker text.
9. BO Profile section-boundary detection was exact-substring, single-line only — silently returns nothing if a real PDF's text extraction wraps a header across a line break, or differs slightly in case/spacing. Hardened with a case/whitespace-tolerant fallback and a joined-2-line-window fallback; `parse_bifa()`'s rigid ">=8 tokens" threshold relaxed; a new `bo_profile_parser.py --diagnose <path>` command added for precise future debugging. Not fully re-verified against a real BO Profile PDF this session (none was supplied) — the false-REVIEW/silently-empty-sheet symptom is eliminated either way (see item 10), but the underlying parse success itself is unconfirmed.
10. "FY Total vs BIFA" had the FY hardcoded to the literal string "2022-23", and treated an empty BIFA lookup as if it were a verified zero, producing false REVIEW flags on every line. Fixed: FY is now passed in dynamically; an empty/unmatched BIFA now shows explicit N/A with a red warning explaining why, never a fake comparison.
11. `folder_classifier.py`'s PDF-classification loop silently swallowed a `pdfplumber` import failure (inside a bare `except Exception`), which — when `pdfplumber` genuinely wasn't installed in the person's venv — misclassified EVERY PDF (GSTR-9, GSTR-9C, BO Profile, scanned BS/PL) as "BS/PL," making GSTR-9/GSTR-9C appear as "not supplied" even though the files were present. Root cause confirmed on the person's real machine: a corrupted/copied venv where `pip` pointed at a nonexistent old project path. Fixed: pdfplumber's availability is checked once, loudly, before the loop; a missing library now produces an unmissable `[WARNING]` instead of silent misclassification.
12. `bs_pl_input.py`'s BS/PL rule engine (R0-R12) was built and unit-tested standalone, but never actually called from `master_build.py`'s real pipeline — found during a genericity audit; would never have appeared in the output workbook regardless of whether `bs_pl_input.py` was filled in. Fixed and wired in, with the GSTIN-tag safety check (§7.6) added at the same time.
13. `mcp-india-stack`'s HSN data, even freshly refreshed, is 99.9% all-zero-rate — discovered by inspecting a real refresh export the person ran locally; not a bug in this tool's own code, but a critical data-quality finding about the third-party source that changed how much this tool can rely on it (§7.2). Responded to by adding an all-zero-rate defensive filter to the loader and being explicit throughout the tool and this document about the resulting near-zero real coverage.

Two recurring tool-authoring bugs worth noting for future development in a new chat: on at least three occasions this session, a `str_replace` edit accidentally dropped the `def` line of the function immediately following the edited block (leaving an orphaned docstring one level deeper than its `def`). Always re-view the file immediately after any multi-line `str_replace` near a function boundary and confirm the next function's `def` line survived, before moving on.

---

## 9. Any-taxpayer readiness — what's generic vs what needs manual updating

Fully generic, zero editing needed for a new taxpayer: GSTIN/company/FY/month detection, output filename, all core document parsers, ARN/filing-compliance, cancelled-e-invoice detection, all monthly/cross-month/annual/forensic checks' logic, the HSN/SAC code-validity master (already covers 21,935+681 official codes for ANY taxpayer).

Needs manual updating per taxpayer:
- `bs_pl_input.py` — GSTIN-tagged, refuses to apply to a mismatched taxpayer (safe by construction, but needs real figures typed in per taxpayer to do anything at all).
- `HSN_RATE_HISTORY` in `hsn_fraud_checks.py` — currently covers only this taxpayer's 9 real HSN codes. A different industry needs its own codes added (10-15 minutes of research per code, following the exact pattern already in the table — see §7.5).
- `hsn_fraud_checks.py`'s older, unrelated fraud-pattern checks that reference specific HSN chapters (a handful of the ~25 named fraud patterns assume pharma-chapter context) may need review for a very different industry — not audited this session, flagged as a possible gap.

---

## 10. What's tested vs what's architecturally-ready-but-unverified — honest status

Tested against real data, high confidence:
- `annual_return_parser.py`'s GSTR-9/GSTR-9C/Table-8A parsers — every extracted figure matches independently-computed reference numbers exactly (see §11).
- `forensic_checks.py`'s R13/R14/R0/R1/R12 — same, exact matches.
- Every graceful-degradation fix in §8 (items 1-4, 6-13) — isolated unit tests, all passing, several re-verified against the exact real-world text/data that originally exposed the bug.
- The HSN rate-reference system's three sources (§7.1-7.3) — data quality independently verified against real exports this session (including the critical mcp-india-stack finding).
- The new per-month HSN Rate Review table (§6.8) — tested standalone with match/mismatch/unknown-code scenarios.

Architecturally built, NOT yet run end-to-end against real data:
- Multi-year (2+ FY) support — no real multi-year folder was supplied this session; the file-discovery and month-ordering logic is reasoned through carefully but unproven live. Three sub-layers (36-check HSN/Fraud layer, annual-reconciliation layer, GSTR-9/9C/Table-8A forensic layer) still resolve to the FIRST matched file per document type only — true multi-year support for those three needs per-FY looping, a contained follow-up. Irrelevant for single-FY use, which is this project's current actual use case.
- BO Profile's BIFA/Related-Party/Top-Counterparty sections — hardened defensively (§8 item 9), but real parse success not re-confirmed; run `bo_profile_parser.py --diagnose <real_path>` to close this out precisely.
- QRMP due-date category (Category X/Y state list) in `filing_compliance.py` — written from general knowledge, not cross-verified against the exact current CBIC notification; low-risk since this taxpayer is a confirmed monthly filer (this code path isn't exercised).
- ARN-date extraction from a REAL merged GSTR-1 file — tested only against a hand-built synthetic workbook using the exact real marker format the person confirmed; the general mechanism is verified, but a fresh real file should be spot-checked once available.

---

## 11. Verified results (exact numbers reproduced from real files this session)

- Table 8A: 2,534 total B2B rows, 2,408 "ITC available = Yes." CGST Rs 97,11,690.70, SGST Rs 97,11,690.70, IGST Rs 2,81,96,772.63, Total Rs 4,76,20,154.03.
- GSTR-9: Table 6A (ITC via 3B) total Rs 4,78,88,427.54. Table 4B taxable Rs 29,98,31,494.95. Table 9 liability IGST Rs 2,90,78,901.00.
- GSTR-9C: Turnover Rs 46,56,39,087.14. Exempt/nil/non-GST adjustment Rs 16,70,77,200.00. ITC per books Rs 4,68,38,857.26. ARN AA0503234208977, ARN Date 30-12-2023.
- R14 (four-way ITC): 3B minus Books = Rs 10,49,570.28; 2B minus Books = Rs 9,39,250.40; 8A minus Books = Rs 7,81,296.77 — exact matches to the independently-produced forensic analysis this project started from.
- R13: correctly flags the Rs 16.7 Cr unsupported-exempt-turnover scenario.
- R0/R1/R12 (BS/PL): all PASS, exact-match figures.
- mcp-india-stack data quality: 22,471 of 22,500 rows (99.9%) all-zero-rate, confirmed on a real fresh `--refresh-all` export.
- HSN/SAC master: 21,935 HSN + 681 SAC codes, zero duplicates, all 9 of this taxpayer's curated codes present.

---

## 12. Complete output workbook structure

| Sheet(s) | Content |
|---|---|
| Master Dashboard | Every FLAG/MISMATCH/REVIEW across ALL layers (monthly + HSN/Fraud + Forensic R13/R14 + Cancelled E-Invoices), ranked together |
| `Comparison <Month>` x N | Per-month raw comparison, plus the new HSN Rate Review table appended below it (§6.8) |
| `Analysis14 <Month>` x N | Per-month 14-check findings |
| `EWB <Month>` / `EWB Detail <Month>` x N | Per-month 27-check EWB findings |
| Doc-Series Integrity | Table-13 vs actual, cross-referenced against declared cancellations + cancelled e-invoices |
| Rectification Pairs | Amendments linked to original month + DRC payments |
| HSN & Fraud Pattern Checks | All 36+ checks' findings (unchanged this final round, per explicit instruction) |
| Filing Compliance & Late Fee | Per-month ARN, filed date, due date, Rs late fee + interest |
| Forensic Checks (R13-R14) | Turnover-gap rule + four-way ITC reconciliation |
| Cancelled E-Invoices | Cross-check findings + full list |
| Annual Cover & Caveats | Sources, known limitations, departmental proceedings |
| Annual Ledger Walkthrough | Monthly ledger/TPST/portal reconciliation |
| FY Total vs BIFA | FY totals vs department's BIFA figures |
| Related-Party Alerts / Top Counterparties | BO Profile lists |

---

## 13. Recommended next steps (priority order, for whoever continues this project)

1. Real BO Profile PDF (or its `--diagnose` output) — the single highest-value remaining verification gap (§10).
2. A second FY's worth of merged files — proves the multi-year architecture end-to-end instead of by code review alone.
3. Extend `HSN_RATE_HISTORY` the moment a new taxpayer/industry is assessed — 10-15 minutes of targeted research per new HSN code, following the exact pattern already there.
4. Re-run `master_build.py` on a real multi-month GSTR-1/3B set to spot-check the ARN-date extraction (§10) against real marker text end-to-end, not just the synthetic reproduction.
5. If BS/P&L becomes a recurring need, ask the taxpayer's accountant for a text-based (non-scanned) export — would let §7.6's OCR limitation be lifted using the same watermark-tolerant PDF-text technique already proven on GSTR-9/GSTR-9C.
