# Bug-fix round — QA validation feedback

Five bugs were reported by the verification team after testing against real
data files. Status of each:

---

## 🔴 Bug 1 — Cancelled E-Invoices detection completely failed — **FIXED + VERIFIED**

**Root cause (two separate bugs, both in `gst_scrutiny_tool.parse_einv()`):**
1. Header matching was **case-sensitive exact-dict-key lookup**. The real
   column is named `E-invoice status` (lowercase i, lowercase s); the
   candidate list had `E-Invoice Status` (capital I, capital S) — these
   never matched even though the right name was "in the list".
2. Deeper bug: even once a status column IS found, the code was adding
   **every row's** taxable/IGST/CGST/SGST/CESS/count into the main totals
   and the `lines` dict used for every E-Invoice-vs-GSTR-1 comparison —
   cancelled rows included. Since a cancelled e-invoice is correctly ABSENT
   from GSTR-1 (marked "Deleted" there), including it on the E-Invoice side
   manufactured a false gap on every single cancelled invoice.

**Fix:** (a) header matching is now case-insensitive against a normalized
header map; (b) cancellation status is checked **before** any totals are
accumulated — a cancelled row is recorded separately in `out["cancelled"]`
and `continue`s past the totals/lines accumulation entirely.

**Tested:** synthetic workbook using the exact real header name
(`E-invoice status`, values `Valid`/`Cancelled`) — confirms the column is
now found, the cancelled row is captured correctly, and critically that it
no longer pollutes the main taxable/IGST/count totals.

**Linked fix — Doc-Series Integrity:** the Aug-22 `MR22-23/0226`-style
"genuinely unexplained" gap is a direct downstream effect of Bug 1 — Table
13 declares the serial, GSTR-1 correctly omits it (cancelled), but nothing
previously cross-referenced that omission against the cancellation reason.
New `forensic_checks.enrich_doc_gap_with_cancelled_einvoices()` cross-checks
every "missing" Doc-Series serial against that month's cancelled-e-invoice
list; a match now shows as **"EXPLAINED BY CANCELLED E-INVOICE"** (amber,
not red) instead of "UNEXPLAINED — REVIEW". Tested with a synthetic
reproduction of the exact Aug-22 scenario reported.

---

## 🔴 Bug 2 — GSTR-1 ARN filing-date not extracted — **FIXED + VERIFIED**

**Root cause:** the marker-parsing regex in `filing_compliance.py` only
recognised the date label in "Date [of] [Filing/ARN]" word order. The real
marker text uses **"ARN Date"** (ARN comes first): `ARN: AA050422057237G |
ARN Date: 10-05-2022` — a word order the original regex's optional group
never covered, so the ARN number matched but the date group always came
back empty.

**Fix:** rewrote the label-matching group to accept `ARN Date`, `Date of
Filing`, `Filing Date`, or bare `Date`, in any of these word orders.

**Tested:** reproduced the exact bug with the exact real marker text quoted
in the bug report first (confirmed `date=None`), then re-tested after the
fix (confirmed `date=2022-05-10`) — plus 3 more word-order variants, all
passing.

---

## 🔴 Bug 3 — "FY Total vs BIFA" sheet: BIFA column always 0 — **PARTIALLY FIXED, NOT FULLY VERIFIED (no real BO Profile PDF available this round)**

**What was fixed with confidence:**
- `build_annual_workbook.write_fy_total_vs_bifa()` had the FY **hardcoded**
  to the literal string `"2022-23"` for the `bifa_by_fy` lookup — a
  genericity bug (would silently break for any other taxpayer/FY) fixed
  regardless of whether it's the direct cause here. Now takes the real
  detected FY as a parameter.
- **The specific symptom reported — every BIFA figure showing 0 and every
  row flagging REVIEW — is now impossible.** Previously, an empty
  `bifa_by_fy` dict (FY not found) made every `bifa.get(x) or 0` silently
  become `0`, which then got diffed against the real recomputed value as
  if it were a genuine department figure. Now, when `bifa` is empty, every
  BIFA-side value is explicitly `None` ("n/a"), the Check column shows
  "N/A" (not a fake REVIEW), and the sheet prints an explicit red warning
  naming which FY was looked for and which FYs (if any) were actually
  found in the BO Profile.
- `bo_profile_parser._section_bounds()` (used by `parse_bifa()` and every
  other BO-Profile table parser) was hardened: case-insensitive/whitespace-
  normalized matching, plus a fallback that searches a joined 2-line window
  in case a section header got split across a line break during PDF text
  extraction (a plausible, common cause of "the parser suddenly finds
  nothing" on a real-world PDF that a synthetic test file wouldn't surface).
- `parse_bifa()`'s row-acceptance threshold was relaxed from "must have
  ≥8 numeric tokens or silently drop the whole FY" to "≥6 required, tail
  EWB-related fields individually degrade to `None` if genuinely absent" —
  the previous all-or-nothing threshold could have been discarding a real
  row that had, say, 6-7 tokens because an optional column produced no
  token at all (not even a blank placeholder) in the real extraction.

**What could NOT be verified:** I do not have your real BO Profile PDF this
round, so I cannot confirm whether `parse_bifa()` now actually FINDS the
section in your specific file (as opposed to just being more tolerant in
general). **If the underlying section-detection is still failing after this
fix, the sheet will now say so explicitly (per the point above) instead of
showing a misleading 0 — so at minimum the false-REVIEW symptom is
resolved even if the root parse still needs another round.**

**To close this out with full certainty:** run
`python3 bo_profile_parser.py --diagnose <your_BO_Profile.pdf>` (new
diagnostic command, ships in this package) and paste the output back, or
share the PDF directly — this will show exactly which sections are found
vs not, with a text preview, so the fix (if still needed) can be precise
rather than another guess.

---

## 🔴 Bug 4 & 5 — Related-Party Alerts / Top Counterparties sheets empty — **SAME ROOT-CAUSE FAMILY AS BUG 3, SAME CAVEAT**

Both sheets are driven by `bo_profile_parser.parse_related_party()` and
`parse_top_list()`, which depend on the same `_section_bounds()` function
hardened above, PLUS two very rigid multi-field regexes (`REL_PARTY_RE`,
`TOP_LIST_RE`) that require an exact sequence of sub-fields (dates in a
specific format, status as exactly `Active`/`Cancelled`, risk label as
exactly `LOW`/`MEDIUM`/`HIGH`, etc.). If your real PDF's text extraction
produces even a slightly different token order or spacing for these rows,
the whole regex fails to match ANY row — consistent with both sheets being
**completely** empty rather than partially populated.

**What was fixed:** the `_section_bounds()` hardening above applies here
too (same shared dependency). Both sheets now also print an explicit note
when zero rows are found — *"either genuinely none in the source BO
Profile, or this section's marker text wasn't found during parsing"* —
instead of an empty sheet that looks identical to a genuinely clean result.

**What could NOT be fixed with confidence:** `REL_PARTY_RE` and
`TOP_LIST_RE` themselves were **not** rewritten, because doing so blind
(without seeing your real PDF's actual extracted text for these rows) risks
producing a *differently* wrong regex rather than a correct one — exactly
the kind of unverified guess this tool's whole design philosophy exists to
avoid. This is the most honest thing to do given the constraint, not a
shortcut.

**To close this out:** the same `--diagnose` command above will show the
raw lines for these sections too (or share the actual PDF) — with the real
text in front of me, fixing these two regexes precisely is a contained,
fast follow-up, the same way Bugs 1/2 were fixed with full certainty once
the exact real text was known.

---

## Net effect of this round

- Bugs 1 and 2: **fully fixed, fully tested, high confidence** — these were
  precise, reproducible bugs with exact real-world text provided in the
  report, so the fix could be verified the same way Table 8A/GSTR-9/GSTR-9C
  were in the previous round.
- Bugs 3/4/5: **defensively hardened + the misleading-output symptom is
  eliminated either way** (false REVIEW flags / silently-empty sheets both
  now come with an explicit "here's why this might be incomplete" note
  instead of looking like a clean verified result) — but the underlying
  section-detection fix itself is **not guaranteed** without the real BO
  Profile PDF (or the `--diagnose` output) to verify against.
