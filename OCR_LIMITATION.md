# Why the Balance Sheet / P&L PDF is NOT auto-parsed

## What was tested

`MRHC_PL_AND_BS_FY_22-23.pdf` was opened with `pdfplumber` — **zero extractable
text on either page, two raster images per page**. This confirms it is a
**scanned document** (the file itself carries a "Scanned with CamScanner"
watermark), not a digitally-generated PDF like the GSTR-9/GSTR-9C exports.

To test feasibility, both pages were rendered at 3x scale and run through
`pytesseract` (Tesseract OCR). Result, compared line-by-line against the
correct figures:

| Line item | Correct value (from the document) | OCR output | Correct? |
|---|---|---|---|
| Finance Costs (FY23) | 49,73,007.06 | 43,73,007.05 | **NO — wrong by Rs 6,00,000** |
| Current tax expense | 1,86,70,817.00 | 1,86,70,817.60 | NO — wrong by 60 paise |
| Most other rupee figures | — | — | Correct |
| Row-label ↔ value alignment (P&L page) | — | Several labels and values became separated onto different lines during OCR | Structurally unreliable |

Two real, silent digit-errors in a ~30-figure sample, one of them large enough
(Rs 6 lakh) to itself **generate a false forensic FLAG** in the R10 (Finance
Costs) check — exactly the failure mode this tool is built to avoid everywhere
else (see the "no safety net" rule and the historical A1/A3 bugs documented
in `GST_Tool_Full_Project_Context.md`, both of which were caught by refusing
to trust an unverified number).

## Decision

**A forensic/scrutiny tool cannot silently trust OCR'd rupee figures.** Every
other parser in this codebase is content-based but never *guesses* a number —
it reads it verbatim from a real cell or a clean regex match on machine-
generated text. OCR on a scanned image is fundamentally different: it is a
*statistical guess* at what a digit probably is, and it will occasionally
guess wrong with no way to detect which figures are the wrong ones.

## Two supported paths forward

**Path A — structured input (built, ready now).** `bs_pl_input.py` in this
package is a plain-Python template: fill in the line items you can read
directly off the PDF (or off a proper trial balance / Tally export) as a
dict of `{fy_prior, fy_current}` pairs. `forensic_checks.check_bs_pl_rules()`
runs the full R0–R12 rule engine against it. This is the same effort as
transcribing the figures once, and every number in the output is then
provably correct (typed by a human who can see the real PDF at full
resolution, not guessed by a compression-lossy OCR pass on a phone-scanned
image).

**Path B — ask for a text-based export.** If the company's accountant can
re-export the Balance Sheet/P&L from Tally/Zoho/whatever ERP as a native PDF
or Excel (not a CamScanner photo), `annual_return_parser.py`'s own
`_extract_pdf_text()` + label-matching approach (already proven on the real
GSTR-9/GSTR-9C exports, including tolerating their watermark noise) would
extract every figure with the same verified accuracy as those two documents
— zero OCR risk, because there would be real text in the PDF to read.

## What this means for this specific taxpayer's Part 2 findings

Every number quoted in `GST_Forensic_Comparison_Framework_v1.md` Part 2
(§2.1–2.6) was manually transcribed by a human reading the scanned PDF, not
OCR'd — so those specific figures are trustworthy as given. If you want the
**R0–R12 rule engine** (not just the one-off manual findings already in that
document) to run against this taxpayer's real numbers, fill in
`bs_pl_input.py`'s dict from the same PDF (10 minutes of typing, ~12 line
items) and pass it to `forensic_checks.check_bs_pl_rules()` — see that
module's docstring for the exact call.
