#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
RUN ONE MONTH  --  drives the three existing single-month engines
(gst_scrutiny_tool, gst_analysis_checks, gst_eway_recon) for one period,
now reading that period's data out of the MERGED (whole-FY) workbooks
identified by folder_classifier.py, instead of one file per month.

Two things changed from the original per-month-file design:
  1. The merged-workbook paths are the SAME for every month (there is only
     one file per document type) -- only PERIOD_LABEL changes between calls.
     Every parser below (parse_gstr1, parse_gstr3b, parse_einv, 2B, amendments,
     EWB invoice readers) now takes an explicit `month` argument and reads
     only that month's block out of the shared file.
  2. E-Way-Bill input still comes from the whole-FY annual workbooks, filtered
     down to this month via ewb_annual_parser.filter_by_month() -- unchanged
     from before, since those files were never per-month to begin with.

CHANGELOG (this revision):
  - ewb_out_file_supplied / ewb_in_file_supplied now passed through to
    gst_eway_recon.run() so a totally-absent EWB direction produces an
    honest SKIP instead of misleading PASS/REVIEW noise (see that module's
    docstring).
  - Cancelled-e-invoice rows (gst_scrutiny_tool.parse_einv()'s new
    'cancelled' key) are collected here per month and returned in the
    result dict, for the master build to aggregate FY-wide.
  - ARN dates (filing_compliance.py) are looked up per month here (once
    per run via the cache passed in by master_build.py, not re-read from
    disk every month) and the resulting late-fee/interest record is
    returned in the result dict for check #10/#8 to actually use --
    previously always fell through to INFO because GSTR1_FILING_DATE/
    GSTR3B_FILING_DATE were never actually set anywhere in this pipeline.
"""

import gst_scrutiny_tool as raw
import gst_analysis_checks as ana
import gst_eway_recon as eway
import gstr2b_parser as g2b
import ewb_annual_parser as ewbp
import amendments as amd
import filing_compliance as fc


def run_month(month_label, files, ewb_out_annual_rows, ewb_in_annual_rows,
              self_gstin, company_name, ewb_out_file_supplied=True, ewb_in_file_supplied=True,
              gstr1_arn_by_month=None, gstr3b_arn_by_month=None,
              gstr1_is_qrmp=False, gstr3b_is_qrmp=False, annual_turnover=None):
    """files: {'gstr1':path,'gstr3b':path,'einv':path or None,'gstr2b':path or None}
    -- these are the MERGED workbook paths, the same on every call; only
    month_label changes which block gets read out of each of them.

    gstr1_arn_by_month / gstr3b_arn_by_month: pre-computed once per run (not
    per month -- these read the WHOLE merged file each time) by the caller
    via filing_compliance.gstr1_arn_dates_by_month()/gstr3b_arn_dates_by_month(),
    then passed in here so this function stays a pure per-month reader like
    every other parser call in this file."""
    g1path, g3bpath, einvpath, twobpath = (files.get("gstr1"), files.get("gstr3b"),
                                            files.get("einv"), files.get("gstr2b"))

    # ---- set the shared modules up for this period ----
    # (file paths are the same across months; only PERIOD_LABEL actually changes)
    raw.GSTR1_FILE = g1path
    raw.GSTR3B_FILE = g3bpath
    raw.EINV_FILE = einvpath
    raw.GSTR2B_FILE = twobpath
    raw.SELF_GSTIN = self_gstin
    raw.COMPANY_NAME = company_name
    raw.PERIOD_LABEL = month_label
    eway.SELF_GSTIN = self_gstin
    eway.GSTR2B_FILE = twobpath

    # ---- filing compliance: ARN dates + late fee/interest for THIS month ----
    # (fixes the previously-broken/unwired legacy extraction -- see filing_compliance.py)
    compliance = None
    if gstr1_arn_by_month is not None or gstr3b_arn_by_month is not None:
        compliance = fc.month_filing_compliance(
            month_label, gstr1_arn_by_month or {}, gstr3b_arn_by_month or {},
            gstr1_is_qrmp=gstr1_is_qrmp, gstr3b_is_qrmp=gstr3b_is_qrmp, self_gstin=self_gstin,
            annual_turnover=annual_turnover)
        # feed the real filing dates into gst_analysis_checks' checks #8/#10
        # (CONFIG-based, same mechanism the codebase already had -- just actually populated now)
        raw.GSTR1_FILING_DATE = compliance.get("gstr1_filing_date")
        raw.GSTR3B_FILING_DATE = compliance.get("gstr3b_filing_date")

    # ---- pipeline 1: comparison + Sooraj's 14 checks (unchanged engines,
    #      now called with the explicit month) ----
    comparisons, comp_raw = raw.build_comparisons()
    g1 = comp_raw["g1"]; g3b = comp_raw["g3b"]; einv = comp_raw["einv"]; b2b = comp_raw["b2b"]
    g1_lines = ana.read_gstr1_lines(g1path, month_label)
    einv_lines = ana.read_einv_lines(einvpath, month_label) if einvpath else []
    findings14 = ana.run_checks(g1, g3b, einv, b2b, g1_lines, einv_lines)

    # ---- pipeline 2: E-Way-Bill 27-check matrix, fed from the ANNUAL EWB lists
    #      filtered to this month (by EWB date) -- unchanged ----
    ewb_out = ewbp.filter_by_month(ewb_out_annual_rows, month_label)
    ewb_in = ewbp.filter_by_month(ewb_in_annual_rows, month_label)
    g1inv = eway.read_gstr1_invoices(g1path, month_label)
    einv_ew = eway.read_einv_invoices(einvpath, month_label) if einvpath else {}
    b2b_ew = g2b.summary_for_month(twobpath, month_label)
    findings27 = eway.run(ewb_out, ewb_in, g1inv, einv_ew, g3b, b2b_ew,
                           ewb_out_file_supplied=ewb_out_file_supplied,
                           ewb_in_file_supplied=ewb_in_file_supplied)

    # ---- amendments + doc-series-integrity (this month's OWN GSTR-1 block only;
    #      cross-month linkage to the ORIGINAL month happens in master_build) ----
    b2ba = amd.parse_b2ba(g1path, month_label)
    cdnra = amd.parse_cdnra(g1path, month_label)
    docs = amd.parse_docs(g1path, month_label)
    actual_invnos = set(k[0] for k in g1.get("lines", {}).keys())
    doc_gap = amd.doc_series_gap_check(docs, actual_invnos)

    # ---- cancelled e-invoices this month (new) ----
    cancelled_this_month = einv.get("cancelled", []) if einv.get("available") else []
    einv_column_found = einv.get("cancel_col_found", False) if einv.get("available") else False

    return dict(
        month=month_label, comparisons=comparisons, comp_raw=comp_raw,
        findings14=findings14, findings27=findings27,
        b2ba=b2ba, cdnra=cdnra, docs=docs, doc_gap=doc_gap,
        compliance=compliance,
        cancelled_einvoices=cancelled_this_month, einv_cancel_col_found=einv_column_found,
        g1_named_invnos=set(k[0] for k in g1.get("lines", {}).keys() if k[0]),
        meta=dict(
            ewb_out_n=len(ewb_out), ewb_in_n=len(ewb_in),
            twob_src=b2b_ew.get("_source"), twob_file=b2b_ew.get("_file"),
            twob_available=b2b_ew.get("available", True),
            einv_file=einvpath,
        ),
    )
