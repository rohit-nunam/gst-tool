#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BS/PL STRUCTURED INPUT  --  fill this in by hand from the real PDF/Excel,
then pass MRHC_BS_PL (or your own copy of this dict) to
forensic_checks.check_bs_pl_rules(). See OCR_LIMITATION.md for why this is
a typed-in dict and not an auto-OCR'd one.

Every key is OPTIONAL -- only fill in what you can read off the real
document. check_bs_pl_rules() emits an explicit "not tested -- <key> not
supplied" finding for anything left out, so the final report always shows
a complete checklist (per the framework's own instruction in section 2.6),
never a silently-skipped row.

Values are {"fy_prior": <number>, "fy_current": <number>} pairs. Use the
SAME two financial years as the rest of the run (this taxpayer: FY21-22 as
"fy_prior", FY22-23 as "fy_current" -- matching the two columns printed on
the real Balance Sheet/P&L).
"""

# Filled in below from MRHC_PL_AND_BS_FY_22-23.pdf, transcribed by hand
# (NOT OCR -- see OCR_LIMITATION.md) directly from the document text visible
# in this session. Cross-checked: Total Assets = Total Equity+Liab for BOTH
# years (R0 passes), matching the document's own totals row exactly.
#
# GENERIC NAME (BS_PL_DATA, not a taxpayer-specific name): master_build.py
# tries to import this exact variable from this exact file -- for a NEW
# taxpayer, replace the values below with their real figures (rename
# nothing, keep the variable name BS_PL_DATA so the pipeline picks it up
# automatically), or empty the dict entirely to skip R0-R12 cleanly.
BS_PL_DATA = {
    # SAFETY TAG (checked by master_build.py before use): this dict is only applied if it
    # matches the GSTIN actually being processed this run. Prevents a stale/wrong-taxpayer's
    # BS/PL figures from silently being used against a different taxpayer's GST returns.
    "_gstin": "05AAECM6380J1ZA",
    "total_assets": {"fy_prior": 22_88_22_678.61, "fy_current": 30_39_29_544.01},
    "total_equity_liab": {"fy_prior": 22_88_22_678.61, "fy_current": 30_39_29_544.01},

    "share_capital": {"fy_prior": 5_71_70_067.00, "fy_current": 5_71_70_067.00},
    "reserves_and_surplus": {"fy_prior": 3_94_91_383.22, "fy_current": 8_76_77_882.15},

    "trade_payables": {"fy_prior": 8_64_87_870.59, "fy_current": 7_21_81_503.53},
    "short_term_provisions": {"fy_prior": 79_74_340.00, "fy_current": 2_04_27_100.00},

    "fixed_assets_tangible": {"fy_prior": 2_71_88_489.34, "fy_current": 2_50_97_042.97},
    "non_current_investments": {"fy_prior": 16_22_500.00, "fy_current": 2_23_20_050.00},
    "inventories": {"fy_prior": 11_54_36_693.00, "fy_current": 5_35_67_690.00},
    "trade_receivables": {"fy_prior": 5_82_00_987.98, "fy_current": 6_92_71_811.89},

    "revenue_from_operations": {"fy_prior": 32_26_17_617.19, "fy_current": 46_56_39_087.14},
    "other_income": {"fy_prior": 2_71_66_064.25, "fy_current": 6_38_593.04},
    "other_expenses": {"fy_prior": 5_02_91_468.33, "fy_current": 4_53_37_472.48},
    "finance_costs": {"fy_prior": 15_61_947.50, "fy_current": 49_73_007.06},
    "depreciation": {"fy_prior": 32_79_627.00, "fy_current": 37_15_267.00},
    "net_profit_after_tax": {"fy_prior": 2_17_06_887.75, "fy_current": 4_81_86_498.93},
}


if __name__ == "__main__":
    import forensic_checks as fchk
    import annual_return_parser as arp

    gstr9c = arp.parse_gstr9c("/mnt/user-data/uploads/1783794742279_GSTR-9C_05AAECM6380J1ZA_032023.pdf")
    findings = fchk.check_bs_pl_rules(BS_PL_DATA, gstr9c=gstr9c, bo_profile=None)
    for f in findings:
        print(f"[{f.severity:8}] {f.ref:6} {f.title}")
        print(f"           {f.detail[:200]}")
