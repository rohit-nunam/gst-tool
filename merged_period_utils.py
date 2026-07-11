#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MERGED PERIOD UTILS
====================
Shared helpers for the new MERGED input files (one workbook per document
type covering many months, instead of one file per month).

Every GSTR-1 / E-Invoice / GSTR-2B sub-sheet in a merged workbook carries a
period-MARKER row before each month's (or quarter's) block of data, e.g.:

    "Financial Year: 2022-23  |  Tax Period: January  |  ARN: ..."          (GSTR-1)
    "Financial Year: 2022-23  |  Tax Period: 042022  |  Date Updated ..."   (E-Invoice, numeric MMYYYY)
    "Financial Year: 2022-23  |  Tax Period: Apr-Jun  |  Date of Gen ..."   (GSTR-2B, quarterly)

This module finds those marker rows by CONTENT (never by position/sheet name)
and slices the sheet into {month_label: [data_rows]} blocks. GSTR-3B is
different -- it merges as one SHEET PER MONTH rather than marker rows inside
one sheet -- so GSTR-3B sheets are identified by their own in-sheet
'Year'/'Tax Period' key-value rows (content-based, per user instruction:
ignore the sheet's NAME entirely, e.g. 'Jan_2022-23' is not to be trusted).

HARD RULE (per explicit instruction): no safety nets. If a marker can't be
parsed, or a requested month isn't present, this raises -- it does not
silently return zero/empty/a guess.
"""

import re

MONTH_NAME_TO_NUM = {
    "JANUARY": 1, "FEBRUARY": 2, "MARCH": 3, "APRIL": 4, "MAY": 5, "JUNE": 6,
    "JULY": 7, "AUGUST": 8, "SEPTEMBER": 9, "OCTOBER": 10, "NOVEMBER": 11, "DECEMBER": 12,
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "JUN": 6, "JUL": 7, "AUG": 8,
    "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}
MONTH_NUM_TO_ABBR = {1: "Apr", 2: "May", 3: "Jun", 4: "Jul", 5: "Aug", 6: "Sep",
                     7: "Oct", 8: "Nov", 9: "Dec", 10: "Jan", 11: "Feb", 12: "Mar"}
# ^ deliberately NOT used -- keep a plain calendar map instead (clearer, no FY-offset tricks here)
CAL_MONTH_ABBR = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
                   7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}

QUARTER_TO_MONTHS = {
    "JAN-MAR": [1, 2, 3], "APR-JUN": [4, 5, 6], "JUL-SEP": [7, 8, 9], "OCT-DEC": [10, 11, 12],
}

MARKER_RE = re.compile(
    r"Financial Year:\s*([0-9]{4}\s*-\s*[0-9]{2,4})\s*\|\s*Tax Period:\s*([^\|]+?)\s*(?:\||$)"
)


class PeriodParseError(ValueError):
    pass


def fy_years(fy):
    """'2022-23' -> (2022, 2023). Raises PeriodParseError if malformed."""
    m = re.match(r"^\s*(\d{4})\s*-\s*(\d{2,4})\s*$", fy)
    if not m:
        raise PeriodParseError(f"Unrecognised Financial Year format: {fy!r}")
    y1 = int(m.group(1))
    y2s = m.group(2)
    y2 = int(y2s) if len(y2s) == 4 else int(str(y1)[:2] + y2s)
    return y1, y2


def months_for_tax_period(fy, tax_period):
    """Return list of 'Mon-YY' calendar-month labels this marker covers.
    1 label for a month marker, 3 for a quarter marker. Raises PeriodParseError
    if the Tax Period text isn't in any recognised format (month name, numeric
    MMYYYY, or a quarter like 'Apr-Jun') -- stays flexible across formats
    since real files may use any of the three, but does not guess beyond them."""
    y1, y2 = fy_years(fy)
    tp = tax_period.strip().upper()

    if re.match(r"^\d{6}$", tp):                 # numeric MMYYYY, e.g. '042022'
        month_nums = [int(tp[:2])]
        if not (1 <= month_nums[0] <= 12):
            raise PeriodParseError(f"Unrecognised numeric Tax Period: {tax_period!r}")
    elif tp in MONTH_NAME_TO_NUM:                 # 'January' / 'Jan' style
        month_nums = [MONTH_NAME_TO_NUM[tp]]
    elif tp in QUARTER_TO_MONTHS:                 # 'Apr-Jun' style
        month_nums = QUARTER_TO_MONTHS[tp]
    else:
        raise PeriodParseError(f"Unrecognised Tax Period format: {tax_period!r} (FY {fy})")

    labels = []
    for mm in month_nums:
        cal_year = y2 if mm <= 3 else y1
        labels.append(f"{CAL_MONTH_ABBR[mm]}-{str(cal_year)[2:]}")
    return labels


def parse_marker_text(text):
    """Return (fy, tax_period_raw, [month_labels]) for one marker cell's text.
    Raises PeriodParseError if the text is not a period marker at all."""
    m = MARKER_RE.search(text or "")
    if not m:
        raise PeriodParseError(f"Not a period-marker cell: {text!r}")
    fy, tp = m.group(1).strip(), m.group(2).strip()
    return fy, tp, months_for_tax_period(fy, tp)


def is_marker_row(row):
    """A marker row carries its text in cell 0 only; every other cell is empty."""
    if not row or not row[0]:
        return False
    return bool(MARKER_RE.search(str(row[0])))


def split_rows_by_month(data_rows):
    """data_rows: rows AFTER the header row (may start with a marker).
    Returns {month_label: [row, row, ...]}, excluding marker rows themselves
    and genuinely blank spacer rows. A quarter marker fans its rows out into
    all 3 of that quarter's month buckets (used for GSTR-2B's quarter-level
    summary sheet; invoice-level 2B sheets should instead use each row's own
    per-line period column -- see gstr2b_parser.py).
    A month whose block has ZERO data rows (its marker is immediately
    followed by the next marker) still gets registered with an empty list --
    that is a legitimate 'this month had nothing on this sub-sheet' state,
    not a missing month.
    Raises PeriodParseError if data is found before any marker has been seen."""
    blocks = {}
    current_labels = None
    for row in data_rows:
        if is_marker_row(row):
            _, _, current_labels = parse_marker_text(str(row[0]))
            for lbl in current_labels:
                blocks.setdefault(lbl, [])
            continue
        if not any(c not in (None, "") for c in row):
            continue
        if current_labels is None:
            raise PeriodParseError(
                "Data row encountered before any period marker in this sheet -- "
                "cannot determine which month it belongs to: " + repr(row)
            )
        for lbl in current_labels:
            blocks.setdefault(lbl, []).append(row)
    return blocks


def rows_for_month(all_rows, header_row_idx, month_label):
    """Convenience wrapper: split rows[header_row_idx+1:] by month, return only
    the requested month's rows. Raises PeriodParseError if that month has no
    block in this sheet at all (distinct from the month having zero DATA rows,
    which is legitimate and returns [])."""
    blocks = split_rows_by_month(all_rows[header_row_idx + 1:])
    if month_label not in blocks:
        raise PeriodParseError(
            f"Month {month_label!r} not found as a period marker in this sheet. "
            f"Months present: {sorted(blocks)}"
        )
    return blocks[month_label]


def find_block_for_month(all_rows, month_label):
    """For sheets where period markers sit directly among the data (no single
    fixed header row to split from -- e.g. GSTR-2B's 'ITC Available' summary,
    which has several small tables per quarter block), return (start, end)
    row-index bounds (start is the row right after the marker; end is the
    next marker's row, or len(all_rows)) for the block covering `month_label`.
    Raises PeriodParseError if no marker in the sheet covers that month."""
    marker_positions = []  # (row_idx, [month_labels])
    for i, row in enumerate(all_rows):
        if is_marker_row(row):
            _, _, labels = parse_marker_text(str(row[0]))
            marker_positions.append((i, labels))
    for idx, (row_idx, labels) in enumerate(marker_positions):
        if month_label in labels:
            start = row_idx + 1
            end = marker_positions[idx + 1][0] if idx + 1 < len(marker_positions) else len(all_rows)
            return start, end
    raise PeriodParseError(
        f"Month {month_label!r} not covered by any period marker in this sheet. "
        f"Markers found: {[lbl for _, lbl in marker_positions]}"
    )


def months_present(all_rows, header_row_idx):
    """Return sorted set of every month label found via markers in this sheet
    (including months whose block turned out to have zero data rows)."""
    months = set()
    current_labels = None
    for row in all_rows[header_row_idx + 1:]:
        if is_marker_row(row):
            _, _, current_labels = parse_marker_text(str(row[0]))
            months.update(current_labels)
    return months
