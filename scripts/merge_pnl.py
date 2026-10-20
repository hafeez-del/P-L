#!/usr/bin/env python3
"""
Merge single-store P&L "GL Historical Trend" exports into ONE formatted workbook.

 * one tab per store (tab name = store number + name)
 * report-title block ("Historical Trend", "For Period...", trend labels) and '_' spacer rows removed
 * months ordered oldest -> newest, plus a "N-Month Total" column
 * formatting:
     - section headers (MERCHANDISE SALES, INVENTORY ADJUSTMENT ...)  -> yellow, full-row (merged)
     - TOTAL rows                                                      -> blue fill, white bold text
     - a blank row between sections (new "vertical")
     - sub-headers under ADJUSTED OPERATING EXPENSES (LABOUR EXPENSE, OTHER LABOUR EXPENSE)
                                                                       -> bold, no fill, full-row
     - lines that have no header of their own (e.g. insurance/store supplies block after UTILITIES)
                                                                       -> plain bordered rows
     - NET INCOME                                                      -> big blue/white row
 * "Summary" tab: every vertical (all TOTAL lines + NET INCOME) for all stores, with live formulas

Usage:
    1) Script ke top par INPUT_DIR (aur chahein to OUTPUT_FILE) set karein, phir:
           python merge_pnl.py
    2) ya command line se:
           python merge_pnl.py <input_dir> [out.xlsx]
"""
import glob
import os
import re
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

# =====================================================================
#  >>>  YAHAN APNA PATH DALEN  <<<
#  INPUT_PATHS: folders / files ki list jahan store ki xlsx files hain
#  OUTPUT_FILE: final workbook kahan save ho (khali "" chhodo to INPUT_DIR me
#               "Combined_PnL_by_Store.xlsx" pehle folder me ban jayegi)
#  Windows path ke aage r zaroor lagayen (raw string), jaise r"C:\Users\..."
# =====================================================================
# INPUT_PATHS me jitne chahein folder YA seedhe file ke path dal sakte hain (comma se alag).
# Folder do to uski saari .xlsx files uthayi jayengi.
INPUT_PATHS = [
    r"C:\Users\YourName\Documents\PnL_Files",
    # r"D:\Another\Folder",
    # r"C:\Users\YourName\Downloads\Parker.xlsx",
]
OUTPUT_FILE = r""   # e.g. r"C:\Users\YourName\Documents\PnL_Files\Combined_PnL_by_Store.xlsx"
# =====================================================================

DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")

# headers that sit INSIDE a parent header and are shown as bold sub-headers (no fill)
SUBHEADERS = {"LABOUR EXPENSE", "OTHER LABOUR EXPENSE"}
# value rows that act as a section header in the report (shown as yellow header)
VALUE_HEADERS = {"NET VENDING INCOME"}

# ---- styles -------------------------------------------------------------
YELLOW = PatternFill("solid", fgColor="FFFF00")
BLUE = PatternFill("solid", fgColor="548DD4")
NAVY = PatternFill("solid", fgColor="1F3864")
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center")
NUM_FMT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'
F_DATA = Font(name="Verdana", size=9, color="000000")
F_BOLD = Font(name="Verdana", size=9, bold=True, color="000000")
F_TOTAL = Font(name="Verdana", size=9, bold=True, color="FFFFFF")
F_NET = Font(name="Verdana", size=14, bold=True, color="FFFFFF")
F_HDR = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
HDR_ROW = 4  # month header row on every store tab


def norm(s):
    return re.sub(r"\s+", " ", s or "").strip()


# ---- read one source file -----------------------------------------------
def load_store(path):
    ws = openpyxl.load_workbook(path).active
    data = list(ws.iter_rows(values_only=True))

    hdr_idx = next(i for i, r in enumerate(data)
                   if any(isinstance(v, str) and DATE_RE.match(v) for v in r[1:]))
    hdr = data[hdr_idx]
    date_cols = [j for j, v in enumerate(hdr) if isinstance(v, str) and DATE_RE.match(v)]
    total_col = next(j for j, v in enumerate(hdr) if isinstance(v, str) and "Total" in v)
    dates = [datetime.strptime(hdr[j], "%m/%d/%y") for j in date_cols]
    order = sorted(range(len(date_cols)), key=lambda k: dates[k])  # oldest -> newest

    top = [norm(r[0]) for r in data[:hdr_idx] if isinstance(r[0], str) and norm(r[0])]
    top = [t for t in top if "Historical" not in t and "For Period" not in t]
    store, address = top[0], (top[1] if len(top) > 1 else "")

    items = []
    for r in data[hdr_idx + 1:]:
        label = norm(r[0])
        if label in ("", "_"):
            continue                                   # blanks & '_' spacers are rebuilt later
        vals = [r[date_cols[k]] for k in order] + [r[total_col]]
        vals = [v if isinstance(v, (int, float)) else None for v in vals]
        up = label.upper()
        if all(v is None for v in vals) and up not in VALUE_HEADERS:
            kind = "subheader" if up in SUBHEADERS else "header"
        elif up == "NET INCOME":
            kind = "net"
        elif up.startswith("TOTAL"):
            kind = "total"
        elif up in VALUE_HEADERS:
            kind = "vheader"
        else:
            kind = "data"
        items.append({"label": label, "key": up, "vals": vals, "kind": kind})
    return dict(store=store, address=address, months=[dates[k] for k in order], items=items)


def blank_before(prev, kind):
    """Should a blank separator row sit between `prev` and `kind`?"""
    if prev is None or kind == "subheader":
        return False
    if kind in ("header", "vheader"):
        return prev not in ("header", "subheader")     # first thing under a parent header: no gap
    if kind == "total":
        return prev == "total"
    if kind == "net":
        return True
    if kind == "data":
        return prev in ("total", "net")                # headerless block starts after a total
    return False


# ---- cell helpers ---------------------------------------------------------
def fill_row(ws, row, ncols, fill=None, font=None, border=BOX):
    for c in range(1, ncols + 1):
        x = ws.cell(row, c)
        if fill: x.fill = fill
        if font: x.font = font
        x.border, x.alignment = border, CENTER


def full_row_header(ws, row, label, ncols, fill, font):
    ws.cell(row, 1, label)
    fill_row(ws, row, ncols, fill, font)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)


def month_header(ws, row, first_label, months, n):
    ws.cell(row, 1, first_label)
    for k, d in enumerate(months):
        ws.cell(row, 2 + k, d).number_format = "mmm-yy"
    ws.cell(row, 2 + n, f"{n}-Month Total")
    fill_row(ws, row, n + 2, NAVY, F_HDR, border=Border())
    ws.row_dimensions[row].height = 23.25


def set_widths(ws, n):
    ws.column_dimensions["A"].width = 44
    for c in range(2, n + 3):
        ws.column_dimensions[L(c)].width = 17.5


# ---- store tab --------------------------------------------------------------
def write_store(ws, s):
    n = len(s["months"]); ncols = n + 2
    ws["A1"], ws["A2"] = s["store"], s["address"]
    ws["A1"].font, ws["A2"].font = Font(name="Calibri", size=14), Font(name="Calibri", size=11, color="595959")
    ws["A1"].alignment = ws["A2"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 18.75
    month_header(ws, HDR_ROW, "Line Item", s["months"], n)

    row, prev, rowmap = HDR_ROW, None, {}
    for it in s["items"]:
        k = it["kind"]
        if blank_before(prev, k):
            row += 1
        row += 1
        label, vals = it["label"], it["vals"]
        if k == "header":
            full_row_header(ws, row, label, ncols, YELLOW, F_BOLD)
        elif k == "subheader":
            full_row_header(ws, row, label, ncols, None, F_BOLD)
        elif k == "vheader" and all(not v for v in vals):   # e.g. Net Vending Income = 0 -> plain header
            full_row_header(ws, row, label, ncols, YELLOW, F_BOLD)
        else:
            style = {"total": (BLUE, F_TOTAL), "net": (BLUE, F_NET),
                     "vheader": (YELLOW, F_BOLD), "data": (None, F_DATA)}[k]
            ws.cell(row, 1, label)
            for j, v in enumerate(vals):
                if v is not None:
                    ws.cell(row, 2 + j, v).number_format = NUM_FMT
            fill_row(ws, row, ncols, *style)
            if k == "net":
                ws.row_dimensions[row].height = 20
        if k in ("total", "net"):
            rowmap[it["key"]] = row
        prev = k
    ws.freeze_panes = f"B{HDR_ROW + 1}"
    set_widths(ws, n)
    s["rowmap"] = rowmap

    # sanity check: months must add up to the report's own total column
    for it in s["items"]:
        v = it["vals"]
        if it["kind"] != "header" and all(x is not None for x in v) and abs(sum(v[:-1]) - v[-1]) > 0.05:
            print(f"  ! {s['store']}: '{it['label']}' months add to {sum(v[:-1]):,.2f} "
                  f"but the file's own total says {v[-1]:,.2f} (kept as in source)")


# ---- summary tab ---------------------------------------------------------------
def write_summary(ws, stores):
    months = stores[0]["months"]; n = len(months); ncols = n + 2
    ws["A1"] = "P&L SUMMARY - ALL STORES"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True)
    ws.row_dimensions[1].height = 18.75

    # every vertical (TOTAL line + NET INCOME), in report order, union over all stores
    keys, names = [], {}
    for s in stores:
        last = -1
        for it in s["items"]:
            if it["kind"] in ("total", "net"):
                names.setdefault(it["key"], it["label"])
                if it["key"] in keys:
                    last = keys.index(it["key"])
                else:
                    last += 1
                    keys.insert(last, it["key"])

    def all_zero(key):  # a vertical that is 0 in every store (e.g. FUEL INCOME) is just noise
        return all(not any(it["vals"]) for s in stores for it in s["items"] if it["key"] == key)
    keys = [k for k in keys if not all_zero(k)]

    row = 3
    for key in keys:
        full_row_header(ws, row, names[key], ncols, YELLOW, F_BOLD)
        row += 1
        month_header(ws, row, "Store", months, n)
        first = row + 1
        for s in stores:
            row += 1
            ws.cell(row, 1, s["store"])
            src = s["rowmap"].get(key)
            for k in range(n + 1):
                ws.cell(row, 2 + k, f"='{s['tab']}'!{L(2 + k)}{src}" if src else "n/a").number_format = NUM_FMT
            fill_row(ws, row, ncols, None, F_DATA)
        row += 1
        ws.cell(row, 1, "ALL STORES")
        for k in range(n + 1):
            ws.cell(row, 2 + k, f"=SUM({L(2 + k)}{first}:{L(2 + k)}{row - 1})").number_format = NUM_FMT
        fill_row(ws, row, ncols, BLUE, F_TOTAL)
        row += 3
    set_widths(ws, n)
    ws.column_dimensions["A"].width = 44
    ws.sheet_view.showGridLines = False


def sheet_name(store, used):
    name = re.sub(r"[\[\]\*\?/\\:]", "", store.replace(" - ", " "))
    name = re.sub(r"\s*\(.*?\)", "", name).strip()[:31]
    base, i = name, 2
    while name in used:
        name = f"{base[:28]}_{i}"; i += 1
    used.add(name)
    return name


def build(in_paths, out_path):
    if isinstance(in_paths, str):
        in_paths = [in_paths]
    cand = []
    for p in in_paths:
        cand += glob.glob(os.path.join(p, "*.xlsx")) if os.path.isdir(p) else [p]
    files = sorted(set(f for f in cand
                   if not os.path.basename(f).startswith("~")
                   and os.path.abspath(f) != os.path.abspath(out_path)
                   and "Combined_PnL" not in os.path.basename(f)))
    print(f"{len(files)} files mili")
    stores = sorted((load_store(f) for f in files), key=lambda s: s["store"])
    wb = openpyxl.Workbook()
    summary = wb.active; summary.title = "Summary"
    used = {"Summary"}
    for s in stores:
        ws = wb.create_sheet(sheet_name(s["store"], used))
        s["tab"] = ws.title
        write_store(ws, s)
    write_summary(summary, stores)
    wb.save(out_path)
    print(f"Saved {out_path}  ({len(stores)} stores: {', '.join(s['tab'] for s in stores)})")


if __name__ == "__main__":
    # command-line args (optional) override INPUT_PATHS / OUTPUT_FILE set at the top
    in_paths = [sys.argv[1]] if len(sys.argv) > 1 else INPUT_PATHS
    first = in_paths[0]
    out = sys.argv[2] if len(sys.argv) > 2 else (
        OUTPUT_FILE or os.path.join(first if os.path.isdir(first) else os.path.dirname(first) or ".",
                                    "Combined_PnL_by_Store.xlsx"))
    bad = [p for p in in_paths if not os.path.exists(p)]
    if bad:
        sys.exit("Path nahi mila: " + ", ".join(bad) + "\nScript ke top par INPUT_PATHS check karein.")
    # OUTPUT_FILE agar folder hai (ya .xlsx nahi likha) to khud theek kar do
    if os.path.isdir(out):
        out = os.path.join(out, "Combined_PnL_by_Store.xlsx")
    elif not out.lower().endswith(".xlsx"):
        out += ".xlsx"
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    build(in_paths, out)
