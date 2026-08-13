"""
Match anode-shortage part numbers against available anode inventory.

Pipeline
--------
1. "Loading - Practice.xlsx" / sheet "KEMET_GOP_ATO_ANALYSIS_DAILY"
   - column D  (PARENT_ITEM)     -> part number
   - column F  (Actual Start Qty)-> quantity needed for that demand row
   - column E  (SUGG_START_DATE) -> demand date, filtered to a configurable window
   - column AE (Kpcs/Cycles)     -> filtered against a configurable threshold
   - column C  (Anode)           -> the anode/component code for that part number
   Rows that pass the date-window and AE-threshold filters are summed per part
   number to get RequiredQty.

2. "PlanningResult -Practice.xlsx" / sheet "Sheet1"
   - column B (partNumber) / column F (quantity)
   Summed per part number to get PlannedQty.

   ShortageQty = RequiredQty - PlannedQty (only kept when > 0, i.e. planning
   has not covered the requirement).

3. For every shortage part number, look up its anode/component code (from
   step 1) in "Anode 3-4-7.31.xlsx" / sheet "details", matching column C
   (ITEM_NAME). A lot is usable if either:
     a) column E (PlanDate) is blank AND column H (SUBINVENTORY) is
        "ANODE-KO" or "ANODE-INSP", or
     b) column H is "Intransit" AND the date encoded in column G (WAYBILL,
        e.g. "MES-260724B384" -> positions 5-10 = YYMMDD = 2026-07-24) is
        before a configurable cutoff date.
   LOT_NUMBERs that appear more than once anywhere in column A (the same rows
   Excel's "Highlight Duplicate Values" conditional format flags in pink) are
   excluded entirely, since a repeated lot number in the source data is
   unreliable. Usable lots' column F (QTY) are summed to AvailableAnodeQty.

4. Output: every shortage part number that has at least one usable anode lot
   (RequiredQty, PlannedQty, ShortageQty, AvailableAnodeQty, ...), plus a
   detail sheet listing the individual anode lots that matched.

All the filters mentioned above (demand date window, AE threshold, anode
cutoff date) are exposed as CLI options so they can be re-tuned per run.

Usage
-----
    python anode_shortage_matching.py \
        --loading "Loading - Practice.xlsx" \
        --planning-result "PlanningResult -Practice.xlsx" \
        --anode "Anode 3-4-7.31.xlsx" \
        --output shortage_with_anode_report.xlsx \
        --date-start 2026-08-06 --months 1 \
        --kpcs-threshold 6 \
        --anode-cutoff-date 2026-07-27
"""

import argparse
import re
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta

import openpyxl
from openpyxl.styles import PatternFill

LOADING_SHEET = "KEMET_GOP_ATO_ANALYSIS_DAILY"
PLANNING_SHEET = "Sheet1"
ANODE_SHEET = "details"

# Columns are located by header text, not fixed letters, because source files are
# re-uploaded from time to time with columns inserted/reordered.
LOADING_COL = {"part_number": "PARENT_ITEM", "qty": "Actual Start Qty", "date": "SUGG_START_DATE",
               "anode_code": "Anode", "kpcs": "Kpcs/Cycles", "category": "Anode Category",
               "total_cycle": "Total Cycle", "elect_type": "Elect Type"}
PLANNING_COL = {"part_number": "partNumber", "qty": "quantity", "lot_id": "anodeLotID"}
ANODE_COL = {"anode_code": "ITEM_NAME", "powder_type": "PowderType", "plan_date": "PlanDate", "qty": "QTY",
             "waybill": "WAYBILL", "subinventory": "SUBINVENTORY", "lot": "LOT_NUMBER"}

WAYBILL_DATE_RE = re.compile(r"^MES-(\d{2})(\d{2})(\d{2})")
USABLE_DIRECT_STATUSES = {"ANODE-KO", "ANODE-INSP"}
INTRANSIT_STATUS = "Intransit"


def _header_indices(ws, wanted):
    """Map {key: 0-based column index} for the header names in `wanted` (a dict of key->header text)."""
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True))
    name_to_idx = {name: i for i, name in enumerate(header_row) if name is not None}
    result = {}
    missing = []
    for key, header_name in wanted.items():
        if header_name in name_to_idx:
            result[key] = name_to_idx[header_name]
        else:
            missing.append(header_name)
    if missing:
        raise ValueError(f"Column header(s) not found in sheet '{ws.title}': {missing}. "
                          f"Available headers: {sorted(n for n in name_to_idx)}")
    return result


def _to_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and value.strip():
        for fmt in ("%Y-%m-%d", "%m/%d/%Y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None


def _add_months(d, months):
    month = d.month - 1 + months
    year = d.year + month // 12
    month = month % 12 + 1
    day = min(d.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
                       31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
    return date(year, month, day)


def parse_waybill_date(waybill):
    if not waybill:
        return None
    m = WAYBILL_DATE_RE.match(str(waybill).strip())
    if not m:
        return None
    yy, mm, dd = (int(g) for g in m.groups())
    try:
        return date(2000 + yy, mm, dd)
    except ValueError:
        return None


def build_cumulative_demand(demand_rows):
    """[(date, qty), ...] (any order, duplicate dates allowed) -> [(date, cumulative_qty), ...] sorted ascending.

    All rows sharing the same SUGG_START_DATE are summed into one step: a date
    only counts as "covered" once every row on it is fully satisfied.
    """
    by_date = defaultdict(float)
    for d, qty in demand_rows:
        by_date[d] += qty or 0
    curve = []
    running = 0.0
    for d in sorted(by_date):
        running += by_date[d]
        curve.append((d, running))
    return curve


def covered_through_date(cumulative_curve, supply_qty):
    """Latest date whose full-day cumulative demand is <= supply_qty (i.e. fully paid off),
    or None if supply_qty doesn't even cover the earliest date's full-day total. Meant to be
    evaluated with the running supply as of *after* a given lot is added, so the lot is
    labeled with the furthest date its addition genuinely finishes covering."""
    result = None
    for d, cum_qty in cumulative_curve:
        if cum_qty <= supply_qty:
            result = d
        else:
            break
    return result


def next_uncovered_date(cumulative_curve, supply_qty):
    """Earliest date whose full-day cumulative demand still exceeds supply_qty, or None if
    supply_qty already covers the entire backlog."""
    for d, cum_qty in cumulative_curve:
        if cum_qty > supply_qty:
            return d
    return None


def get_demand(loading_path, date_start, date_end, kpcs_threshold):
    """Return (required_qty_by_part, anode_code_by_part)."""
    wb = openpyxl.load_workbook(loading_path, read_only=True, data_only=True)
    ws = wb[LOADING_SHEET]

    idx = _header_indices(ws, LOADING_COL)
    part_i = idx["part_number"]
    qty_i = idx["qty"]
    date_i = idx["date"]
    anode_i = idx["anode_code"]
    kpcs_i = idx["kpcs"]
    category_i = idx["category"]
    total_cycle_i = idx["total_cycle"]
    elect_type_i = idx["elect_type"]

    required = defaultdict(float)
    anode_code_by_part = {}
    category_by_part = {}
    total_cycle_by_part = {}
    elect_type_by_part = {}
    demand_rows_by_part = defaultdict(list)
    full_demand_rows_by_part = defaultdict(list)

    for row in ws.iter_rows(min_row=2, values_only=True):
        part = row[part_i]
        if not part:
            continue
        if anode_code_by_part.get(part) is None and row[anode_i]:
            anode_code_by_part[part] = row[anode_i]
        if category_by_part.get(part) is None and row[category_i]:
            category_by_part[part] = row[category_i]
        if total_cycle_by_part.get(part) is None and isinstance(row[total_cycle_i], (int, float)):
            total_cycle_by_part[part] = row[total_cycle_i]
        if elect_type_by_part.get(part) is None and row[elect_type_i]:
            elect_type_by_part[part] = row[elect_type_i]

        row_date = _to_date(row[date_i])
        kpcs = row[kpcs_i]
        if not isinstance(kpcs, (int, float)):
            kpcs = None
        if row_date is None or kpcs is None or not (kpcs > kpcs_threshold):
            continue

        qty = row[qty_i] or 0
        # Full order backlog for this part (any date), used to see how far the
        # covered date reaches beyond the near-term demand window.
        full_demand_rows_by_part[part].append((row_date, qty))

        if date_start <= row_date <= date_end:
            required[part] += qty
            demand_rows_by_part[part].append((row_date, qty))

    wb.close()
    return (required, anode_code_by_part, category_by_part, total_cycle_by_part, elect_type_by_part,
            demand_rows_by_part, full_demand_rows_by_part)


def get_planned(planning_path):
    wb = openpyxl.load_workbook(planning_path, read_only=True, data_only=True)
    ws = wb[PLANNING_SHEET]

    idx = _header_indices(ws, PLANNING_COL)
    part_i = idx["part_number"]
    qty_i = idx["qty"]
    lot_id_i = idx["lot_id"]

    planned = defaultdict(float)
    used_lot_ids = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        part = row[part_i]
        if part:
            planned[part] += row[qty_i] or 0
        if row[lot_id_i]:
            used_lot_ids.add(row[lot_id_i])

    wb.close()
    return planned, used_lot_ids


def is_usable_lot(plan_date, subinventory, waybill, anode_cutoff_date, cutoff_inclusive=False):
    sub = (subinventory or "").strip()
    if plan_date is None and sub in USABLE_DIRECT_STATUSES:
        return True
    if sub == INTRANSIT_STATUS:
        wb_date = parse_waybill_date(waybill)
        if wb_date is not None and (wb_date <= anode_cutoff_date if cutoff_inclusive
                                     else wb_date < anode_cutoff_date):
            return True
    return False


def get_available_anode(anode_path, anode_codes_needed, anode_cutoff_date, used_lot_ids=frozenset(),
                         cutoff_inclusive=False):
    """Return (available_qty_by_code, lot_rows_by_code) for the requested anode codes.

    LOT_NUMBERs that appear more than once anywhere in the sheet's column A (the
    same duplicates Excel's "Highlight Duplicate Values" conditional formatting
    flags in pink on the source file) are treated as unreliable and excluded
    entirely, regardless of anode code or usability status. LOT_NUMBERs already
    listed as an anodeLotID in PlanningResult (i.e. already consumed by planning
    and counted in PlannedQty) are excluded too, so the same physical lot is
    never counted as both "planned" and "still available".
    """
    wb = openpyxl.load_workbook(anode_path, read_only=True, data_only=True)
    ws = wb[ANODE_SHEET]

    idx = _header_indices(ws, ANODE_COL)
    code_i = idx["anode_code"]
    powder_i = idx["powder_type"]
    plan_date_i = idx["plan_date"]
    qty_i = idx["qty"]
    waybill_i = idx["waybill"]
    sub_i = idx["subinventory"]
    lot_i = idx["lot"]

    rows = []
    lot_counts = defaultdict(int)
    for row in ws.iter_rows(min_row=2, values_only=True):
        lot = row[lot_i]
        if lot:
            lot_counts[lot] += 1
        rows.append(row)
    wb.close()

    available = defaultdict(float)
    lots = defaultdict(list)

    for row in rows:
        code = row[code_i]
        lot = row[lot_i]
        if code not in anode_codes_needed:
            continue
        if not lot or lot_counts[lot] > 1:
            continue  # duplicated LOT_NUMBER in the source sheet -> excluded
        if lot in used_lot_ids:
            continue  # already consumed by a PlanningResult row -> excluded
        plan_date = _to_date(row[plan_date_i])
        subinventory = row[sub_i]
        waybill = row[waybill_i]
        qty = row[qty_i] or 0
        powder_type = row[powder_i]

        if not is_usable_lot(plan_date, subinventory, waybill, anode_cutoff_date, cutoff_inclusive):
            continue

        available[code] += qty
        lots[code].append((lot, code, powder_type, plan_date, qty, waybill,
                            subinventory, parse_waybill_date(waybill)))

    return available, lots


def build_report(loading_path, planning_path, anode_path, date_start, date_end,
                  kpcs_threshold, anode_cutoff_date, cutoff_inclusive=False):
    (required, anode_code_by_part, category_by_part, total_cycle_by_part, elect_type_by_part,
     demand_rows_by_part, full_demand_rows_by_part) = get_demand(
        loading_path, date_start, date_end, kpcs_threshold)
    planned, used_lot_ids = get_planned(planning_path)

    shortages = {}
    for part, req_qty in required.items():
        planned_qty = planned.get(part, 0)
        shortage_qty = req_qty - planned_qty
        if shortage_qty > 0:
            shortages[part] = (req_qty, planned_qty, shortage_qty)

    anode_codes_needed = {anode_code_by_part[p] for p in shortages if anode_code_by_part.get(p)}
    available, lots = get_available_anode(anode_path, anode_codes_needed, anode_cutoff_date, used_lot_ids,
                                           cutoff_inclusive)

    summary_rows = []
    detail_rows = []
    for part, (req_qty, planned_qty, shortage_qty) in shortages.items():
        anode_code = anode_code_by_part.get(part)
        available_qty = available.get(anode_code, 0) if anode_code else 0
        if not anode_code or available_qty <= 0:
            continue  # no anode available for this shortage -> excluded per requirement

        summary_rows.append({
            "part_number": part,
            "anode_code": anode_code,
            "required_qty": req_qty,
            "planned_qty": planned_qty,
            "shortage_qty": shortage_qty,
            "available_anode_qty": available_qty,
            "available_lot_count": len(lots.get(anode_code, [])),
            "anode_covers_shortage": available_qty >= shortage_qty,
        })

        cumulative_curve = build_cumulative_demand(full_demand_rows_by_part.get(part, []))
        part_lots = sorted(lots.get(anode_code, []), key=lambda entry: entry[7] or date.min)
        running_supply = planned_qty
        for lot_number, code, powder_type, plan_date, qty, waybill, subinventory, wb_date in part_lots:
            running_supply += qty
            covered_through = covered_through_date(cumulative_curve, running_supply)
            if covered_through is not None:
                covered_display = covered_through
            else:
                # Doesn't even cover the earliest date's full-day demand yet -> name
                # that still-unmet date on every such row, not just the last one.
                still_short_date = next_uncovered_date(cumulative_curve, running_supply)
                covered_display = f"{still_short_date}不够" if still_short_date else "不够"
            detail_rows.append({
                "covered_through_date": covered_display,
                "category": category_by_part.get(part),
                "total_cycle": total_cycle_by_part.get(part),
                "elect_type": elect_type_by_part.get(part),
                "part_number": part,
                "anode_code": code,
                "lot_number": lot_number,
                "powder_type": powder_type,
                "plan_date": plan_date,
                "qty": qty,
                "waybill": waybill,
                "subinventory": subinventory,
                "waybill_date": wb_date,
            })

    summary_rows.sort(key=lambda r: r["shortage_qty"], reverse=True)
    return summary_rows, detail_rows


def save(summary_rows, detail_rows, params, output_path):
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Shortage_With_Anode"
    header = ["No.", "PartNumber", "AnodeCode", "RequiredQty", "PlannedQty",
              "ShortageQty", "AvailableAnodeQty", "AvailableLotCount", "AnodeCoversShortage"]
    ws.append(header)
    for i, r in enumerate(summary_rows, 1):
        ws.append([i, r["part_number"], r["anode_code"], r["required_qty"], r["planned_qty"],
                   r["shortage_qty"], r["available_anode_qty"], r["available_lot_count"],
                   "Y" if r["anode_covers_shortage"] else "N"])
    widths = (6, 26, 22, 14, 14, 14, 18, 16, 18)
    for col, width in zip("ABCDEFGHI", widths):
        ws.column_dimensions[col].width = width
    for row in range(2, ws.max_row + 1):
        for col in "DEFG":
            ws[f"{col}{row}"].number_format = "#,##0"

    ws_detail = wb.create_sheet("Anode_Lot_Detail")
    ws_detail.append(["Category", "TotalCycle", "ElectType", "CoveredThroughDate", "LotNumber", "PartNumber",
                       "AnodeCode", "PowderType", "PlanDate", "QTY", "WAYBILL", "SUBINVENTORY", "WaybillDate",
                       "PartNumber1", "Anode1", "LotNumber1", "QTY1"])
    for r in detail_rows:
        ws_detail.append([r["category"], r["total_cycle"], r["elect_type"], r["covered_through_date"],
                           r["lot_number"], r["part_number"], r["anode_code"], r["powder_type"], r["plan_date"],
                           r["qty"], r["waybill"], r["subinventory"], r["waybill_date"],
                           r["part_number"], r["anode_code"], r["lot_number"], r["qty"]])
    widths = (14, 12, 12, 18, 16, 26, 22, 12, 12, 12, 16, 16, 14, 26, 22, 16, 12)
    for col, width in zip("ABCDEFGHIJKLMNOPQ", widths):
        ws_detail.column_dimensions[col].width = width
    for row in range(2, ws_detail.max_row + 1):
        ws_detail[f"D{row}"].number_format = "yyyy-mm-dd"
        ws_detail[f"I{row}"].number_format = "yyyy-mm-dd"
        ws_detail[f"M{row}"].number_format = "yyyy-mm-dd"

    lot_counts = defaultdict(int)
    for r in detail_rows:
        lot_counts[r["lot_number"]] += 1
    duplicate_fill = PatternFill(start_color="FFD9A0", end_color="FFD9A0", fill_type="solid")
    for row_idx, r in enumerate(detail_rows, start=2):
        if lot_counts[r["lot_number"]] > 1:
            ws_detail[f"E{row_idx}"].fill = duplicate_fill

    mirror_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    for row_idx in range(1, ws_detail.max_row + 1):
        for col in "NOPQ":
            ws_detail[f"{col}{row_idx}"].fill = mirror_fill

    ws_params = wb.create_sheet("Parameters")
    ws_params.append(["Parameter", "Value"])
    for k, v in params.items():
        ws_params.append([k, str(v)])
    ws_params.column_dimensions["A"].width = 24
    ws_params.column_dimensions["B"].width = 30

    wb.save(output_path)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--loading", default="Loading - Practice.xlsx",
                         help="Path to the Loading .xlsx file (KEMET_GOP_ATO_ANALYSIS_DAILY sheet)")
    parser.add_argument("--planning-result", default="PlanningResult -Practice.xlsx",
                         help="Path to the PlanningResult .xlsx file")
    parser.add_argument("--anode", default="Anode 3-4-7.31.xlsx",
                         help="Path to the Anode inventory .xlsx file (details sheet)")
    parser.add_argument("--output", default="shortage_with_anode_report.xlsx",
                         help="Path to write the output report .xlsx file")

    parser.add_argument("--date-start", default=None,
                         help="Start of the demand date window (YYYY-MM-DD). Default: today.")
    parser.add_argument("--date-end", default=None,
                         help="End of the demand date window (YYYY-MM-DD). "
                              "Default: --date-start plus --months.")
    parser.add_argument("--months", type=int, default=1,
                         help="Width of the demand date window in months when --date-end is not "
                              "given (default: 1, i.e. 'the next month').")

    parser.add_argument("--kpcs-threshold", type=float, default=6.0,
                         help="AE column (Kpcs/Cycles) must be strictly greater than this value "
                              "for a demand row to count (default: 6).")

    parser.add_argument("--anode-cutoff-date", default="2026-07-27",
                         help="An 'Intransit' anode lot is usable only if its WAYBILL date "
                              "(column G, positions 5-10 = YYMMDD) is before this date (YYYY-MM-DD), "
                              "strictly before unless --anode-cutoff-inclusive is set. Default: 2026-07-27.")
    parser.add_argument("--anode-cutoff-inclusive", action="store_true",
                         help="Include the cutoff date itself (WAYBILL date <= cutoff) instead of the "
                              "default strictly-before comparison.")

    args = parser.parse_args()

    date_start = _to_date(args.date_start) if args.date_start else date.today()
    if args.date_end:
        date_end = _to_date(args.date_end)
    else:
        date_end = _add_months(date_start, args.months)

    anode_cutoff_date = _to_date(args.anode_cutoff_date)

    summary_rows, detail_rows = build_report(
        args.loading, args.planning_result, args.anode,
        date_start, date_end, args.kpcs_threshold, anode_cutoff_date,
        args.anode_cutoff_inclusive,
    )

    params = {
        "loading_file": args.loading,
        "planning_result_file": args.planning_result,
        "anode_file": args.anode,
        "demand_date_start": date_start,
        "demand_date_end": date_end,
        "kpcs_threshold (AE column, strictly greater than)": args.kpcs_threshold,
        "anode_cutoff_date (Intransit WAYBILL date, {} cutoff)".format(
            "on or before" if args.anode_cutoff_inclusive else "strictly before"): anode_cutoff_date,
    }
    save(summary_rows, detail_rows, params, args.output)

    print(f"{len(summary_rows)} part numbers are short on planned quantity AND have usable anode "
          f"-> {args.output}")


if __name__ == "__main__":
    sys.exit(main())
