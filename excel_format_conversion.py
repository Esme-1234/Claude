"""
Convert a "wide" Excel sheet (one column per PlanDate, one row per Category)
into a "long" Excel sheet with columns: UploadType, Category, PlanDate, Quantity.

By default the date range is extended by one extra trailing day, carrying
forward the last date column's quantity (e.g. last date 7/3 with qty 80 ->
also emit 7/4 with qty 80). Use --extend-days to change how many extra days
are added, or 0 to disable.

Rows with a blank or zero Quantity are dropped from the output.

If --reference is given, it points to a two-column (Category, quantity)
sheet holding the already-known quantities for the earliest date in the
output. For that earliest date's rows only: if the reference quantity
differs from the computed Quantity and the reference quantity is not 0,
UploadType is changed from ADD to UPDATE for that row (categories missing
from the reference, or where the reference is 0, are left as ADD).
Conversely, any reference Category that isn't present at all among the
earliest date's output rows gets a new row appended with UploadType=DELETE
and Quantity set to that reference quantity.

Usage:
    python excel_format_conversion.py <input.xlsx> <output.xlsx> [--sheet SHEET_NAME] [--upload-type ADD] [--extend-days 1] [--reference reference.xlsx]
"""

import argparse
import datetime
import sys

import pandas as pd


def _is_date_header(value) -> bool:
    if isinstance(value, (pd.Timestamp, datetime.date, datetime.datetime)):
        return True
    if isinstance(value, str):
        try:
            pd.to_datetime(value)
            return True
        except (ValueError, TypeError):
            return False
    return False


def convert(input_path: str, sheet_name=0, upload_type: str = "ADD", extend_days: int = 1) -> pd.DataFrame:
    raw = pd.read_excel(input_path, sheet_name=sheet_name)

    category_col = next(
        (col for col in raw.columns if str(col).strip().lower() == "category"),
        None,
    )
    if category_col is None:
        raise ValueError(
            f"Could not find a 'Category' column in {input_path}. "
            f"Columns found: {list(raw.columns)}"
        )

    # Only take the contiguous run of real date-typed headers right after
    # Category. Sheets can have unrelated trailing columns (totals, a second
    # "actuals" block, etc.) whose headers are not genuine date values, or
    # are duplicate dates pandas has renamed to strings (e.g. "...0.1") -
    # both are excluded by requiring an actual date/datetime type.
    cols = list(raw.columns)
    start = cols.index(category_col) + 1
    date_cols = []
    for col in cols[start:]:
        if not _is_date_header(col):
            break
        date_cols.append(col)
    if not date_cols:
        raise ValueError("No date columns found immediately after the 'Category' column.")

    # Restrict to the contiguous block of real category rows starting right
    # below the header. Sheets often have unrelated notes/totals further
    # down that happen to reuse the same columns - stop at the first row
    # whose Category value isn't a fresh, non-empty category name.
    seen = set()
    end = len(raw)
    for i, value in enumerate(raw[category_col]):
        if not isinstance(value, str) or not value.strip() or value in seen:
            end = i
            break
        seen.add(value)
    raw = raw.iloc[:end]

    # Extend the date range by carrying the last date column's values forward
    # onto extra_days new trailing dates (e.g. last date 7/3 -> also emit
    # 7/4 with the same quantity as 7/3), before melting to long format.
    last_date_col = date_cols[-1]
    last_date = pd.to_datetime(last_date_col)
    extra_cols = []
    for i in range(1, extend_days + 1):
        new_date_col = last_date + pd.Timedelta(days=i)
        raw[new_date_col] = raw[last_date_col]
        extra_cols.append(new_date_col)
    all_date_cols = date_cols + extra_cols

    long_df = raw.melt(
        id_vars=[category_col],
        value_vars=all_date_cols,
        var_name="PlanDate",
        value_name="Quantity",
    )
    long_df = long_df.rename(columns={category_col: "Category"})
    long_df["PlanDate"] = pd.to_datetime(long_df["PlanDate"]).dt.date
    long_df.insert(0, "UploadType", upload_type)

    long_df = long_df[["UploadType", "Category", "PlanDate", "Quantity"]]
    long_df = long_df[long_df["Quantity"].notna() & (long_df["Quantity"] != 0)]
    return long_df


def apply_update_flag(long_df: pd.DataFrame, reference_path: str) -> pd.DataFrame:
    ref = pd.read_excel(reference_path, sheet_name=0)
    category_col, value_col = ref.columns[0], ref.columns[1]
    reference = dict(zip(ref[category_col], ref[value_col]))

    earliest_date = long_df["PlanDate"].min()
    is_earliest = long_df["PlanDate"] == earliest_date

    def decide(upload_type, category, quantity, is_earliest_row):
        if not is_earliest_row:
            return upload_type
        ref_qty = reference.get(category)
        if ref_qty is not None and ref_qty != 0 and ref_qty != quantity:
            return "UPDATE"
        return upload_type

    long_df = long_df.copy()
    long_df["UploadType"] = [
        decide(ut, cat, qty, earliest)
        for ut, cat, qty, earliest in zip(
            long_df["UploadType"], long_df["Category"], long_df["Quantity"], is_earliest
        )
    ]

    # Reference categories absent from the earliest date's output rows are
    # no longer part of the plan - flag them for deletion.
    existing_categories = set(long_df.loc[is_earliest, "Category"])
    missing = [cat for cat in reference if cat not in existing_categories]
    if missing:
        delete_rows = pd.DataFrame(
            {
                "UploadType": "DELETE",
                "Category": missing,
                "PlanDate": earliest_date,
                "Quantity": [reference[cat] for cat in missing],
            }
        )
        long_df = pd.concat([long_df, delete_rows], ignore_index=True)

    return long_df


def save(df: pd.DataFrame, output_path: str) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")
        ws = writer.sheets["Sheet1"]
        widths = {"A": 12, "B": 12, "C": 12, "D": 10}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        date_col_letter = "C"
        for row in range(2, ws.max_row + 1):
            ws[f"{date_col_letter}{row}"].number_format = "m/d/yyyy"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", help="Path to the source .xlsx file (wide format)")
    parser.add_argument("output", help="Path to write the converted .xlsx file (long format)")
    parser.add_argument("--sheet", default=0, help="Sheet name or index to read (default: first sheet)")
    parser.add_argument("--upload-type", default="ADD", help="Value to fill the UploadType column (default: ADD)")
    parser.add_argument(
        "--extend-days",
        type=int,
        default=1,
        help="Number of extra trailing dates to add after the last date column, "
        "each carrying forward the last date's quantity (default: 1, use 0 to disable)",
    )
    parser.add_argument(
        "--reference",
        default=None,
        help="Path to a (Category, quantity) sheet for the earliest date; rows whose "
        "quantity differs from a non-zero reference value get UploadType=UPDATE",
    )
    args = parser.parse_args()

    sheet = args.sheet
    try:
        sheet = int(sheet)
    except (TypeError, ValueError):
        pass

    df = convert(args.input, sheet_name=sheet, upload_type=args.upload_type, extend_days=args.extend_days)
    if args.reference:
        df = apply_update_flag(df, args.reference)
    save(df, args.output)
    print(f"Converted {len(df)} rows -> {args.output}")


if __name__ == "__main__":
    sys.exit(main())
