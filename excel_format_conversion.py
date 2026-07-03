"""
Convert a "wide" Excel sheet (one column per PlanDate, one row per Category)
into a "long" Excel sheet with columns: UploadType, Category, PlanDate, Quantity.

Usage:
    python excel_format_conversion.py <input.xlsx> <output.xlsx> [--sheet SHEET_NAME] [--upload-type ADD]
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


def convert(input_path: str, sheet_name=0, upload_type: str = "ADD") -> pd.DataFrame:
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

    long_df = raw.melt(
        id_vars=[category_col],
        value_vars=date_cols,
        var_name="PlanDate",
        value_name="Quantity",
    )
    long_df = long_df.rename(columns={category_col: "Category"})
    long_df["PlanDate"] = pd.to_datetime(long_df["PlanDate"]).dt.date
    long_df.insert(0, "UploadType", upload_type)

    long_df = long_df[["UploadType", "Category", "PlanDate", "Quantity"]]
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
    args = parser.parse_args()

    sheet = args.sheet
    try:
        sheet = int(sheet)
    except (TypeError, ValueError):
        pass

    df = convert(args.input, sheet_name=sheet, upload_type=args.upload_type)
    save(df, args.output)
    print(f"Converted {len(df)} rows -> {args.output}")


if __name__ == "__main__":
    sys.exit(main())
