"""
Convert a "wide" Excel sheet (one column per PlanDate, one row per Category)
into a "long" Excel sheet with columns: UploadType, Category, PlanDate, Quantity.

Usage:
    python excel_format_conversion.py <input.xlsx> <output.xlsx> [--sheet SHEET_NAME] [--upload-type ADD]
"""

import argparse
import sys

import pandas as pd


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

    date_cols = [col for col in raw.columns if col != category_col]
    if not date_cols:
        raise ValueError("No date columns found after the 'Category' column.")

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
