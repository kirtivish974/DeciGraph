from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV_DIR = ROOT / "data" / "raw"
EXCEL_PATH = ROOT / "data" / "processed" / "procurement_data.xlsx"

excel = pd.ExcelFile(EXCEL_PATH)

for csv_path in sorted(CSV_DIR.glob("*.csv")):
    table = csv_path.stem

    csv_df = pd.read_csv(csv_path).fillna("").astype(str)
    excel_df = pd.read_excel(
        EXCEL_PATH,
        sheet_name=table
    ).fillna("").astype(str)

    if csv_df.equals(excel_df):
        continue

    print(f"\n{'=' * 55}")
    print(f"TABLE: {table}")
    print(f"{'=' * 55}")

    if list(csv_df.columns) != list(excel_df.columns):
        print("COLUMN DIFFERENCE:")
        print("CSV:", list(csv_df.columns))
        print("Excel:", list(excel_df.columns))
        continue

    if csv_df.shape != excel_df.shape:
        print("SHAPE DIFFERENCE:")
        print("CSV:", csv_df.shape)
        print("Excel:", excel_df.shape)
        continue

    differences = 0

    for row in range(len(csv_df)):
        for col in csv_df.columns:
            csv_value = csv_df.at[row, col]
            excel_value = excel_df.at[row, col]

            if csv_value != excel_value:
                differences += 1

                if differences <= 15:
                    print(
                        f"Row {row + 1}, Column '{col}'\n"
                        f"  CSV:   {repr(csv_value)}\n"
                        f"  Excel: {repr(excel_value)}"
                    )

    print(f"\nTotal differing cells: {differences}")

print("\nDiagnosis completed.")