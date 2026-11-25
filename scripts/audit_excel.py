import pandas as pd
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

tables = [
    "suppliers",
    "products",
    "purchase_orders",
    "supplier_deliveries",
    "invoices",
    "payments",
    "contracts",
    "quality_reports",
]

excel_path = ROOT / "data" / "processed" / "procurement_data.xlsx"

print("\n========== CSV vs EXCEL AUDIT ==========")

all_passed = True

for table in tables:
    csv_path = ROOT / "data" / "raw" / f"{table}.csv"

    csv_df = pd.read_csv(csv_path)
    excel_df = pd.read_excel(excel_path, sheet_name=table)

    if list(csv_df.columns) != list(excel_df.columns):
        print(f"[FAIL] {table}: Column mismatch")
        all_passed = False
        continue

    if csv_df.shape != excel_df.shape:
        print(f"[FAIL] {table}: Shape mismatch")
        all_passed = False
        continue

    try:
        pd.testing.assert_frame_equal(
            csv_df,
            excel_df,
            check_dtype=False,
            check_exact=False,
            rtol=1e-9,
            atol=1e-9,
        )

        print(f"[PASS] {table}: Data matches")

    except AssertionError as error:
        print(f"[FAIL] {table}: Actual data mismatch")
        print(error)
        all_passed = False

print("\n========== FINAL RESULT ==========")

if all_passed:
    print("[PASS] All 8 CSV tables match their Excel sheets.")
else:
    print("[FAIL] Some tables need investigation.")