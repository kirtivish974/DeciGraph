
import json
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(".")
DATA = ROOT / "data/raw"

# Load project metadata
with open("SOURCE_MANIFEST.json", encoding="utf-8") as f:
    manifest = json.load(f)

with open("evaluation/decision_scenarios.json", encoding="utf-8") as f:
    scenarios = json.load(f)

with open("evaluation/ground_truth_anomalies.json", encoding="utf-8") as f:
    anomalies = json.load(f)

print("\n========== FILE EXISTENCE AUDIT ==========")

source_paths = []

for table in manifest["csv_files"].values():
    source_paths.append(table["path"])

source_paths.append(manifest["excel_workbook"]["path"])
source_paths.extend(manifest["markdown_documents"])
source_paths.extend(manifest["pdf_documents"])

missing_files = []

for path in source_paths:
    if Path(path).exists():
        print("[PASS]", path)
    else:
        print("[FAIL]", path)
        missing_files.append(path)

print("\n========== CSV ROW COUNT AUDIT ==========")

for table, details in manifest["csv_files"].items():
    df = pd.read_csv(details["path"])

    expected = details["rows"]
    actual = len(df)

    if expected == actual:
        print(f"[PASS] {table}: {actual} rows")
    else:
        print(
            f"[FAIL] {table}: expected {expected}, got {actual}"
        )

print("\n========== EXCEL SHEET AUDIT ==========")

excel_path = manifest["excel_workbook"]["path"]

if Path(excel_path).exists():
    workbook = load_workbook(excel_path, read_only=True)
    actual_sheets = set(workbook.sheetnames)
    expected_sheets = set(manifest["excel_workbook"]["sheets"])

    print("Sheets found:", workbook.sheetnames)

    missing_sheets = expected_sheets - actual_sheets

    if not missing_sheets:
        print("[PASS] All expected Excel sheets exist")
    else:
        print("[FAIL] Missing sheets:", missing_sheets)

    workbook.close()

print("\n========== EVIDENCE ID AUDIT ==========")

# Collect IDs from all CSV records
available_ids = set()

for details in manifest["csv_files"].values():
    df = pd.read_csv(details["path"], dtype=str)

    for column in df.columns:
        available_ids.update(df[column].dropna().str.strip())

# Add registered document IDs
available_ids.update(manifest["document_ids"])

missing_evidence = []

# Check scenario evidence
for scenario in scenarios:
    for evidence_id in scenario["key_evidence_ids"]:
        if evidence_id not in available_ids:
            missing_evidence.append(
                (scenario["scenario_id"], evidence_id)
            )

# Check anomaly evidence
for anomaly in anomalies:
    ids = [
        anomaly["source_record_id"],
        *anomaly["related_record_ids"]
    ]

    for evidence_id in ids:
        if evidence_id not in available_ids:
            missing_evidence.append(
                (anomaly["anomaly_id"], evidence_id)
            )

if not missing_evidence:
    print("[PASS] All evaluation evidence IDs are registered")
else:
    for owner, evidence_id in missing_evidence:
        print(f"[FAIL] {owner}: Missing ID {evidence_id}")

print("\n========== EVALUATION SUMMARY ==========")

print("Decision scenarios:", len(scenarios))
print("Ground truth anomalies:", len(anomalies))
print("Registered document IDs:", len(manifest["document_ids"]))
print("Missing files:", len(missing_files))
print("Missing evidence IDs:", len(missing_evidence))

print("\nEvidence audit completed.")
