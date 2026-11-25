"""
DecisionGraph Dataset Validator
===============================
Runs 18 structural, arithmetic, referential, document, and anomaly-discovery
checks on the generated dataset, distinguishes expected planted anomalies from
unexpected errors, writes reports/validation_report.json, and exits non-zero
if any critical check fails.
"""

from __future__ import annotations

import json
import re
import sys
from datetime import date
from pathlib import Path

import openpyxl
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent.parent
RAW_DIR = BASE_DIR / "data" / "raw"
PROCESSED_DIR = BASE_DIR / "data" / "processed"
DOCS_DIR = BASE_DIR / "documents"
EVAL_DIR = BASE_DIR / "evaluation"
REPORTS_DIR = BASE_DIR / "reports"

DATASET_START = date(2025, 4, 1)
DATASET_END = date(2026, 3, 31)
ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

EXPECTED_SCHEMAS = {
    "suppliers": (5, "supplier_id", [
        "supplier_id", "supplier_name", "contact_person", "email", "phone",
        "address", "city", "state", "gstin", "registration_date",
        "supplier_category", "risk_rating", "status",
    ]),
    "products": (10, "product_id", [
        "product_id", "product_name", "category", "unit_of_measure",
        "hsn_code", "standard_unit_price", "gst_rate", "description",
    ]),
    "purchase_orders": (25, "po_id", [
        "po_id", "supplier_id", "product_id", "order_date",
        "expected_delivery_date", "quantity", "unit_price", "total_amount",
        "gst_amount", "grand_total", "status", "approved_by",
    ]),
    "supplier_deliveries": (25, "delivery_id", [
        "delivery_id", "po_id", "supplier_id", "expected_delivery_date",
        "actual_delivery_date", "quantity_ordered", "quantity_delivered",
        "delivery_status", "delay_days", "remarks",
    ]),
    "invoices": (25, "invoice_id", [
        "invoice_id", "invoice_number", "po_id", "supplier_id",
        "invoice_date", "taxable_amount", "gst_rate", "gst_amount",
        "total_amount", "due_date", "payment_status", "remarks",
    ]),
    "payments": (25, "payment_id", [
        "payment_id", "invoice_id", "supplier_id", "payment_date",
        "payment_amount", "payment_mode", "reference_number", "status",
    ]),
    "contracts": (5, "contract_id", [
        "contract_id", "supplier_id", "contract_start_date",
        "contract_end_date", "contract_value", "payment_terms",
        "delivery_terms", "penalty_clause", "auto_renewal", "status",
    ]),
    "quality_reports": (15, "qr_id", [
        "qr_id", "supplier_id", "po_id", "product_id", "inspection_date",
        "batch_size", "defects_found", "defect_rate", "inspection_result",
        "inspector", "remarks",
    ]),
}


def run_validation() -> dict:
    checks: list[dict] = []
    expected_anomalies_found: list[str] = []
    unexpected_errors: list[str] = []
    warnings: list[str] = []

    def record(name: str, passed: bool, detail: str, critical: bool = True) -> None:
        checks.append({"check": name, "status": "PASS" if passed else "FAIL", "detail": detail})
        if not passed:
            if critical:
                unexpected_errors.append(f"{name}: {detail}")
            else:
                warnings.append(f"{name}: {detail}")

    # 1. Required file inventory
    required_files = (
        [RAW_DIR / f"{k}.csv" for k in EXPECTED_SCHEMAS]
        + [
            PROCESSED_DIR / "procurement_data.xlsx",
            EVAL_DIR / "decision_scenarios.json",
            EVAL_DIR / "ground_truth_anomalies.json",
            BASE_DIR / "README.md",
            BASE_DIR / "DATA_DICTIONARY.md",
            BASE_DIR / "SOURCE_MANIFEST.json",
        ]
    )
    missing = [str(p.relative_to(BASE_DIR)) for p in required_files if not p.exists()]
    record("1_required_file_inventory", len(missing) == 0, "All required files present" if not missing else f"Missing: {missing}")

    dfs = {k: pd.read_csv(RAW_DIR / f"{k}.csv") for k in EXPECTED_SCHEMAS if (RAW_DIR / f"{k}.csv").exists()}

    # 2. CSV row counts and schemas
    schema_errs = []
    for name, (exp_rows, _, exp_cols) in EXPECTED_SCHEMAS.items():
        df = dfs.get(name)
        if df is None:
            schema_errs.append(f"{name} missing")
            continue
        if len(df) != exp_rows:
            schema_errs.append(f"{name}: expected {exp_rows} rows, got {len(df)}")
        if list(df.columns) != exp_cols:
            schema_errs.append(f"{name}: column mismatch")
    record("2_csv_row_counts_and_schemas", len(schema_errs) == 0, "All 8 CSV schemas and row counts match" if not schema_errs else "; ".join(schema_errs))

    # 3. Duplicate primary IDs
    pk_errs = []
    for name, (_, pk_col, _) in EXPECTED_SCHEMAS.items():
        df = dfs[name]
        if not df[pk_col].is_unique:
            pk_errs.append(f"{name}.{pk_col} has duplicates")
    record("3_duplicate_primary_ids", len(pk_errs) == 0, "All primary keys unique" if not pk_errs else "; ".join(pk_errs))

    # 4. Foreign-key integrity
    fk_errs = []
    sups = set(dfs["suppliers"]["supplier_id"])
    prds = set(dfs["products"]["product_id"])
    pos = set(dfs["purchase_orders"]["po_id"])
    invs = set(dfs["invoices"]["invoice_id"])
    fk_checks = [
        ("contracts", "supplier_id", sups),
        ("purchase_orders", "supplier_id", sups),
        ("purchase_orders", "product_id", prds),
        ("supplier_deliveries", "po_id", pos),
        ("supplier_deliveries", "supplier_id", sups),
        ("invoices", "po_id", pos),
        ("invoices", "supplier_id", sups),
        ("payments", "invoice_id", invs),
        ("payments", "supplier_id", sups),
        ("quality_reports", "supplier_id", sups),
        ("quality_reports", "po_id", pos),
        ("quality_reports", "product_id", prds),
    ]
    for tbl, col, valid in fk_checks:
        bad = set(dfs[tbl][col]) - valid
        if bad:
            fk_errs.append(f"{tbl}.{col} invalid FKs: {bad}")
    record("4_foreign_key_integrity", len(fk_errs) == 0, "All foreign keys resolve" if not fk_errs else "; ".join(fk_errs))

    # 5. Date ranges and date formats
    date_errs = []
    date_cols = {
        "purchase_orders": ["order_date", "expected_delivery_date"],
        "supplier_deliveries": ["expected_delivery_date", "actual_delivery_date"],
        "invoices": ["invoice_date", "due_date"],
        "payments": ["payment_date"],
        "contracts": ["contract_start_date", "contract_end_date"],
        "quality_reports": ["inspection_date"],
    }
    for tbl, cols in date_cols.items():
        for col in cols:
            for val in dfs[tbl][col].astype(str):
                if not ISO_DATE_RE.match(val):
                    date_errs.append(f"{tbl}.{col} bad ISO date {val}")
                    continue
                dt = date.fromisoformat(val)
                # due_date for March invoices can extend 45 days past DATASET_END
                max_dt = date(2026, 5, 15) if col == "due_date" else DATASET_END
                if not (DATASET_START <= dt <= max_dt):
                    date_errs.append(f"{tbl}.{col} out of range: {val}")
    record("5_date_ranges_and_formats", len(date_errs) == 0, "All dates valid ISO and in range" if not date_errs else "; ".join(date_errs))

    # 6. GST arithmetic and invoice totals
    gst_errs = []
    for _, inv in dfs["invoices"].iterrows():
        exp_gst = round(inv["taxable_amount"] * inv["gst_rate"], 2)
        exp_tot = round(inv["taxable_amount"] + inv["gst_amount"], 2)
        if abs(inv["gst_amount"] - exp_gst) >= 0.01 or abs(inv["total_amount"] - exp_tot) >= 0.01:
            gst_errs.append(inv["invoice_id"])
    record("6_gst_arithmetic_and_totals", len(gst_errs) == 0, "All 25 invoices have exact GST arithmetic" if not gst_errs else f"Errors: {gst_errs}")

    # 7. Invoice-to-PO discrepancies
    po_amt = dfs["purchase_orders"].set_index("po_id")["total_amount"].to_dict()
    mismatches = [
        r["invoice_id"] for _, r in dfs["invoices"].iterrows()
        if abs(r["taxable_amount"] - po_amt[r["po_id"]]) >= 0.01
    ]
    if mismatches == ["INV-2025-018"]:
        expected_anomalies_found.append("ANM-03 (INV-2025-018)")
    record("7_invoice_to_po_discrepancies", mismatches == ["INV-2025-018"], f"Found expected mismatch: {mismatches}")

    # 8. Duplicate invoice numbers
    dup_nums = dfs["invoices"]["invoice_number"].value_counts()
    dups = list(dup_nums[dup_nums > 1].index)
    if dups == ["SVE/25-26/0147"]:
        expected_anomalies_found.append("ANM-04 (SVE/25-26/0147)")
    record("8_duplicate_invoice_numbers", dups == ["SVE/25-26/0147"], f"Found expected duplicate invoice_number: {dups}")

    # 9. Payment reconciliation
    inv_tot = dfs["invoices"].set_index("invoice_id")["total_amount"].to_dict()
    inv_stat = dfs["invoices"].set_index("invoice_id")["payment_status"].to_dict()
    overpaid = [r["payment_id"] for _, r in dfs["payments"].iterrows() if r["payment_amount"] > inv_tot[r["invoice_id"]] + 0.01]
    partials = [r["payment_id"] for _, r in dfs["payments"].iterrows() if inv_stat[r["invoice_id"]] == "Partial"]
    if partials == ["PAY-015"]:
        expected_anomalies_found.append("ANM-12 (PAY-015)")
    record("9_payment_reconciliation", len(overpaid) == 0 and partials == ["PAY-015"], f"Overpaid={overpaid}, Partials={partials}")

    # 10. Delivery delay records
    delays = list(dfs["supplier_deliveries"].loc[dfs["supplier_deliveries"]["delay_days"] > 0, "delivery_id"])
    if delays == ["DEL-003", "DEL-008"]:
        expected_anomalies_found.extend(["ANM-01 (DEL-003)", "ANM-02 (DEL-008)"])
    record("10_delivery_delay_records", delays == ["DEL-003", "DEL-008"], f"Delayed deliveries: {delays}")

    # 11. Quality failures
    q_issues = list(dfs["quality_reports"].loc[dfs["quality_reports"]["defect_rate"] >= 0.09, "qr_id"])
    if set(q_issues) == {"QR-010", "QR-012"}:
        expected_anomalies_found.extend(["ANM-05 (QR-010)", "ANM-06 (QR-012)"])
    record("11_quality_failures", set(q_issues) == {"QR-010", "QR-012"}, f"High-defect quality reports: {q_issues}")

    # 12. Certificate status and expiry
    cert1 = (DOCS_DIR / "pdf" / "cert_SUP004_iso9001.pdf").read_bytes()
    cert2 = (DOCS_DIR / "pdf" / "cert_SUP002_iso9001.pdf").read_bytes()
    cert_ok = (b"2025-09-30" in cert1 and b"EXPIRED" in cert1 and b"2027-03-31" in cert2 and b"VALID" in cert2)
    if cert_ok:
        expected_anomalies_found.append("ANM-07 (DOC-PDF-001)")
    record("12_certificate_status_and_expiry", cert_ok, "SUP004 expired 2025-09-30; SUP002 valid 2027-03-31")

    # 13. Markdown document count
    md_files = list(DOCS_DIR.rglob("*.md"))
    record("13_markdown_document_count", len(md_files) == 12, f"Found {len(md_files)} Markdown documents")

    # 14. PDF document count and readable text
    pdf_files = list((DOCS_DIR / "pdf").glob("*.pdf"))
    pdf_text_ok = len(pdf_files) == 5 and all(b"%PDF-" in p.read_bytes() and b"DOC-PDF-" in p.read_bytes() for p in pdf_files)
    record("14_pdf_count_and_readable_text", pdf_text_ok, f"Found {len(pdf_files)} text-selectable PDFs")

    # 15. Excel worksheet names and record counts
    wb = openpyxl.load_workbook(PROCESSED_DIR / "procurement_data.xlsx", read_only=True)
    xl_errs = []
    if wb.sheetnames != list(EXPECTED_SCHEMAS.keys()):
        xl_errs.append(f"Sheets mismatch: {wb.sheetnames}")
    else:
        for sname, (exp_rows, _, _) in EXPECTED_SCHEMAS.items():
            ws = wb[sname]
            data_rows = ws.max_row - 1
            if data_rows != exp_rows:
                xl_errs.append(f"{sname}: expected {exp_rows}, got {data_rows}")
    wb.close()
    record("15_excel_worksheets_and_counts", len(xl_errs) == 0, "All 8 Excel sheets match CSV row counts" if not xl_errs else "; ".join(xl_errs))

    # 16. Evaluation scenario count
    scenarios = json.loads((EVAL_DIR / "decision_scenarios.json").read_text(encoding="utf-8"))
    record("16_evaluation_scenario_count", len(scenarios) == 10, f"Found {len(scenarios)} decision scenarios")

    # 17. Ground-truth references
    gt_anomalies = json.loads((EVAL_DIR / "ground_truth_anomalies.json").read_text(encoding="utf-8"))
    all_ids = (
        sups | prds | pos | invs
        | set(dfs["contracts"]["contract_id"])
        | set(dfs["supplier_deliveries"]["delivery_id"])
        | set(dfs["payments"]["payment_id"])
        | set(dfs["quality_reports"]["qr_id"])
        | {f"DOC-MD-{i:03d}" for i in range(1, 13)}
        | {f"DOC-PDF-{i:03d}" for i in range(1, 6)}
    )
    gt_errs = []
    for a in gt_anomalies:
        if not (BASE_DIR / a["source_file"]).exists():
            gt_errs.append(f"{a['anomaly_id']} missing file {a['source_file']}")
        if a["source_record_id"] not in all_ids:
            gt_errs.append(f"{a['anomaly_id']} unknown source_record_id {a['source_record_id']}")
        for rid in a["related_record_ids"]:
            if rid not in all_ids:
                gt_errs.append(f"{a['anomaly_id']} unknown related_record_id {rid}")
    record("17_ground_truth_references", len(gt_anomalies) == 12 and len(gt_errs) == 0, "All 12 ground-truth references resolve" if not gt_errs else "; ".join(gt_errs))

    # 18. Whether all planted anomalies are discoverable
    email_sup3 = (DOCS_DIR / "emails" / "email_SUP003_pricing_revision.md").read_text(encoding="utf-8")
    con_sup3 = (DOCS_DIR / "contracts" / "contract_SUP003_CON-2025-003.md").read_text(encoding="utf-8")
    mgmt_h1 = (DOCS_DIR / "management" / "management_report_H1_supplier_reliability.md").read_text(encoding="utf-8")
    email_sup1 = (DOCS_DIR / "emails" / "email_SUP001_performance_claim.md").read_text(encoding="utf-8")
    po23 = dfs["purchase_orders"].loc[dfs["purchase_orders"]["po_id"] == "PO-2025-023"].iloc[0]

    anm08_ok = "2,100.00" in email_sup3 and "1,850.00" in con_sup3
    anm09_ok = "one late delivery" in mgmt_h1 and (dfs["supplier_deliveries"].loc[dfs["supplier_deliveries"]["supplier_id"] == "SUP002", "delay_days"] == 0).all()
    anm10_ok = "95%" in email_sup1 and (dfs["supplier_deliveries"].loc[dfs["supplier_deliveries"]["supplier_id"] == "SUP001", "delay_days"] == 0).sum() == 3
    anm11_ok = po23["unit_price"] == 3400.0 and "Pending Buyer Sign-Off" in str(po23["approved_by"])

    if anm08_ok:
        expected_anomalies_found.append("ANM-08 (DOC-MD-003 vs DOC-MD-006)")
    if anm09_ok:
        expected_anomalies_found.append("ANM-09 (DOC-MD-012 vs deliveries)")
    if anm10_ok:
        expected_anomalies_found.append("ANM-10 (DOC-MD-007 vs deliveries)")
    if anm11_ok:
        expected_anomalies_found.append("ANM-11 (PO-2025-023)")

    discoverable = len(expected_anomalies_found) == 12
    record("18_all_planted_anomalies_discoverable", discoverable, f"Discovered {len(expected_anomalies_found)}/12 planted anomalies")

    overall_status = "PASS" if len(unexpected_errors) == 0 else "FAIL"
    report = {
        "final_overall_status": overall_status,
        "checks_passed": sum(1 for c in checks if c["status"] == "PASS"),
        "checks_total": len(checks),
        "checks": checks,
        "expected_anomalies_found": expected_anomalies_found,
        "unexpected_errors_found": unexpected_errors,
        "warnings": warnings,
    }
    REPORTS_DIR.mkdir(parents=True, exist_ok=True)
    (REPORTS_DIR / "validation_report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    return report


def main() -> None:
    report = run_validation()
    print(f"Validation Status: {report['final_overall_status']} ({report['checks_passed']}/{report['checks_total']} checks passed)")
    print(f"Expected Anomalies Found: {len(report['expected_anomalies_found'])}/12")
    if report["unexpected_errors_found"]:
        print("Unexpected Errors:")
        for err in report["unexpected_errors_found"]:
            print(f"  - {err}")
        sys.exit(1)
    sys.exit(0)


if __name__ == "__main__":
    main()
