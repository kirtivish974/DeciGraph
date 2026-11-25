
from pathlib import Path
import pandas as pd

DATA = Path("data/raw")

# Load datasets
tables = {
    name: pd.read_csv(DATA / f"{name}.csv")
    for name in [
        "suppliers",
        "products",
        "contracts",
        "purchase_orders",
        "invoices",
        "payments",
        "supplier_deliveries",
        "quality_reports"
    ]
}

issues = []



def check(condition, message):
    if condition:
        print("[PASS]", message)
    else:
        issues.append(message)
        print("[FAIL]", message)


def check_unique(table, column):
    df = tables[table]
    duplicates = df[column].duplicated().sum()

    check(
        duplicates == 0,
        f"{table}.{column} has {duplicates} duplicate IDs"
    )


def check_fk(child, column, parent, parent_column):
    invalid = set(tables[child][column].dropna()) - set(
        tables[parent][parent_column].dropna()
    )

    check(
        len(invalid) == 0,
        f"{child}.{column} references valid {parent} IDs"
        if not invalid else
        f"{child}.{column} has invalid references: {invalid}"
    )


print("\n========== PRIMARY KEY CHECKS ==========")

primary_keys = {
    "suppliers": "supplier_id",
    "products": "product_id",
    "contracts": "contract_id",
    "purchase_orders": "po_id",
    "invoices": "invoice_id",
    "payments": "payment_id",
    "supplier_deliveries": "delivery_id",
    "quality_reports": "qr_id"
}

for table, column in primary_keys.items():
    check_unique(table, column)


print("\n========== FOREIGN KEY CHECKS ==========")

relationships = [
    ("contracts", "supplier_id", "suppliers", "supplier_id"),
    ("purchase_orders", "supplier_id", "suppliers", "supplier_id"),
    ("purchase_orders", "product_id", "products", "product_id"),
    ("invoices", "po_id", "purchase_orders", "po_id"),
    ("invoices", "supplier_id", "suppliers", "supplier_id"),
    ("payments", "invoice_id", "invoices", "invoice_id"),
    ("payments", "supplier_id", "suppliers", "supplier_id"),
    ("supplier_deliveries", "po_id", "purchase_orders", "po_id"),
    ("supplier_deliveries", "supplier_id", "suppliers", "supplier_id"),
    ("quality_reports", "po_id", "purchase_orders", "po_id"),
    ("quality_reports", "supplier_id", "suppliers", "supplier_id"),
    ("quality_reports", "product_id", "products", "product_id"),
]

for child, col, parent, parent_col in relationships:
    check_fk(child, col, parent, parent_col)


print("\n========== CROSS-TABLE CONSISTENCY ==========")

# Invoice supplier must match its purchase order supplier
invoice_check = tables["invoices"].merge(
    tables["purchase_orders"][["po_id", "supplier_id"]],
    on="po_id",
    suffixes=("_invoice", "_po")
)

mismatch = (
    invoice_check["supplier_id_invoice"]
    != invoice_check["supplier_id_po"]
).sum()

check(
    mismatch == 0,
    f"Invoice supplier matches PO supplier ({mismatch} mismatches)"
)

# Payment supplier must match invoice supplier
payment_check = tables["payments"].merge(
    tables["invoices"][["invoice_id", "supplier_id"]],
    on="invoice_id",
    suffixes=("_payment", "_invoice")
)

mismatch = (
    payment_check["supplier_id_payment"]
    != payment_check["supplier_id_invoice"]
).sum()

check(
    mismatch == 0,
    f"Payment supplier matches invoice supplier ({mismatch} mismatches)"
)

# Delivery supplier must match PO supplier
delivery_check = tables["supplier_deliveries"].merge(
    tables["purchase_orders"][["po_id", "supplier_id"]],
    on="po_id",
    suffixes=("_delivery", "_po")
)

mismatch = (
    delivery_check["supplier_id_delivery"]
    != delivery_check["supplier_id_po"]
).sum()

check(
    mismatch == 0,
    f"Delivery supplier matches PO supplier ({mismatch} mismatches)"
)

# Quality report supplier and product must match PO
quality_check = tables["quality_reports"].merge(
    tables["purchase_orders"][
        ["po_id", "supplier_id", "product_id"]
    ],
    on="po_id",
    suffixes=("_report", "_po")
)

supplier_mismatch = (
    quality_check["supplier_id_report"]
    != quality_check["supplier_id_po"]
).sum()

product_mismatch = (
    quality_check["product_id_report"]
    != quality_check["product_id_po"]
).sum()

check(
    supplier_mismatch == 0,
    f"Quality report supplier matches PO ({supplier_mismatch} mismatches)"
)

check(
    product_mismatch == 0,
    f"Quality report product matches PO ({product_mismatch} mismatches)"
)


print("\n========== FINANCIAL VALIDATION ==========")

# PO total = quantity * unit price
po = tables["purchase_orders"]

difference = (
    po["total_amount"] -
    (po["quantity"] * po["unit_price"])
).abs()

check(
    (difference < 0.01).all(),
    f"PO subtotal calculation ({(difference >= 0.01).sum()} mismatches)"
)

# PO grand total = subtotal + GST
difference = (
    po["grand_total"] -
    (po["total_amount"] + po["gst_amount"])
).abs()

check(
    (difference < 0.01).all(),
    f"PO grand total calculation ({(difference >= 0.01).sum()} mismatches)"
)

# Invoice total = taxable amount + GST
inv = tables["invoices"]

difference = (
    inv["total_amount"] -
    (inv["taxable_amount"] + inv["gst_amount"])
).abs()

check(
    (difference < 0.01).all(),
    f"Invoice total calculation ({(difference >= 0.01).sum()} mismatches)"
)

# Invoice GST = taxable amount * GST rate
difference = (
    inv["gst_amount"] -
    (inv["taxable_amount"] * inv["gst_rate"])
).abs()

check(
    (difference < 0.01).all(),
    f"Invoice GST calculation ({(difference >= 0.01).sum()} mismatches)"
)


print("\n========== DELIVERY VALIDATION ==========")

delivery = tables["supplier_deliveries"].copy()

expected = pd.to_datetime(
    delivery["expected_delivery_date"], errors="coerce"
)

actual = pd.to_datetime(
    delivery["actual_delivery_date"], errors="coerce"
)

calculated_delay = (actual - expected).dt.days.clip(lower=0)

difference = (
    delivery["delay_days"] - calculated_delay
).abs()

check(
    (difference == 0).all(),
    f"Delivery delay calculation ({(difference != 0).sum()} mismatches)"
)

check(
    (delivery["quantity_delivered"] <=
     delivery["quantity_ordered"]).all(),
    "Delivered quantity does not exceed ordered quantity"
)


print("\n========== QUALITY VALIDATION ==========")

quality = tables["quality_reports"]

calculated_rate = (
    quality["defects_found"] / quality["batch_size"]
)

difference = (
    quality["defect_rate"] - calculated_rate
).abs()

check(
    (difference < 0.0001).all(),
    f"Quality defect rate calculation ({(difference >= 0.0001).sum()} mismatches)"
)


print("\n========== FINAL SUMMARY ==========")

print("Total checks:", 32)
print("Issues found:", len(issues))

print("Issues found:", len(issues))

for issue in issues:
    print("-", issue)

# Save findings
report = Path("reports/day01_relationship_audit.txt")
report.parent.mkdir(parents=True, exist_ok=True)

report.write_text(
    "DECISIONGRAPH RELATIONSHIP AUDIT\n\n"
    + "\n".join(issues if issues else ["No issues detected."]),
    encoding="utf-8"
)

print("\nReport saved:", report)
