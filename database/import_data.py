
import csv
from decimal import Decimal
from pathlib import Path

import pandas as pd
from sqlalchemy.orm import Session

from database.db import engine
from database.models import (
    Supplier,
    Product,
    Contract,
    PurchaseOrder,
    Invoice,
    Payment,
    SupplierDelivery,
    QualityReport,
)


BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data" / "raw"


# Import in foreign-key dependency order
TABLE_MODELS = [
    ("suppliers", Supplier),
    ("products", Product),
    ("contracts", Contract),
    ("purchase_orders", PurchaseOrder),
    ("invoices", Invoice),
    ("payments", Payment),
    ("supplier_deliveries", SupplierDelivery),
    ("quality_reports", QualityReport),
]


DATE_COLUMNS = {
    "registration_date",
    "contract_start_date",
    "contract_end_date",
    "order_date",
    "expected_delivery_date",
    "invoice_date",
    "due_date",
    "payment_date",
    "actual_delivery_date",
    "inspection_date",
}

DECIMAL_COLUMNS = {
    "contract_value",
    "standard_unit_price",
    "gst_rate",
    "unit_price",
    "total_amount",
    "gst_amount",
    "grand_total",
    "taxable_amount",
    "payment_amount",
    "defect_rate",
}

INTEGER_COLUMNS = {
    "quantity",
    "batch_size",
    "defects_found",
    "quantity_ordered",
    "quantity_delivered",
    "delay_days",
}

BOOLEAN_COLUMNS = {
    "auto_renewal",
}


def convert_value(column, value):
    """Convert CSV values into appropriate Python types."""

    if pd.isna(value):
        return None

    if column in DATE_COLUMNS:
        return pd.to_datetime(value, errors="raise").date()

    if column in DECIMAL_COLUMNS:
        return Decimal(str(value))

    if column in INTEGER_COLUMNS:
        return int(value)

    if column in BOOLEAN_COLUMNS:
        normalized = str(value).strip().lower()

        if normalized in {"true", "1", "yes"}:
            return True

        if normalized in {"false", "0", "no"}:
            return False

        raise ValueError(
            f"Invalid boolean value for {column}: {value}"
        )

    return str(value)


def import_table(session, table_name, model):
    """Import one CSV file into its corresponding database table."""

    file_path = DATA_DIR / f"{table_name}.csv"

    if not file_path.exists():
        raise FileNotFoundError(
            f"Dataset not found: {file_path}"
        )

    # Read values as strings to avoid losing identifier formatting.
    df = pd.read_csv(file_path, dtype=str)

    imported_count = 0

    for row_number, row in enumerate(
        df.to_dict(orient="records"), start=2
    ):
        converted_record = {
            column: convert_value(column, value)
            for column, value in row.items()
        }

        record = model(**converted_record)

        session.add(record)
        imported_count += 1

    return imported_count


def import_all_data():
    """Import all eight datasets in one transaction."""

    results = {}

    try:
        with Session(engine) as session:
            with session.begin():

                for table_name, model in TABLE_MODELS:
                    count = import_table(
                        session,
                        table_name,
                        model
                    )

                    results[table_name] = count

    except Exception:
        # The transaction is rolled back automatically.
        raise

    print("\nDATABASE IMPORT COMPLETED")
    print("=" * 40)

    for table_name, count in results.items():
        print(f"{table_name}: {count} records")

    print("=" * 40)
    print(f"Total imported: {sum(results.values())}")


if __name__ == "__main__":
    import_all_data()