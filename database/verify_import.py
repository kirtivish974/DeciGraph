from sqlalchemy import text

from database.db import engine


TABLES = [
    "suppliers",
    "products",
    "contracts",
    "purchase_orders",
    "invoices",
    "payments",
    "supplier_deliveries",
    "quality_reports",
]


def verify_database():
    with engine.connect() as connection:

        print("DATABASE ROW COUNTS")
        print("=" * 40)

        total_records = 0

        for table in TABLES:
            count = connection.execute(
                text(f"SELECT COUNT(*) FROM {table}")
            ).scalar_one()

            print(f"{table}: {count}")
            total_records += count

        print(f"\nTotal records: {total_records}")

        # Check foreign-key violations
        violations = connection.execute(
            text("PRAGMA foreign_key_check")
        ).fetchall()

        print("\nFOREIGN KEY CHECK")
        print("Violations:", len(violations))

        # Check duplicate invoice numbers
        duplicates = connection.execute(
            text("""
                SELECT invoice_number, COUNT(*) AS occurrences
                FROM invoices
                GROUP BY invoice_number
                HAVING COUNT(*) > 1
            """)
        ).fetchall()

        print("\nDUPLICATE INVOICE NUMBERS")
        print("Found:", len(duplicates))

        for invoice_number, count in duplicates:
            print(f"{invoice_number}: {count} occurrences")


if __name__ == "__main__":
    verify_database()