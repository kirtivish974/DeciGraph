import pandas as pd


REQUIRED_COLUMNS = {
    "suppliers": ["supplier_id", "supplier_name"],
    "products": ["product_id", "product_name"],
    "purchase_orders": ["po_id", "supplier_id"],
    "supplier_deliveries": ["delivery_id", "po_id"],
    "invoices": ["invoice_id", "supplier_id"],
    "payments": ["payment_id", "invoice_id"],
    "contracts": ["contract_id", "supplier_id"],
    "quality_reports": ["qr_id", "supplier_id"],
}


BUSINESS_KEYS = {
    "suppliers": [("supplier_id",)],
    "products": [("product_id",)],
    "contracts": [("contract_id",)],
    "purchase_orders": [("po_id",)],
    "invoices": [
        ("invoice_id",),
        ("supplier_id", "invoice_number")
    ],
    "payments": [("payment_id",)],
    "supplier_deliveries": [("delivery_id",)],
    "quality_reports": [("qr_id",)],
}

DATE_COLUMNS = {
    "suppliers": [
        "registration_date"
    ],
    "purchase_orders": [
        "order_date",
        "expected_delivery_date"
    ],
    "supplier_deliveries": [
        "expected_delivery_date",
        "actual_delivery_date"
    ],
    "invoices": [
        "invoice_date",
        "due_date"
    ],
    "payments": [
        "payment_date"
    ],
    "contracts": [
        "contract_start_date",
        "contract_end_date"
    ],
    "quality_reports": [
        "inspection_date"
    ]
}

def validate_dataframe(df, dataset_name):
    """
    Validate a DataFrame and report potential data-quality issues.

    Records are never deleted or modified.
    """

    # 1. Check required columns
    required = REQUIRED_COLUMNS.get(dataset_name, [])

    missing_columns = [
        column
        for column in required
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"{dataset_name}: Missing required columns: "
            f"{missing_columns}"
        )

    # Convert configured date columns
    date_columns = DATE_COLUMNS.get(dataset_name, [])

    for column in date_columns:
        if column in df.columns:
            try:
                df[column] = pd.to_datetime(
                    df[column],
                    errors="raise"
                )
            except (ValueError, TypeError) as error:
                raise ValueError(
                    f"{dataset_name}: Invalid date in "
                    f"column '{column}': {error}"
                )

    # 2. Detect missing values
    missing_values = {
        column: int(count)
        for column, count in df.isnull().sum().items()
        if count > 0
    }

    # 3. Detect exact duplicate rows
    duplicate_count = int(df.duplicated().sum())

    # 4. Detect business-level duplicate identifiers
    business_duplicates = {}

    key_groups = BUSINESS_KEYS.get(dataset_name, [])

    for columns in key_groups:

        # Ignore rows with missing business-key values.
        # Missing identifiers should be reported separately.
        valid_rows = df.dropna(subset=list(columns))

        duplicate_mask = valid_rows.duplicated(
            subset=list(columns),
            keep=False
        )

        duplicate_rows = valid_rows.loc[duplicate_mask]

        if not duplicate_rows.empty:

            duplicate_values = (
                duplicate_rows[list(columns)]
                .drop_duplicates()
                .to_dict(orient="records")
            )

            key_name = " + ".join(columns)

            business_duplicates[key_name] = duplicate_values

    return {
        "missing_values": missing_values,
        "duplicate_rows": duplicate_count,
        "business_duplicates": business_duplicates
    }