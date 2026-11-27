
from pathlib import Path
import pandas as pd

from ingestion.validation import validate_dataframe


def load_excel(file_path):
    """
    Load all sheets from an Excel workbook
    and preserve source information.
    """

    file_path = Path(file_path)

    # 1. Check whether file exists
    if not file_path.exists():
        raise FileNotFoundError(
            f"Excel file not found: {file_path}"
        )

    # 2. Check file extension
    if file_path.suffix.lower() not in [".xlsx", ".xls"]:
        raise ValueError(
            f"Unsupported Excel file type: {file_path.suffix}"
        )

    # 3. Read all sheets
    try:
        sheets = pd.read_excel(
            file_path,
            sheet_name=None
        )
    except Exception as error:
        raise ValueError(
            f"Could not read Excel workbook: {error}"
        )

    # 4. Store results for every sheet
    results = {}

    for sheet_name, df in sheets.items():

        dataset_name = sheet_name.strip().lower()

        # 5. Run shared validation
        quality_checks = validate_dataframe(
            df,
            dataset_name
        )

        # 6. Preserve source metadata
        records = []

        for index, row in df.iterrows():
            records.append({
                "source_file": file_path.name,
                "sheet_name": sheet_name,
                "dataset": dataset_name,
                "row_number": int(index) + 2,
                "content": row.to_dict()
            })

        # 7. Store sheet results
        results[dataset_name] = {
            "sheet_name": sheet_name,
            "total_rows": len(df),
            "columns": df.columns.tolist(),
            "records": records,
            "quality_checks": quality_checks
        }

    return {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "total_sheets": len(results),
        "sheets": results
    }