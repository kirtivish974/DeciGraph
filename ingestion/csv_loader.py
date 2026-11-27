
from pathlib import Path
import pandas as pd

from ingestion.validation import validate_dataframe


def load_csv(file_path):
    """
    Load a CSV file and preserve its source information.
    """

    file_path = Path(file_path)

    # 1. Check whether file exists
    if not file_path.exists():
        raise FileNotFoundError(
            f"CSV file not found: {file_path}"
        )

    # 2. Check file extension
    if file_path.suffix.lower() != ".csv":
        raise ValueError(
            f"Unsupported file type: {file_path.suffix}"
        )

    # 3. Read CSV
    try:
        df = pd.read_csv(file_path)

    except pd.errors.EmptyDataError:
        raise ValueError("CSV file is empty.")

    except pd.errors.ParserError as error:
        raise ValueError(f"Malformed CSV file: {error}")

    except UnicodeDecodeError:
        raise ValueError("CSV encoding is not supported.")

    # 4. Identify dataset using filename
    dataset_name = file_path.stem.lower()

    # 5. Run shared validation
    quality_checks = validate_dataframe(
        df,
        dataset_name
    )

    # 6. Preserve source metadata for every row
    records = []

    for index, row in df.iterrows():
        records.append({
            "source_file": file_path.name,
            "dataset": dataset_name,
            "row_number": int(index) + 2,
            "content": row.to_dict()
        })

    # 7. Return extracted data and audit information
    return {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "dataset": dataset_name,
        "total_rows": len(df),
        "columns": df.columns.tolist(),
        "records": records,
        "quality_checks": quality_checks
    }