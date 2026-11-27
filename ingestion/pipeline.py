from pathlib import Path

from ingestion.csv_loader import load_csv
from ingestion.excel_loader import load_excel
from ingestion.pdf_loader import load_pdf


def load_file(file_path):
    """
    Automatically route a file to its appropriate loader.
    """

    file_path = Path(file_path)

    # Check whether file exists
    if not file_path.exists():
        raise FileNotFoundError(
            f"File not found: {file_path}"
        )

    # Identify file extension
    extension = file_path.suffix.lower()

    # Route to appropriate loader
    if extension == ".csv":
        return load_csv(file_path)

    elif extension in [".xlsx", ".xls"]:
        return load_excel(file_path)

    elif extension == ".pdf":
        return load_pdf(file_path)

    else:
        raise ValueError(
            f"Unsupported file format: {extension}"
        )