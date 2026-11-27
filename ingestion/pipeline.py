from pathlib import Path

from ingestion.csv_loader import load_csv
from ingestion.excel_loader import load_excel
from ingestion.pdf_loader import load_pdf
from ingestion.text_loader import load_text_document


def load_file(file_path):
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    extension = file_path.suffix.lower()

    if extension == ".csv":
        return load_csv(file_path)

    elif extension in [".xlsx", ".xls"]:
        return load_excel(file_path)

    elif extension == ".pdf":
        return load_pdf(file_path)

    elif extension in [".md", ".html", ".htm"]:
        return load_text_document(file_path)

    else:
        raise ValueError(f"Unsupported file format: {extension}")