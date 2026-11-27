from pathlib import Path
import pymupdf


def load_pdf(file_path):
    """
    Extract text from a PDF page by page
    while preserving source metadata.
    """

    file_path = Path(file_path)

    # 1. Check whether file exists
    if not file_path.exists():
        raise FileNotFoundError(
            f"PDF file not found: {file_path}"
        )

    # 2. Check file extension
    if file_path.suffix.lower() != ".pdf":
        raise ValueError(
            f"Unsupported file type: {file_path.suffix}"
        )

    # 3. Open PDF
    try:
        pdf = pymupdf.open(file_path)
    except Exception as error:
        raise ValueError(
            f"Could not open PDF: {error}"
        )

    # 4. Extract text page by page
    pages = []

    try:
        for page_number, page in enumerate(pdf, start=1):

            text = page.get_text("text").strip()

            pages.append({
                "source_file": file_path.name,
                "file_path": str(file_path),
                "page_number": page_number,
                "total_pages": len(pdf),
                "content": text
            })

    finally:
        pdf.close()

    # 5. Return extracted content and metadata
    return {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "total_pages": len(pages),
        "pages": pages,
        "quality_checks": {
            "empty_pages": [
                page["page_number"]
                for page in pages
                if not page["content"]
            ]
        }
    }