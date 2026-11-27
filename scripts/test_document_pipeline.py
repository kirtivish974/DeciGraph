from pathlib import Path
from ingestion.pipeline import load_file

files = (
    list(Path("documents").rglob("*.md"))
    + list(Path("documents").rglob("*.pdf"))
    + list(Path("documents").rglob("*.html"))
)

print("\nDOCUMENT PIPELINE TEST")
print("=" * 60)

successful = 0
failed = 0

for file_path in files:
    try:
        result = load_file(file_path)

        if "pages" in result:
            content_length = sum(
                len(page["content"]) for page in result["pages"]
            )
            document_type = "pdf"
        else:
            content_length = len(result["content"])
            document_type = result["document_type"]

        print(
            f"{result['file_name']} | "
            f"Type: {document_type} | "
            f"Content: {content_length} characters"
        )

        successful += 1

    except Exception as error:
        print(f"FAILED: {file_path} | Error: {error}")
        failed += 1

print("=" * 60)
print(f"Total files: {len(files)}")
print(f"Successful: {successful}")
print(f"Failed: {failed}")