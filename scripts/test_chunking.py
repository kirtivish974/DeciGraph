from pathlib import Path

from ingestion.pipeline import load_file
from ingestion.chunker import chunk_document


files = (
    list(Path("documents").rglob("*.md"))
    + list(Path("documents").rglob("*.pdf"))
    + list(Path("documents").rglob("*.html"))
)

total_chunks = 0
failed_files = 0

print("\nDOCUMENT CHUNKING AUDIT")
print("=" * 75)

for file_path in files:
    try:
        document = load_file(file_path)
        chunks = chunk_document(document)

        total_chunks += len(chunks)

        print(
            f"{document['file_name']:<48} "
            f"Chunks: {len(chunks):<4} "
            f"Status: {'PASS' if chunks else 'EMPTY'}"
        )

        for chunk in chunks:
            assert chunk["content"].strip()
            assert len(chunk["content"]) <= 700
            assert chunk["metadata"]["source_file"] == document["file_name"]

    except Exception as error:
        failed_files += 1
        print(f"FAILED: {file_path} | {error}")

print("=" * 75)
print(f"Documents tested: {len(files)}")
print(f"Total chunks: {total_chunks}")
print(f"Failed files: {failed_files}")