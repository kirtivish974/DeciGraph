
from pathlib import Path
import fitz
import json

ROOT = Path(".")
REPORT = ROOT / "reports/document_content_audit.txt"

with open("SOURCE_MANIFEST.json", encoding="utf-8") as f:
    manifest = json.load(f)

results = []

def record(text):
    results.append(text)
    print(text)

record("=" * 65)
record("DECISIONGRAPH DOCUMENT CONTENT AUDIT")
record("=" * 65)

# Markdown audit
record("\n========== MARKDOWN DOCUMENTS ==========")

for file_path in manifest["markdown_documents"]:
    path = ROOT / file_path

    if not path.exists():
        record(f"[FAIL] Missing: {file_path}")
        continue

    content = path.read_text(encoding="utf-8")
    lines = content.splitlines()

    record(f"\nFILE: {file_path}")
    record(f"Characters: {len(content)}")
    record(f"Lines: {len(lines)}")

    headings = [
        line for line in lines
        if line.strip().startswith("#")
    ]

    record(f"Headings: {headings}")

    if len(content.strip()) == 0:
        record("[FAIL] Empty document")
    else:
        record("[PASS] Document contains text")

# PDF audit
record("\n========== PDF DOCUMENTS ==========")

for file_path in manifest["pdf_documents"]:
    path = ROOT / file_path

    if not path.exists():
        record(f"[FAIL] Missing: {file_path}")
        continue

    try:
        with fitz.open(path) as pdf:
            page_count = len(pdf)
            extracted_text = "\n".join(
                page.get_text() for page in pdf
            )

        record(f"\nFILE: {file_path}")
        record(f"Pages: {page_count}")
        record(f"Extracted characters: {len(extracted_text)}")

        if len(extracted_text.strip()) == 0:
            record("[WARNING] No extractable text. May be scanned.")
        else:
            record("[PASS] Text extraction successful")

    except Exception as e:
        record(f"[FAIL] PDF error: {e}")

record("\n========== AUDIT SUMMARY ==========")
record(f"Markdown documents: {len(manifest['markdown_documents'])}")
record(f"PDF documents: {len(manifest['pdf_documents'])}")
record("\nDocument audit completed.")

REPORT.parent.mkdir(parents=True, exist_ok=True)
REPORT.write_text("\n".join(results), encoding="utf-8")

print(f"\nReport saved: {REPORT}")
