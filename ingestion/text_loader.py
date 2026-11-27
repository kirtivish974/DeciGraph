
from pathlib import Path
from html.parser import HTMLParser


class HTMLTextExtractor(HTMLParser):
    """Extract readable text from HTML."""

    def __init__(self):
        super().__init__()
        self.text_parts = []
        self.skip_content = False

    def handle_starttag(self, tag, attrs):
        if tag in {"script", "style", "noscript"}:
            self.skip_content = True

    def handle_endtag(self, tag):
        if tag in {"script", "style", "noscript"}:
            self.skip_content = False

    def handle_data(self, data):
        if not self.skip_content:
            text = data.strip()
            if text:
                self.text_parts.append(text)


def load_text_document(file_path):
    """
    Load Markdown or HTML while preserving source metadata.
    """

    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(
            f"Document not found: {file_path}"
        )

    extension = file_path.suffix.lower()

    if extension not in {".md", ".html", ".htm"}:
        raise ValueError(
            f"Unsupported text document format: {extension}"
        )

    try:
        raw_text = file_path.read_text(encoding="utf-8")
    except UnicodeDecodeError as error:
        raise ValueError(
            f"Unsupported document encoding: {error}"
        )

    if extension in {".html", ".htm"}:
        parser = HTMLTextExtractor()
        parser.feed(raw_text)
        content = "\n".join(parser.text_parts)
    else:
        content = raw_text.strip()

    return {
        "file_name": file_path.name,
        "file_path": str(file_path),
        "document_type": extension.lstrip("."),
        "content": content,
        "quality_checks": {
            "empty_document": not bool(content.strip())
        }
    }