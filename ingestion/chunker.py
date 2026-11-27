from langchain_text_splitters import RecursiveCharacterTextSplitter


text_splitter = RecursiveCharacterTextSplitter(
    chunk_size=700,
    chunk_overlap=100,
    separators=["\n\n", "\n", ". ", " ", ""],
    length_function=len,
)


def chunk_document(document):
    """
    Split a loaded document into smaller chunks.

    Supports:
    - Markdown and HTML
    - PDF pages
    """

    chunks = []

    if "pages" in document:
        # PDF: preserve page-level metadata
        for page in document["pages"]:
            content = page["content"]

            if not content.strip():
                continue

            page_chunks = text_splitter.split_text(content)

            for chunk in page_chunks:
                chunk_index = len(chunks)
                chunks.append({
                    "content": chunk,
                    "metadata": {
                        "source_file": document["file_name"],
                        "file_path": document["file_path"],
                        "document_type": "pdf",
                        "page_number": page["page_number"],
                        "chunk_index": chunk_index,
                    },
                })

    else:
        # Markdown / HTML: split the complete document
        content = document["content"]

        if content.strip():
            document_chunks = text_splitter.split_text(content)

            for chunk_index, chunk in enumerate(document_chunks):
                chunks.append({
                    "content": chunk,
                    "metadata": {
                        "source_file": document["file_name"],
                        "file_path": document["file_path"],
                        "document_type": document["document_type"],
                        "page_number": None,
                        "chunk_index": chunk_index,
                    },
                })

    return chunks