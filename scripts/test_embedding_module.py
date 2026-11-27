from embeddings.gemini_embeddings import GeminiEmbeddings


def main():
    embeddings = GeminiEmbeddings()

    # Test single text embedding
    text = "Supplier delivered defective industrial components."

    vector = embeddings.embed_text(text)

    print("SINGLE TEXT EMBEDDING")
    print("=" * 40)
    print("Dimensions:", len(vector))
    print("Vector type:", type(vector).__name__)

    # Test document embedding
    documents = [
        "Supplier delivered defective components.",
        "Invoice payment is overdue.",
        "Purchase order delivery was delayed.",
    ]

    vectors = embeddings.embed_documents(documents)

    print("\nDOCUMENT EMBEDDING")
    print("=" * 40)
    print("Documents:", len(documents))
    print("Generated vectors:", len(vectors))
    print("Vector dimensions:", len(vectors[0]))

    # Test query embedding
    query = "Which supplier has delivery delays?"

    query_vector = embeddings.embed_query(query)

    print("\nQUERY EMBEDDING")
    print("=" * 40)
    print("Query:", query)
    print("Query dimensions:", len(query_vector))

    print("\nAll embedding tests passed!")


if __name__ == "__main__":
    main()