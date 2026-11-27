from pathlib import Path

from qdrant_client import QdrantClient
from qdrant_client.models import Distance, VectorParams


BASE_DIR = Path(__file__).resolve().parent.parent

COLLECTION_NAME = "decigraph_documents"
VECTOR_SIZE = 768


class QdrantStore:

    def __init__(self):

        self.storage_path = BASE_DIR / "data" / "vectorstore"

        self.storage_path.mkdir(
            parents=True,
            exist_ok=True
        )

        self.client = QdrantClient(
            path=str(self.storage_path)
        )

    def create_collection(self):

        if self.client.collection_exists(COLLECTION_NAME):
            print(f"Collection '{COLLECTION_NAME}' already exists.")
            return

        self.client.create_collection(
            collection_name=COLLECTION_NAME,
            vectors_config=VectorParams(
                size=VECTOR_SIZE,
                distance=Distance.COSINE
            )
        )

        print(f"Collection '{COLLECTION_NAME}' created successfully.")

    def collection_info(self):

        if not self.client.collection_exists(COLLECTION_NAME):
            raise ValueError(
                f"Collection '{COLLECTION_NAME}' does not exist."
            )

        return self.client.get_collection(COLLECTION_NAME)

    def close(self):

        self.client.close()