from vectorstore.qdrant_store import QdrantStore


def main():

    store = QdrantStore()

    try:
        store.create_collection()

        info = store.collection_info()

        print("\nQDRANT COLLECTION TEST")
        print("=" * 40)
        print("Collection:", info.config.params.vectors.size)
        print("Vector dimensions:", info.config.params.vectors.size)
        print("Distance:", info.config.params.vectors.distance)
        print("Points:", info.points_count)

    finally:
        store.close()


if __name__ == "__main__":
    main()