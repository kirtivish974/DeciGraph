from database.db import engine, Base
from database import models


def initialize_database():
    # Create all tables registered with Base
    Base.metadata.create_all(bind=engine)

    print("Database initialized successfully!")

    print("\nTables created:")

    for table_name in Base.metadata.tables:
        print(f"- {table_name}")


if __name__ == "__main__":
    initialize_database()