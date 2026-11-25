
from pathlib import Path
import pandas as pd

DATA_DIR = Path("data/raw")

csv_files = list(DATA_DIR.glob("*.csv"))

print("=" * 65)
print("           DECISIONGRAPH - DATASET AUDIT")
print("=" * 65)

for file in csv_files:

    print(f"\n\nFILE: {file.name}")
    print("-" * 50)

    try:
        df = pd.read_csv(file)

        print("Rows:", df.shape[0])
        print("Columns:", df.shape[1])

        print("\nColumn names:")
        print(list(df.columns))

        print("\nData types:")
        print(df.dtypes)

        print("\nMissing values:")
        print(df.isnull().sum())

        print("\nDuplicate rows:", df.duplicated().sum())

        print("\nSample records:")
        print(df.head(3).to_string(index=False))

    except Exception as e:
        print("ERROR:", e)

print("\n" + "=" * 65)
print("AUDIT COMPLETED")
print("=" * 65)
