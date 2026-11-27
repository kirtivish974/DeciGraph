import os
from pathlib import Path

from dotenv import load_dotenv
from google import genai
from google.genai import types


# Load API key from project root
BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

api_key = os.getenv("GEMINI_API_KEY")

if not api_key:
    raise ValueError("GEMINI_API_KEY not found in .env file")


# Initialize Gemini client
client = genai.Client(api_key=api_key)

# Test embedding generation
response = client.models.embed_content(
    model="gemini-embedding-2",
    contents="A supplier delivered defective industrial components.",
    config=types.EmbedContentConfig(
        output_dimensionality=768
    ),
)

embedding = response.embeddings[0].values

print("GEMINI EMBEDDING TEST")
print("=" * 40)
print("Model: gemini-embedding-2")
print("Embedding dimensions:", len(embedding))
print("Vector type:", type(embedding).__name__)
print("First 5 values:", embedding[:5])
print("API connection: SUCCESS")