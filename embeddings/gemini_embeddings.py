import os
from pathlib import Path

from dotenv import load_dotenv
from google import genai
from google.genai import types


BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")


class GeminiEmbeddings:

    def __init__(self):
        api_key = os.getenv("GEMINI_API_KEY")

        if not api_key:
            raise ValueError("GEMINI_API_KEY not found in .env")

        self.client = genai.Client(api_key=api_key)
        self.model = "gemini-embedding-2"
        self.dimensions = 768

    def embed_text(self, text):
        if not text or not text.strip():
            raise ValueError("Text cannot be empty")

        response = self.client.models.embed_content(
            model=self.model,
            contents=text,
            config=types.EmbedContentConfig(
                output_dimensionality=self.dimensions
            ),
        )

        return response.embeddings[0].values

    def embed_documents(self, texts):
        if not texts:
            return []

        return [self.embed_text(text) for text in texts]

    def embed_query(self, query):
        return self.embed_text(query)