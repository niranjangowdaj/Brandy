import sys
import json
import numpy as np
from docx import Document
from sentence_transformers import SentenceTransformer

def extract_docx_text(docx_path):
    doc = Document(docx_path)
    return "\n".join([para.text for para in doc.paragraphs])

def chunk_text(text, chunk_size=500):
    words = text.split()
    return [' '.join(words[i:i+chunk_size]) for i in range(0, len(words), chunk_size)]

def main(docx_path, out_prefix):
    model = SentenceTransformer('all-MiniLM-L6-v2')
    text = extract_docx_text(docx_path)
    chunks = chunk_text(text)
    embeddings = model.encode(chunks)
    np.savez_compressed(f"{out_prefix}_embeddings.npz", embeddings=embeddings)
    with open(f"{out_prefix}_chunks.json", "w") as f:
        json.dump(chunks, f)
    print(f"Saved {len(chunks)} chunks and embeddings to {out_prefix}_embeddings.npz and {out_prefix}_chunks.json")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python embed_docx.py <input.docx> <output_prefix>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])