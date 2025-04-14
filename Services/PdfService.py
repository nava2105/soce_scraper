import os
import PyPDF2
import fitz  # PyMuPDF


def allowed_file(filename, allowed_extensions={'pdf'}):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


def extract_text_chunks(pdf_path, chunk_size=1000):
    text_chunks = []
    text = ""

    try:
        with open(pdf_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            text = "".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        print(f"PyPDF2 failed: {e}, trying PyMuPDF...")
        doc = fitz.open(pdf_path)
        text = "".join(page.get_text() for page in doc)

    for i in range(0, len(text), chunk_size):
        text_chunks.append(text[i:i + chunk_size])

    return text_chunks
