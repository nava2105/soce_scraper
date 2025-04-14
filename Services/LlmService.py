import os
import google.generativeai as genai

def configure_gemini_api():
    google_api_key = os.getenv('GOOGLE_API_KEY')
    if google_api_key:
        genai.configure(api_key=google_api_key)

def generate_text_embeddings(text_chunks):
    embeddings = []
    for chunk in text_chunks:
        try:
            result = genai.embed_content(
                model="models/embedding-001",
                content=chunk,
                task_type="retrieval_document",
                title="Document Chunk"
            )
            embeddings.append((chunk, result['embedding']))
        except Exception as e:
            print(f"Error generating embedding: {e}")
    return embeddings

def generate_ai_response(prompt):
    model = genai.GenerativeModel("gemini-1.5-flash")
    response = model.generate_content(prompt)
    return response.text
