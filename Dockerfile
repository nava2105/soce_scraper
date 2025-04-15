FROM python:3.9

ENV GOOGLE_API_KEY="YOUR_GEMINI_API_KEY_HERE"

WORKDIR /app

COPY . .

RUN pip install -r requirements.txt

EXPOSE 5000

CMD ["python", "app.py"]