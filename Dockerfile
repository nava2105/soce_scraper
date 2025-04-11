FROM python:3.9

ENV GOOGLE_API_KEY="Here your GEMINI API KEY"

WORKDIR /app

COPY . .

RUN pip install -r requirements.txt

EXPOSE 5000

CMD ["python", "app.py"]