FROM python:3.9

ENV GOOGLE_API_KEY="AIzaSyAhhOb8cNZRLzH45YHOhTEWM968EAmgbt4"

WORKDIR /app

COPY . .

RUN pip install -r requirements.txt

EXPOSE 5000

CMD ["python", "app.py"]