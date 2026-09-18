FROM python:3.11-slim

WORKDIR /app

ENV PYTHONUNBUFFERED=1

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

# Data (token cache, subscription info) lives here - mount a volume onto it.
RUN mkdir -p /data
ENV DATA_DIR=/data

EXPOSE 5000

CMD ["python", "app.py"]