FROM python:3.11-slim

# Install Tesseract OCR + language data + OpenCV system deps
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-eng \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Create folders the app writes to
RUN mkdir -p /tmp/uploads /tmp/reports

EXPOSE 8000

CMD ["python", "app.py"]
