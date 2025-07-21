FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    poppler-utils \
    tesseract-ocr \
    unrar-free \
    libreoffice \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app files
COPY . .

# Set env var for Google credentials (optional)
ENV GOOGLE_APPLICATION_CREDENTIALS=/app/service_account.json

# Expose FastAPI port
EXPOSE 8001

# Run the app
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8001"]

