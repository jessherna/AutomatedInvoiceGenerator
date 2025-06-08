FROM python:3.11-slim AS base

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Add current directory to Python path
ENV PYTHONPATH=/app

# Default command (can be overridden)
CMD ["python", "sample_invoice.py"]
