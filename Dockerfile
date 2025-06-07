FROM python:3.11-slim
WORKDIR /app

# 1) Install only the deps you need for Issue 6
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 2) Copy your application code explicitly
COPY invoice.py .

# 3) Copy tests (and any data folders)
COPY tests/ tests/
COPY tests/data/ tests/data/

# 4) Add current directory to Python path
ENV PYTHONPATH=/app

# 5) Run pytest against /app
CMD ["pytest", "-q", "--maxfail=1"]
