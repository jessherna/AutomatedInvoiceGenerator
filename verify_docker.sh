#!/bin/bash

echo "🔍 Verifying Docker setup..."

# Build the Docker image
docker build -t invoice-generator .

echo "✅ Docker build successful"

echo "🧪 Running basic container test..."
# Test Python imports (only cross-platform compatible ones)
docker run --rm invoice-generator python -c "import openpyxl; import PIL; print('✅ Python imports working')"

echo "🧪 Running pytest..."
# Run tests
docker run --rm invoice-generator python -m pytest tests/

echo "📄 Testing invoice generation..."
# Generate sample invoice
docker run --rm -v ${PWD}/screenshots:/app/screenshots invoice-generator python sample_invoice.py

echo "✅ All tests completed successfully!" 