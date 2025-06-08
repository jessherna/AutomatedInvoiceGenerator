# Docker Support

This project includes Docker support for both development and testing environments, with cross-platform compatibility for Windows and Linux systems.

## Prerequisites

- Docker
- Docker Compose (optional, but recommended)
- For Windows PDF generation: Microsoft Excel installed on host system

## Quick Start

### Using Docker Directly

1. Build the image:
```bash
docker build -t invoice-generator .
```

2. Run the application:
```bash
# For Linux containers
docker run --rm -v ${PWD}/screenshots:/app/screenshots invoice-generator python sample_invoice.py

# For Windows containers (requires Windows host)
docker run --rm -v ${PWD}/screenshots:/app/screenshots invoice-generator python sample_invoice.py
```

3. Run tests:
```bash
docker run --rm invoice-generator python -m pytest tests/
```

### Using Docker Compose

1. Run the application:
```bash
docker-compose up app
```

2. Run tests:
```bash
docker-compose up test
```

## Platform-Specific Features

### Windows Environment
- Uses win32com for PDF generation
- Requires Microsoft Excel to be installed on the host system
- Supports native Windows COM automation

### Linux Environment
- Uses LibreOffice for PDF generation
- No additional host system requirements
- Fully containerized solution

## Development Workflow

The Docker setup includes:

- Volume mounting for live code updates
- Separate services for running the app and tests
- Environment variables for better debugging
- Cross-platform PDF generation support
- Platform-specific dependency management

## Container Details

- Base image: Python 3.11
- PDF Generation:
  - Windows: Microsoft Excel via win32com
  - Linux: LibreOffice (pre-installed)
- All Python dependencies are installed in the container
- Working directory is mounted as a volume for live updates

## Troubleshooting

### Common Issues

1. PDF Generation Fails
   - Windows: Ensure Microsoft Excel is installed on the host
   - Linux: Verify LibreOffice installation in container
   - Check container logs for specific error messages

2. Volume Mount Issues
   - Ensure correct path format for your OS
   - Check file permissions
   - Verify Docker volume settings

3. Resource Allocation
   - Ensure Docker has enough memory and CPU
   - Check disk space for large PDF generations
   - Monitor container resource usage

4. Platform-Specific Issues
   - Windows: COM automation errors
   - Linux: LibreOffice conversion issues
   - Check platform-specific logs

### Debugging Tips

1. View container logs:
```bash
docker logs invoice-generator
```

2. Access container shell:
```bash
docker exec -it invoice-generator /bin/bash
```

3. Check container status:
```bash
docker ps -a
```

4. Verify environment:
```bash
docker run --rm invoice-generator python -c "import sys; print(sys.platform)"
```