# Automated Invoice Generator

A Python-based tool for automating the generation, export, and delivery of professional invoices from Excel order data.

## Quick Start

### Local Installation
```bash
# Clone and install
git clone https://github.com/yourusername/automated-invoice-generator.git
cd automated-invoice-generator
pip install -r requirements.txt

# Generate invoices
python sample_invoice.py
```

### Docker Installation
```bash
# Clone the repository
git clone https://github.com/yourusername/automated-invoice-generator.git
cd automated-invoice-generator

# Build and verify the Docker container
./verify_docker.sh
```

For detailed Docker setup and troubleshooting, see [Docker Support](Docker_README.md).

## Key Features

- 📊 Excel-based order processing
- 📝 Professional invoice formatting
- 📤 Multiple export formats (XLSX/PDF)
- 📧 Automated email delivery
- 🧮 Automatic tax calculations
- 🔄 PO number management
- ✅ Comprehensive test coverage
- 🐳 Cross-platform Docker support (Windows/Linux)

## Documentation

For detailed documentation, including:
- Complete installation guide
- Usage examples
- Input data format
- Docker support
- Contributing guidelines

Please see the [full documentation](docs/README.md).

## Requirements

### Local Installation
- Python 3.7+
- Windows OS (for PDF generation using Excel)
- Microsoft Excel (for PDF generation)

### Docker Installation
- Docker
- For PDF generation:
  - Windows: Microsoft Excel (via win32com)
  - Linux: LibreOffice (automatically installed in container)

## Docker Support

The application is containerized and supports both Windows and Linux environments. For detailed Docker documentation, see [Docker Support](Docker_README.md).

### Quick Docker Commands
```bash
# Build the container
docker build -t invoice-generator .

# Run a sample invoice generation
docker run --rm -v ${PWD}/screenshots:/app/screenshots invoice-generator python sample_invoice.py

# Run tests
docker run --rm invoice-generator python -m pytest tests/
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Features
- **LoadOrders**: Read `orders.xlsx` → list of dicts  
- **FormatInvoice**: Populate an invoice template with order data  
- **ExportInvoice**: Save invoices as `.xlsx` or `.pdf`  
- **SendInvoice**: Attach & send via Outlook  
- **GenerateAllInvoices**: One-click pipeline to process all orders  
- Fully covered by **pytest** tests under `tests/`

## Email Integration

The application includes Outlook email integration for sending invoices directly from the application. The email functionality is implemented in `email_utils.py` and includes the following features:

- Direct integration with Microsoft Outlook
- Support for multiple recipients and CC recipients
- File attachment support
- Comprehensive error handling and logging
- Unit tests for all email functionality

### Example Usage

```python
from email_utils import OutlookEmailSender

# Create email sender instance
sender = OutlookEmailSender()

# Connect to Outlook
if sender.connect():
    # Send email with attachment
    sender.send_email(
        to_recipients=["recipient@example.com"],
        subject="Invoice #123",
        body="Please find attached the invoice.",
        attachments=["invoice.pdf"]
    )
```

### Requirements for Email Functionality
- Windows OS
- Microsoft Outlook installed
- pywin32 package (included in requirements.txt)

## Screenshots

### Sample Invoice Sent Through Email
![Email Configuration](screenshots/test-email.png)

### Sample Invoice Generation
![Sample Invoice](screenshots/invoice_INV-20250608-0002.png)

## Application Architecture

Here's an overview of the application's components and their interactions:

![Component Diagram](docs/assets/component_diagram.png)

## Getting Started

### Prerequisites
- Python 3.11+  
- pip  

### Installation
```bash
git clone https://github.com/jessherna/automated-invoice-generator.git
cd automated-invoice-generator
pip install -r requirements.txt
```
