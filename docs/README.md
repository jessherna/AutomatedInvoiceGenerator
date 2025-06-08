# Automated Invoice Generator

A Python-based application for automating the generation of professional invoices from order data. This tool processes order information from Excel spreadsheets and generates formatted invoices in both Excel (.xlsx) and PDF formats, with support for both Windows and Linux environments.

## Features

- **Excel-based Data Processing**: Reads order data from structured Excel workbooks
- **Professional Invoice Formatting**: Generates clean, professional-looking invoices
- **Multiple Output Formats**: Supports both Excel (.xlsx) and PDF output
- **Automated Email Delivery**: Sends generated invoices directly to customers
- **Customizable Templates**: Includes company logo and customizable styling
- **Tax Calculation**: Automatically calculates GST (5%)
- **PO Number Management**: Automated PO number generation and tracking
- **Cross-Platform Support**: Works on both Windows and Linux environments
- **Docker Integration**: Containerized deployment with platform-specific optimizations

## Prerequisites

### Local Installation
- Python 3.7 or higher
- Windows OS (for PDF generation via COM)
- Microsoft Excel (for PDF generation on Windows)

### Docker Installation
- Docker
- For PDF generation:
  - Windows: Microsoft Excel (via win32com)
  - Linux: LibreOffice (automatically installed in container)

## Installation

### Local Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/automated-invoice-generator.git
cd automated-invoice-generator
```

2. Create and activate a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

### Docker Installation

1. Build the Docker image:
```bash
docker build -t invoice-generator .
```

2. Run the container:
```bash
# For Linux containers
docker run --rm -v ${PWD}/screenshots:/app/screenshots invoice-generator python sample_invoice.py

# For Windows containers (requires Windows host)
docker run --rm -v ${PWD}/screenshots:/app/screenshots invoice-generator python sample_invoice.py
```

For detailed Docker setup and troubleshooting, see [Docker Support](../Docker_README.md).

## Project Structure

```
automated-invoice-generator/
├── docs/
│   ├── README.md
│   └── assets/
│       └── logo.png
├── tests/
│   ├── test_docker.py
│   ├── test_export_invoice.py
│   ├── test_format_invoice.py
│   ├── test_format_styling.py
│   ├── test_generate_all.py
│   ├── test_load_orders.py
│   └── test_send_invoice.py
├── invoice.py
├── sample_invoice.py
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── verify_docker.sh
└── Docker_README.md
```

## Usage

### Input Data Format

The application expects an Excel workbook with the following sheets:

1. **Orders Sheet**
   - Required columns: CustomerID, ItemID, Qty, Price, CustomerName, Email

2. **BillTo Sheet**
   - Required columns: CustomerID, CustomerName, Email, Phone, Address, City

3. **Items Sheet**
   - Required columns: ItemID, Name, Description, UnitPrice

### Generating Invoices

1. Basic usage:
```python
from invoice import GenerateAllInvoices

# Generate invoices from your orders file
GenerateAllInvoices("path/to/orders.xlsx")
```

2. Individual invoice generation:
```python
from invoice import LoadOrders, LoadBillToData, LoadItemData, FormatInvoice, ExportInvoice

# Load data
orders = LoadOrders("path/to/orders.xlsx")
bill_to_data = LoadBillToData("path/to/orders.xlsx")
item_data = LoadItemData("path/to/orders.xlsx")

# Generate and export invoice
for order in orders:
    formatted_invoice = FormatInvoice(order)
    ExportInvoice(formatted_invoice, f"output/invoice_{order['InvoiceNumber']}", format="pdf")
```

### Platform-Specific Features

#### Windows Environment
- Uses win32com for PDF generation
- Requires Microsoft Excel to be installed on the host system
- Supports native Windows COM automation

#### Linux Environment
- Uses LibreOffice for PDF generation
- No additional host system requirements
- Fully containerized solution

## Testing

### Local Testing
Run the test suite using pytest:
```bash
pytest tests/
```

### Docker Testing
Run tests in the container:
```bash
docker run --rm invoice-generator python -m pytest tests/
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support, please:
1. Check the [Docker Support](../Docker_README.md) for Docker-related issues
2. Open an issue in the GitHub repository
3. Contact the maintainers

## Acknowledgments

- Built with [openpyxl](https://openpyxl.readthedocs.io/)
- PDF generation:
  - Windows: Microsoft Excel COM automation
  - Linux: LibreOffice
- Cross-platform testing with Docker 