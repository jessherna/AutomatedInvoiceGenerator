import os
import pytest
import logging
import sys
from invoice import LoadOrders, FormatInvoice, ExportInvoice

# Configure logging to output to stdout
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

def test_docker_environment():
    """Test if we're running in either Docker or local environment"""
    # Check for Docker environment
    is_docker = os.path.exists('/app')
    
    # Check for local environment
    is_local = os.path.exists('invoice.py')
    
    # Log the environment
    if is_docker:
        logger.info("Running in Docker environment")
    elif is_local:
        logger.info("Running in local environment")
    else:
        logger.error("Environment not properly configured")
    
    # Assert that at least one environment is valid
    assert is_docker or is_local, "Neither Docker nor local environment is properly set up"
    
    # If in Docker, verify the mount
    if is_docker:
        assert os.path.exists('/app/invoice.py'), "Docker volume mount not working"
        logger.info("Docker volume mount verified")
    
    # If local, verify the application files
    if is_local:
        assert os.path.exists('invoice.py'), "Local application files not found"
        logger.info("Local application files verified")

def test_basic_invoice_generation():
    """Test basic invoice generation functionality"""
    # Create a sample order
    sample_order = {
        "InvoiceNumber": "TEST-001",
        "InvoiceDate": "2024-03-20",
        "DueDate": "2024-04-20",
        "PO": "PO-001",
        "CompanyContact": "555-0123",
        "BillTo": {
            "CustomerName": "Test Customer",
            "Address": "123 Test St",
            "City": "Test City",
            "Phone": "555-0000",
            "Email": "test@example.com"
        },
        "ShipTo": {
            "CustomerName": "Test Customer",
            "Address": "123 Test St",
            "City": "Test City",
            "Phone": "555-0000",
            "Email": "test@example.com"
        },
        "Items": [
            {
                "Qty": 2,
                "Description": "Test Item",
                "UnitPrice": 100.00
            }
        ],
        "Terms": "Net 30"
    }

    # Test invoice formatting
    workbook = FormatInvoice(sample_order)
    assert workbook is not None, "Invoice formatting failed"

    # Create test_output directory if it doesn't exist
    os.makedirs("test_output", exist_ok=True)

    # Test export to XLSX
    output_path = "test_output/invoice_test"
    xlsx_file = ExportInvoice(workbook, output_path, format="xlsx")
    assert os.path.exists(xlsx_file), "XLSX file was not created"

    # Clean up
    try:
        os.remove(xlsx_file)
    except:
        pass 