import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.drawing.image import Image
from invoice import FormatInvoice
from datetime import datetime, timedelta

@pytest.fixture
def sample_order():
    return {
        "ItemID": "ABC123",
        "Qty": 2,
        "Price": 9.99,
        "CustomerName": "Acme Corp",
        "Email": "acme@example.com",
        "ItemName": "Test Item",
        "InvoiceNumber": "INV-2024-001",
        "InvoiceDate": datetime.now(),
        "DueDate": datetime.now() + timedelta(days=30),
        "Items": [
            {
                "Qty": 2,
                "Description": "Test Item",
                "UnitPrice": 9.99
            }
        ],
        "BillTo": {
            "CustomerName": "Acme Corp",
            "Address": "123 Main St",
            "City": "Metropolis",
            "Phone": "555-0123",
            "Email": "acme@example.com"
        },
        "ShipTo": {
            "CustomerName": "Acme Corp",
            "Address": "123 Main St",
            "City": "Metropolis",
            "Phone": "555-0123",
            "Email": "acme@example.com"
        },
        "CompanyContact": "555-0123",
        "Terms": "Payment is due within 15 days."
    }

def test_invoice_contains_logo(sample_order):
    wb = FormatInvoice(sample_order)
    ws = wb["Invoice"]

    # openpyxl stores inserted images in _images
    images = ws._images
    assert any(isinstance(img, Image) for img in images), "Logo not found on Invoice sheet"

def test_invoice_uses_table_style(sample_order):
    wb = FormatInvoice(sample_order)
    ws = wb["Invoice"]

    # tables live in ws._tables
    assert ws._tables, "No Excel Table found"
    tbl = list(ws._tables.values())[0] if isinstance(ws._tables, dict) \
          else ws._tables[0]
    assert isinstance(tbl, Table), "Table object not added"
    assert tbl.tableStyleInfo.name == "None", \
           f"Expected style 'None', got {tbl.tableStyleInfo.name}"
    assert tbl.ref == "A21:F22", f"Expected table ref A21:F22, got {tbl.ref}"
