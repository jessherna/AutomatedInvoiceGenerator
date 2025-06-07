import pytest
from invoice import FormatInvoice
from openpyxl import Workbook

@pytest.fixture
def sample_order():
    return {
        "ItemID": "ABC123",
        "Qty": 2,
        "Price": 9.99,
        "CustomerName": "Acme Corp",
        "Email": "acme@example.com"
    }

def test_format_invoice_returns_workbook(sample_order):
    wb = FormatInvoice(sample_order)
    assert isinstance(wb, Workbook)
    assert "Invoice" in wb.sheetnames

def test_format_invoice_header_and_items(sample_order):
    wb = FormatInvoice(sample_order)
    ws = wb["Invoice"]
    # Header checks
    assert ws["A1"].value == "Contoso Logistics"
    assert ws["A2"].value == sample_order["CustomerName"]
    # Column headers at row 5
    headers = [ws.cell(row=5, column=i).value for i in range(1,5)]
    assert headers == ["ItemID", "Qty", "Price", "LineTotal"]
    # Line-item at row 6
    values = [ws.cell(row=6, column=i).value for i in range(1,5)]
    expected_line_total = sample_order["Qty"] * sample_order["Price"]
    assert values == [
        sample_order["ItemID"],
        sample_order["Qty"],
        sample_order["Price"],
        expected_line_total
    ]
    # Total at row 8, column 4
    assert ws.cell(row=8, column=4).value == expected_line_total
