import shutil
from pathlib import Path
import pytest
import invoice

def test_generate_all_invokes_subroutines(tmp_path, monkeypatch):
    # Arrange: copy sample orders file
    sample_src = Path("tests/data/orders_sample_with_id.xlsx")
    sample_dst = tmp_path / "orders_sample_with_id.xlsx"
    shutil.copy(sample_src, sample_dst)

    # Stub LoadOrders to return 3 "orders" with OrderID & Email
    orders_stub = [
        {
            "OrderID": "1001",
            "Email": "a@b.com",
            "CustomerID": "CUST001",
            "CustomerName": "Customer One",
            "ItemID": "ITEM001",
            "Qty": 1,
            "Price": 100.00,
            "ItemName": "Test Item 1"
        },
        {
            "OrderID": "1002",
            "Email": "c@d.com",
            "CustomerID": "CUST002",
            "CustomerName": "Customer Two",
            "ItemID": "ITEM002",
            "Qty": 2,
            "Price": 200.00,
            "ItemName": "Test Item 2"
        },
        {
            "OrderID": "1003",
            "Email": "e@f.com",
            "CustomerID": "CUST003",
            "CustomerName": "Customer Three",
            "ItemID": "ITEM003",
            "Qty": 3,
            "Price": 300.00,
            "ItemName": "Test Item 3"
        }
    ]
    monkeypatch.setattr(invoice, "LoadOrders", lambda path: orders_stub)

    # Counters for subroutine calls
    counts = {"fmt": 0, "exp": 0, "send": 0}

    # Stub FormatInvoice
    def fake_format(order):
        counts["fmt"] += 1
        return f"wb_{order.get('OrderID', '')}"
    monkeypatch.setattr(invoice, "FormatInvoice", fake_format)

    # Stub ExportInvoice
    def fake_export(wb, out_base, format="pdf"):
        counts["exp"] += 1
        return f"{out_base}.pdf"
    monkeypatch.setattr(invoice, "ExportInvoice", fake_export)

    # Stub SendInvoice
    def fake_send(email, path):
        counts["send"] += 1
    monkeypatch.setattr(invoice, "SendInvoice", fake_send)

    # Act
    invoice.GenerateAllInvoices(str(sample_dst))

    # Assert each subroutine was called exactly 3 times
    assert counts["fmt"] == 3
    assert counts["exp"] == 3
    assert counts["send"] == 3
