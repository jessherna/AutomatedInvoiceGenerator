import platform
import pytest
from pathlib import Path
import invoice
from unittest.mock import patch, MagicMock
import sys
import types

class DummyMailItem:
    def __init__(self):
        self.To = None
        self.CC = None
        self.Subject = None
        self.Body = None
        self.Attachments = []
        self.sent = False

    def Attachments_Add(self, path):
        self.Attachments.append(path)

    def Send(self):
        self.sent = True

class DummyOutlook:
    def __init__(self, created):
        self._created = created

    def CreateItem(self, item_type):
        mail = DummyMailItem()
        self._created.append(mail)
        return mail

# Create a fake win32com module for testing
def create_fake_win32com():
    fake_win32com = types.ModuleType("win32com")
    fake_client = types.ModuleType("win32com.client")
    fake_client.Dispatch = MagicMock()
    setattr(fake_win32com, "client", fake_client)
    return fake_win32com

@pytest.fixture(autouse=True)
def mock_win32com():
    """Automatically mock win32com for all tests"""
    fake_win32com = create_fake_win32com()
    with patch.dict(sys.modules, {"win32com": fake_win32com, "win32com.client": fake_win32com.client}):
        yield

@pytest.fixture(autouse=True)
def mock_platform():
    """Automatically mock platform.system to return Windows for all tests"""
    with patch("platform.system", return_value="Windows"):
        yield

def test_send_invoice(tmp_path):
    created = []
    with patch("win32com.client.Dispatch", return_value=DummyOutlook(created)):
        # Prepare a dummy file
        invoice_file = tmp_path / "inv001.pdf"
        invoice_file.write_text("PDF-DATA")

        # Call SendInvoice
        invoice.SendInvoice("foo@bar.com", str(invoice_file))

        # There should be exactly one mail item created
        assert len(created) == 1
        mail = created[0]

        # Validate mail properties
        assert mail.To == "foo@bar.com"
        assert mail.Attachments == [str(invoice_file.absolute())]
        assert mail.sent is True
        assert "Dear Valued Customer" in mail.Body
        assert "Invoice inv001" in mail.Subject
        assert "Contoso Logistics" in mail.Body
        assert "accounting@contosologistics.com" in mail.Body
        assert "(416) 555-0123" in mail.Body

def test_send_invoice_with_cc(tmp_path):
    created = []
    with patch("win32com.client.Dispatch", return_value=DummyOutlook(created)):
        invoice_file = tmp_path / "inv002.pdf"
        invoice_file.write_text("PDF-DATA")
        
        # Call SendInvoice with CC
        invoice.SendInvoice("foo@bar.com", str(invoice_file), cc="cc@bar.com")
        
        assert len(created) == 1
        mail = created[0]
        assert mail.To == "foo@bar.com"
        assert mail.CC == "cc@bar.com"
        assert mail.Attachments == [str(invoice_file.absolute())]
        assert mail.sent is True
        assert "Invoice inv002" in mail.Subject
        assert "Payment Terms" in mail.Body
        assert "Payment is due within 15 days" in mail.Body
        assert "(416) 555-0123" in mail.Body

def test_send_invoice_failure(tmp_path):
    created = []
    with patch("win32com.client.Dispatch", return_value=DummyOutlook(created)):
        invoice_file = tmp_path / "inv003.pdf"
        invoice_file.write_text("PDF-DATA")
        
        # Create a DummyOutlook that will raise an exception when Send is called
        class FailingDummyOutlook(DummyOutlook):
            def CreateItem(self, item_type):
                mail = DummyMailItem()
                def failing_send():
                    raise Exception("Send failed")
                mail.Send = failing_send
                self._created.append(mail)
                return mail
        
        # Use the failing outlook
        with patch("win32com.client.Dispatch", return_value=FailingDummyOutlook(created)):
            # Test that the function handles the error gracefully
            with pytest.raises(Exception):
                invoice.SendInvoice("foo@bar.com", str(invoice_file))

def test_send_invoice_multiple_attachments(tmp_path):
    created = []
    with patch("win32com.client.Dispatch", return_value=DummyOutlook(created)):
        # Create multiple dummy files
        invoice_file = tmp_path / "inv004.pdf"
        additional_file = tmp_path / "additional.pdf"
        xlsx_file = tmp_path / "additional.xlsx"  # This should be ignored
        invoice_file.write_text("PDF-DATA")
        additional_file.write_text("ADDITIONAL-DATA")
        xlsx_file.write_text("XLSX-DATA")
        
        # Call SendInvoice with multiple attachments
        invoice.SendInvoice("foo@bar.com", str(invoice_file), additional_attachments=[str(additional_file), str(xlsx_file)])
        
        assert len(created) == 1
        mail = created[0]
        assert mail.To == "foo@bar.com"
        assert set(mail.Attachments) == {str(invoice_file.absolute()), str(additional_file.absolute())}  # xlsx file should be excluded
        assert mail.sent is True
        assert "Invoice Details" in mail.Body
        assert "Payment Terms" in mail.Body
        assert "(416) 555-0123" in mail.Body

def test_send_invoice_file_not_found(tmp_path):
    created = []
    with patch("win32com.client.Dispatch", return_value=DummyOutlook(created)):
        # Try to send a non-existent file
        non_existent_file = tmp_path / "nonexistent.pdf"
        
        # Test that the function raises a FileNotFoundError
        with pytest.raises(RuntimeError) as exc_info:
            invoice.SendInvoice("foo@bar.com", str(non_existent_file))
        
        assert "File not found" in str(exc_info.value)

def test_send_invoice_non_pdf_file(tmp_path):
    created = []
    with patch("win32com.client.Dispatch", return_value=DummyOutlook(created)):
        # Try to send a non-PDF file
        xlsx_file = tmp_path / "invoice.xlsx"
        xlsx_file.write_text("XLSX-DATA")
        
        # Test that the function raises a ValueError
        with pytest.raises(ValueError) as exc_info:
            invoice.SendInvoice("foo@bar.com", str(xlsx_file))
        
        assert "Only PDF files can be attached" in str(exc_info.value)
