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

def test_send_invoice(tmp_path):
    created = []
    # Inject fake win32com and win32com.client into sys.modules
    fake_win32com = types.ModuleType("win32com")
    fake_client = types.ModuleType("win32com.client")
    fake_client.Dispatch = MagicMock()  # Add Dispatch attribute
    setattr(fake_win32com, "client", fake_client)
    with patch.dict(sys.modules, {"win32com": fake_win32com, "win32com.client": fake_client}):
        # Patch win32com.client.Dispatch to return our DummyOutlook
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
            assert mail.Attachments == [str(invoice_file)]
            assert mail.sent is True
