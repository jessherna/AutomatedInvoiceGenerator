import pytest
from pathlib import Path
import invoice

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

def test_send_invoice(monkeypatch, tmp_path):
    created = []

    # Monkey-patch Dispatch to return our DummyOutlook
    def fake_dispatch(prog_id):
        return DummyOutlook(created)

    monkeypatch.setattr("win32com.client.Dispatch", fake_dispatch)

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
