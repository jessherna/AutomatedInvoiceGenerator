import pytest
from openpyxl import Workbook
from invoice import ExportInvoice

@pytest.fixture
def dummy_wb():
    # Create a minimal “Invoice” workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"
    ws["A1"] = "Contoso Logistics"
    return wb

def test_export_xlsx(tmp_path, dummy_wb):
    out = tmp_path / "invoice001"
    file_path = ExportInvoice(dummy_wb, str(out), format="xlsx")
    assert file_path.endswith(".xlsx")
    assert (tmp_path / "invoice001.xlsx").exists()

def test_export_pdf(tmp_path, dummy_wb):
    out = tmp_path / "invoice002"
    file_path = ExportInvoice(dummy_wb, str(out), format="pdf")
    assert file_path.endswith(".pdf")
    assert (tmp_path / "invoice002.pdf").exists()

def test_export_invalid_format(dummy_wb):
    with pytest.raises(ValueError) as exc:
        ExportInvoice(dummy_wb, "out", format="docx")
    assert "Unsupported format" in str(exc.value)
