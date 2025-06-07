import pytest
import platform
from openpyxl import Workbook
from invoice import ExportInvoice
from unittest.mock import patch, MagicMock
import sys
import types
import os

@pytest.fixture
def dummy_wb():
    # Create a minimal "Invoice" workbook
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
    # Inject fake win32com and win32com.client into sys.modules
    fake_win32com = types.ModuleType("win32com")
    fake_client = types.ModuleType("win32com.client")
    fake_client.Dispatch = MagicMock()  # Add Dispatch attribute
    setattr(fake_win32com, "client", fake_client)
    with patch.dict(sys.modules, {"win32com": fake_win32com, "win32com.client": fake_client}):
        with patch("win32com.client.Dispatch") as mock_dispatch, \
             patch("os.remove") as mock_remove:
            # Set up the mock Excel COM object
            mock_excel = MagicMock()
            mock_wb_com = MagicMock()
            mock_excel.Workbooks.Open.return_value = mock_wb_com
            mock_dispatch.return_value = mock_excel

            file_path = ExportInvoice(dummy_wb, str(out), format="pdf")
            assert file_path.endswith(".pdf")
            # Check that Dispatch was called
            mock_dispatch.assert_called_with("Excel.Application")
            # Check that remove was called for the temp xlsx
            mock_remove.assert_called()

def test_export_invalid_format(dummy_wb):
    with pytest.raises(ValueError) as exc:
        ExportInvoice(dummy_wb, "out", format="docx")
    assert "Unsupported format" in str(exc.value)
