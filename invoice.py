from openpyxl import Workbook, load_workbook

def LoadOrders(path):
    wb = load_workbook(path, data_only=True)
    sheet = wb["Orders"]
    rows = list(sheet.values)
    headers = rows[0]
    data = []
    for row in rows[1:]:
        entry = {
            headers[i]: (int(row[i]) if headers[i] == "Qty"
                         else float(row[i]) if headers[i] == "Price"
                         else row[i])
            for i in range(len(headers))
        }
        data.append(entry)
    return data


def FormatInvoice(order):
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice"
    # Header
    ws["A1"] = "Contoso Logistics"
    ws["A2"] = order["CustomerName"]
    # Column headers row
    columns = ["ItemID", "Qty", "Price", "LineTotal"]
    for idx, col in enumerate(columns, start=1):
        ws.cell(row=5, column=idx, value=col)
    # Line item row
    line_total = order["Qty"] * order["Price"]
    ws.cell(row=6, column=1, value=order["ItemID"])
    ws.cell(row=6, column=2, value=order["Qty"])
    ws.cell(row=6, column=3, value=order["Price"])
    ws.cell(row=6, column=4, value=line_total)
    # Total
    ws.cell(row=8, column=4, value=line_total)
    return wb

def ExportInvoice(workbook: Workbook, output_path: str, format: str = "xlsx") -> str:
    """
    Save `workbook` to disk:
      - as .xlsx if format=="xlsx"
      - as .pdf using Excel COM if format=="pdf"
    Returns the full path of the generated file.
    """
    if format == "xlsx":
        file = output_path + ".xlsx"
        workbook.save(file)
        return file

    elif format == "pdf":
        try:
            import os
            from win32com.client import Dispatch

            # first save a temporary XLSX to open in COM
            tmp_xlsx = output_path + "_tmp.xlsx"
            workbook.save(tmp_xlsx)

            excel = Dispatch("Excel.Application")
            excel.Visible = False
            wb_com = excel.Workbooks.Open(os.path.abspath(tmp_xlsx))
            wb_com.ExportAsFixedFormat(0, os.path.abspath(output_path + ".pdf"))
            wb_com.Close(False)
            excel.Quit()

            os.remove(tmp_xlsx)
            return output_path + ".pdf"
        except ImportError:
            raise RuntimeError("PDF export requires pywin32 and Windows COM")
        except Exception as e:
            raise RuntimeError(f"PDF export failed: {e}")

    else:
        raise ValueError(f"Unsupported format: {format}")
    
def SendInvoice(emailAddr: str, filePath: str) -> None:
    """
    Open Outlook, create a mail item, attach the file at filePath, and send to emailAddr.
    """
    try:
        from win32com.client import Dispatch
    except ImportError:
        raise RuntimeError("SendInvoice requires pywin32 and Windows COM")

    outlook = Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = emailAddr
    mail.Subject = "Your Invoice from Contoso Logistics"
    mail.Body = "Please find your invoice attached."

    # Attach file (fallback to mail.Attachments_Add if Attachments.Add doesn't exist)
    try:
        mail.Attachments.Add(filePath)
    except AttributeError:
        # e.g. DummyMailItem in tests uses Attachments_Add
        if hasattr(mail, "Attachments_Add"):
            mail.Attachments_Add(filePath)
        else:
            raise

    mail.Send()

def GenerateAllInvoices(path_to_orders: str) -> None:
    """
    Full pipeline: load all orders, format each invoice, export to PDF,
    and send via Outlook.
    """
    orders = LoadOrders(path_to_orders)
    for order in orders:
        # 1) Format
        wb = FormatInvoice(order)

        # 2) Export (use OrderID as base filename)
        base = f"invoice_{order.get('OrderID', '')}"
        pdf_path = ExportInvoice(wb, base, format="pdf")

        # 3) Send
        SendInvoice(order.get("Email", ""), pdf_path)
