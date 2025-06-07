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