from openpyxl import load_workbook

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