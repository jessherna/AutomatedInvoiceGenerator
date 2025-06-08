from invoice import LoadOrders, LoadBillToData, TransformOrder, FormatInvoice, ExportInvoice

# 1) Load your orders and bill-to data with matching CustomerIDs
orders = LoadOrders("tests/data/orders_sample_with_id.xlsx")
bill_to_data = LoadBillToData("tests/data/bill_to_list_with_id.xlsx")

# 2) Transform the first order & build the invoice sheet
transformed_order = TransformOrder(orders[1], bill_to_data)
wb = FormatInvoice(transformed_order)

# 3a) Save as Excel
wb.save("screenshots/sample-invoice.xlsx")
print("Saved sample-invoice.xlsx")

# 3b) (Optional) Export as PDF
pdf_path = ExportInvoice(wb, "screenshots/sample-invoice", format="pdf")
print(f"Exported PDF to {pdf_path}")
