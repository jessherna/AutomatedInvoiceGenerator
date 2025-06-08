from invoice import LoadOrders, LoadBillToData, TransformOrder, FormatInvoice, ExportInvoice, SendInvoice
import platform

# 1) Load your orders and bill-to data with matching CustomerIDs
orders = LoadOrders("tests/data/orders_sample_with_id.xlsx")
bill_to_data = LoadBillToData("tests/data/bill_to_list_with_id.xlsx")

# 2) Transform the first order & build the invoice sheet
transformed_order = TransformOrder(orders[1], bill_to_data)
wb = FormatInvoice(transformed_order)

invoice_number = transformed_order["InvoiceNumber"]
invoice_filename_base = f"invoice_{invoice_number}"

# 3a) Save as Excel
wb.save(f"screenshots/{invoice_filename_base}.xlsx")
print(f"Saved {invoice_filename_base}.xlsx")

# 3b) Export as PDF
pdf_path = ExportInvoice(wb, f"screenshots/{invoice_filename_base}", format="pdf")
print(f"Exported PDF to {pdf_path}")

# 4) Send the invoice via email (Windows only)
if platform.system() == "Windows":
    try:
        # Send to the specified email address
        SendInvoice(
            emailAddr="jessicamariep.hernandez@gmail.com",
            filePath=pdf_path,
            cc="accounting@example.com",  # Optional CC recipient
            additional_attachments=[f"screenshots/{invoice_filename_base}.xlsx"]  # Optional additional attachments
        )
        print(f"Sent invoice to jessicamariep.hernandez@gmail.com")
    except Exception as e:
        print(f"Failed to send invoice: {str(e)}")
else:
    print("Note: Email sending is only available on Windows systems")
