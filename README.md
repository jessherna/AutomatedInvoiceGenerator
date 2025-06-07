# Automated Invoice Generator

A TDD-driven VBA/Python tool to automate generating, exporting, and emailing invoices from an Excel “Orders” sheet.

## Features
- **LoadOrders**: Read `orders.xlsx` → list of dicts  
- **FormatInvoice**: Populate an invoice template with order data  
- **ExportInvoice**: Save invoices as `.xlsx` or `.pdf`  
- **SendInvoice**: Attach & send via Outlook  
- **GenerateAllInvoices**: One-click pipeline to process all orders  
- Fully covered by **pytest** tests under `tests/`

## Getting Started

### Prerequisites
- Python 3.11+  
- pip  

### Installation
```bash
git clone https://github.com/jessherna/automated-invoice-generator.git
cd automated-invoice-generator
pip install -r requirements.txt
