from flask import Flask, send_file
import requests
import openpyxl
import shutil
import os
from openpyxl.utils import column_index_from_string

app = Flask(__name__)

# Constants
API_URL = "https://reasolllc.cetecerp.com/api/invoice?invoicedate_from=2023:01:01"
TEMPLATE_PATH = "AR_Template_API.xlsx"
OUTPUT_PATH = "NEW_AR_Report.xlsx"


COLUMN_MAP = {
    "A": "invoicenum",
    "B": "custponum",
    "C": "invoice__bc",
    "D": "invoicedate",
    "E": "ar_duedate",
    "F": "paid_on",
    "I": "terms_desc",
    "K": "invoice_tax_subtotal",
    "M": "paid",
    "O": "total_invoice_amount",
    "Q": "ar_status",
    "R": "invoicedate"  
}


@app.route("/")
def download_excel():
    # Step 1: Fetch data from API
    response = requests.get(API_URL)
    data = response.json()

    # Step 2: Copy the template
    shutil.copy(TEMPLATE_PATH, OUTPUT_PATH)

    # Step 3: Load the workbook
    wb = openpyxl.load_workbook(OUTPUT_PATH)
    ws = wb.active

    # Step 4: Write the data to correct columns
    start_row = 2
    for i, record in enumerate(data, start=start_row):
        for field, col_letter in COLUMN_MAP.items():
            col_index = column_index_from_string(col_letter)
            ws.cell(row=i, column=col_index, value=record.get(field))

    # Step 5: Save and send file
    wb.save(OUTPUT_PATH)
    return send_file(OUTPUT_PATH, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
