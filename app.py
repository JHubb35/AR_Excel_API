from flask import Flask, send_file
import requests
import openpyxl
import shutil
import os
from openpyxl.utils import column_index_from_string

app = Flask(__name__)

# Constants
API_URL = "https://reasolllc.cetecerp.com/api/invoice?invoicedate_from=2023:01:01&preshared_token=8rtpv5gm-dywJH%7C_%5B"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, "AR_Template_API.xlsx")
OUTPUT_PATH = os.path.join(BASE_DIR , "NEW_AR_Report.xlsx")


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
    try:
        response = requests.get(API_URL)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print("API error:", e)
        return "Failed to fetch or parse API data", 500

    if not os.path.exists(TEMPLATE_PATH):
        print("Excel template not found!")
        return "Excel template not found", 500

    shutil.copy(TEMPLATE_PATH, OUTPUT_PATH)
    wb = openpyxl.load_workbook(OUTPUT_PATH)
    ws = wb.active

    start_row = 2
    for i, record in enumerate(data, start=start_row):
        for col_letter, field in COLUMN_MAP.items():
            col_index = column_index_from_string(col_letter)
            ws.cell(row=i, column=col_index, value=record.get(field))

    wb.save(OUTPUT_PATH)
    return send_file(OUTPUT_PATH, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
