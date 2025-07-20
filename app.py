from flask import Flask, send_file
import requests
from openpyxl import load_workbook
from io import BytesIO
import os

app = Flask(__name__)

#  Updated API URL with token
API_URL = "https://reasolllc.cetecerp.com/api/invoice?invoicedate_from=2023:01:01&preshared_token=8rtpv5gm-dywJH%7C_%5B"

TEMPLATE_FILE = "AR_Template_API.xlsx"
SHEET_NAME = "DataSheet"  # Change if needed

COLUMN_MAP = {
    "invoicenum": "A",
    "custponum": "B",
    "invoice__bc": "C",
    "invoicedate": "D",
    "ar_duedate": "E",
    "paid_on": "F",
    "terms_desc": "I",
    "invoice_tax_subtotal": "K",
    "paid": "M",
    "total_invoice_amount": "O",
    "ar_status": "Q",
    "invoicedate": "R",
}

def fetch_data():
    response = requests.get(API_URL)
    response.raise_for_status()
    return response.json()["data"]

def generate_excel(data):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb[SHEET_NAME]
    
    for i, row in enumerate(data, start=2):
        for key, col in COLUMN_MAP.items():
            ws[f"{col}{i}"] = row.get(key, "")
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route("/")
def home():
    return "AR Excel Export API is live. Visit /download-ar to download the latest Excel report."

@app.route("/download-ar")
def download_ar():
    data = fetch_data()
    excel_file = generate_excel(data)
    return send_file(
        excel_file,
        as_attachment=True,
        download_name="AR_Report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)
