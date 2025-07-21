from flask import Flask, send_file, jsonify
import requests
import shutil
import os
import openpyxl
from openpyxl.utils import column_index_from_string
import uuid

app = Flask(__name__)

# Get API URL from environment variable
API_URL = os.environ.get("AR_API_URL")

# Define Excel template file
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILENAME = "AR_Template_API.xlsx"
TEMPLATE_PATH = os.path.join(BASE_DIR, TEMPLATE_FILENAME)

# Define mapping of Excel columns to API fields
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
def home():
    return "Welcome to the AR API Excel generator!"

@app.route("/health")
def health():
    return "OK", 200

@app.route("/download-ar", methods=["GET"])
def download_excel():
    if not API_URL:
        return jsonify({"error": "API URL is not configured"}), 500

    try:
        # Fetch data from API
        response = requests.get(API_URL)
        response.raise_for_status()
        data = response.json()
    except Exception as e:
        print("API error:", e)
        return jsonify({"error": "Failed to fetch or parse API data"}), 500

    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({"error": f"Excel template '{TEMPLATE_FILENAME}' not found"}), 500

    try:
        # Create a unique output filename
        output_filename = f"AR_Report_{uuid.uuid4().hex[:8]}.xlsx"
        output_path = os.path.join(BASE_DIR, output_filename)

        # Copy the template and write data
        shutil.copy(TEMPLATE_PATH, output_path)
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active

        start_row = 2
        for i, record in enumerate(data, start=start_row):
            for col_letter, field in COLUMN_MAP.items():
                col_index = column_index_from_string(col_letter)
                ws.cell(row=i, column=col_index, value=record.get(field))

        wb.save(output_path)
        return send_file(output_path, as_attachment=True)

    except Exception as e:
        print("Excel generation error:", e)
        return jsonify({"error": "Failed to generate Excel file"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))  # For Railway or local
    app.run(host="0.0.0.0", port=port)
