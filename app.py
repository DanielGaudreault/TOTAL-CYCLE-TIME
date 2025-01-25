from flask import Flask, request, render_template, send_file
import pandas as pd
from PyPDF2 import PdfReader
import os
import re

app = Flask(__name__)

# Ensure the uploads directory exists
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    # Check if files are uploaded
    if 'excel' not in request.files or 'pdf' not in request.files:
        return "Please upload both Excel and PDF files."

    excel_file = request.files['excel']
    pdf_files = request.files.getlist('pdf')

    # Save Excel file temporarily
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(excel_path)

    # Read Excel file
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        return f"Error reading Excel file: {e}"

    # Ensure the Excel file has an "Item Number" column
    if 'Item Number' not in df.columns:
        return "Excel file must have an 'Item Number' column."

    # Add "TOTAL CYCLE TIME" column if it doesn't exist
    if 'TOTAL CYCLE TIME' not in df.columns:
        df['TOTAL CYCLE TIME'] = ''

    # Process each PDF file
    for pdf_file in pdf_files:
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        pdf_file.save(pdf_path)

        # Read PDF file
        try:
            pdf_reader = PdfReader(pdf_path)
            pdf_text = ""
            for page in pdf_reader.pages:
                pdf_text += page.extract_text()
        except Exception as e:
            return f"Error reading PDF file: {e}"

        # Extract Item Number and TOTAL CYCLE TIME from PDF text
        item_number = extract_item_number(pdf_text)
        cycle_time = extract_cycle_time(pdf_text)

        if item_number and cycle_time:
            # Match item number with the "Item Number" column in Excel
            matching_row = df['Item Number'] == item_number

            if matching_row.any():
                # Update the "TOTAL CYCLE TIME" column for the matching row
                df.loc[matching_row, 'TOTAL CYCLE TIME'] = cycle_time
            else:
                print(f"No matching item number found for PDF file: {pdf_file.filename}")
        else:
            print(f"Could not extract item number or cycle time from PDF file: {pdf_file.filename}")

    # Save updated Excel file
    updated_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_excel.xlsx')
    df.to_excel(updated_excel_path, index=False)

    # Provide download link
    return send_file(updated_excel_path, as_attachment=True)

def extract_item_number(text):
    # Use regex to find the item number in the PDF text
    regex = r"Item Number\s*:\s*([\w\s-]+)"
    match = re.search(regex, text, re.IGNORECASE)
    return match.group(1).strip() if match else None

def extract_cycle_time(text):
    # Use regex to find "TOTAL CYCLE TIME" and extract the time
    regex = r"TOTAL CYCLE TIME\s*:\s*(\d+\s*HOURS?,\s*\d+\s*MINUTES?,\s*\d+\s*SECONDS?)"
    match = re.search(regex, text, re.IGNORECASE)
    return match.group(1) if match else None

if __name__ == '__main__':
    app.run(debug=True)
