from flask import Flask, request, jsonify, send_file
import pandas as pd
import fitz  # PyMuPDF for PDF processing
import os
import re

app = Flask(__name__)

# Ensure the uploads directory exists
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file."""
    pdf_document = fitz.open(pdf_path)
    text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text

def extract_cycle_time(text):
    """Extract the 'TOTAL CYCLE TIME' from the text."""
    lines = text.split('\n')
    for line in lines:
        if "TOTAL CYCLE TIME" in line:
            regex = r"(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)"
            match = re.search(regex, line, re.IGNORECASE)
            return match.group(0) if match else None
    return None

def extract_part_number(text):
    """Extract the part number from the text."""
    regex = r"Part Number\s*:\s*(\w+)"
    match = re.search(regex, text, re.IGNORECASE)
    return match.group(1) if match else None

@app.route('/process-files', methods=['POST'])
def process_files():
    try:
        # Check if files are uploaded
        if 'files' not in request.files or 'excel' not in request.files:
            return jsonify({"error": "Please upload both Excel and PDF files."}), 400

        # Save the Excel file
        excel_file = request.files['excel']
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
        excel_file.save(excel_path)

        # Read the Excel file
        try:
            df = pd.read_excel(excel_path)
        except Exception as e:
            return jsonify({"error": f"Error reading Excel file: {e}"}), 400

        # Ensure the Excel file has the required columns
        if 'File Name' not in df.columns or 'Part Number' not in df.columns or 'TOTAL CYCLE TIME' not in df.columns:
            return jsonify({"error": "Excel file must contain 'File Name', 'Part Number', and 'TOTAL CYCLE TIME' columns."}), 400

        # Process each uploaded file
        for file in request.files.getlist('files'):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(file_path)

            # Extract text from the file
            if file.filename.lower().endswith('.pdf'):
                text = extract_text_from_pdf(file_path)
            else:
                with open(file_path, 'r') as f:
                    text = f.read()

            # Extract the cycle time
            cycle_time = extract_cycle_time(text)
            if cycle_time:
                # Match the file name or part number with the Excel row
                file_name = file.filename
                part_number = extract_part_number(text)  # Extract part number from the file

                # Update the matching row in the Excel file
                if file_name in df['File Name'].values:
                    df.loc[df['File Name'] == file_name, 'TOTAL CYCLE TIME'] = cycle_time
                elif part_number and part_number in df['Part Number'].values:
                    df.loc[df['Part Number'] == part_number, 'TOTAL CYCLE TIME'] = cycle_time
                else:
                    # Add new row if no match found
                    new_row = pd.DataFrame({
                        'File Name': [file_name],
                        'Part Number': [part_number or 'Unknown'],
                        'TOTAL CYCLE TIME': [cycle_time]
                    })
                    df = pd.concat([df, new_row], ignore_index=True)

        # Save the updated Excel file
        updated_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_excel.xlsx')
        df.to_excel(updated_excel_path, index=False)

        # Provide download link for the updated Excel file
        return send_file(updated_excel_path, as_attachment=True, download_name='updated_excel.xlsx')

    except Exception as e:
        return jsonify({"error": f"An error occurred: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True)
