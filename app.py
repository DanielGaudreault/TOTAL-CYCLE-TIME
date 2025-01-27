from flask import Flask, request, jsonify
import pandas as pd
import fitz
import os

app = Flask(__name__)

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text

# Function to extract cycle time from text
def extract_cycle_time(text):
    lines = text.split('\n')
    for line in lines:
        if "TOTAL CYCLE TIME" in line:
            regex = r"(\d+ HOURS?, \d+ MINUTES?, \d+ SECONDS?)"
            match = re.search(regex, line, re.IGNORECASE)
            return match.group(0) if match else None
    return None

@app.route('/process-files', methods=['POST'])
def process_files():
    upload_dir = 'uploads'
    os.makedirs(upload_dir, exist_ok=True)

    excel_data = pd.DataFrame()
    if 'excel-files' in request.files:
        for excel_file in request.files.getlist('excel-files'):
            file_path = os.path.join(upload_dir, excel_file.filename)
            excel_file.save(file_path)
            df = pd.read_excel(file_path)
            excel_data = excel_data.append(df, ignore_index=True)

    if 'pdf-files' in request.files:
        for pdf_file in request.files.getlist('pdf-files'):
            file_path = os.path.join(upload_dir, pdf_file.filename)
            pdf_file.save(file_path)
            pdf_text = extract_text_from_pdf(file_path)
            cycle_time = extract_cycle_time(pdf_text)
            new_row = {'File_Name': pdf_file.filename, 'Extracted_Cycle_Time': cycle_time}
            excel_data = excel_data.append(new_row, ignore_index=True)

    output_excel = os.path.join(upload_dir, 'updated_excel.xlsx')
    excel_data.to_excel(output_excel, index=False)

    return jsonify({'message': 'Files processed successfully!', 'output_excel': output_excel})

if __name__ == '__main__':
    app.run(debug=True)
