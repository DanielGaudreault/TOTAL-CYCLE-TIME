from flask import Flask, request, jsonify
import pandas as pd
import fitz
import os
import re  # Added import for regular expressions

app = Flask(__name__)

def extract_text_from_pdf(pdf_path):
    pdf_document = fitz.open(pdf_path)
    text = ""
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text

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
    try:
        upload_dir = 'uploads'
        os.makedirs(upload_dir, exist_ok=True)

        excel_data = pd.DataFrame()
        if 'files' in request.files:
            for file in request.files.getlist('files'):
                file_path = os.path.join(upload_dir, file.filename)
                file.save(file_path)
                print(f"Saved file: {file_path}")
                
                if file.filename.endswith('.xlsx'):
                    df = pd.read_excel(file_path)
                    excel_data = excel_data.append(df, ignore_index=True)
                elif file.filename.endswith('.pdf'):
                    pdf_text = extract_text_from_pdf(file_path)
                    cycle_time = extract_cycle_time(pdf_text)
                    print(f"Extracted cycle time: {cycle_time}")
                    new_row = {'File_Name': file.filename, 'Extracted_Cycle_Time': cycle_time}
                    excel_data = excel_data.append(new_row, ignore_index=True)
                else:
                    print(f"Unsupported file type: {file.filename}")

        output_excel = os.path.join(upload_dir, 'updated_excel.xlsx')
        excel_data.to_excel(output_excel, index=False)
        print(f"Updated Excel file saved as: {output_excel}")

        return jsonify({'message': 'Files processed successfully!', 'output_excel': output_excel})
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'message': 'Error processing files.', 'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True)
