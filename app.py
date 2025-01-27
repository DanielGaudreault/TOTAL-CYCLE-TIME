from flask import Flask, request, jsonify
import pandas as pd
import fitz
import os
import re

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
        print("Upload directory created")

        excel_data = None
        results = []

        if 'files' in request.files:
            for file in request.files.getlist('files'):
                file_path = os.path.join(upload_dir, file.filename)
                file.save(file_path)
                print(f"Saved file: {file_path}")
                
                if file.filename.endswith(('.xlsx', '.xls')):
                    excel_data = pd.read_excel(file_path)
                elif file.filename.endswith('.pdf'):
                    pdf_text = extract_text_from_pdf(file_path)
                    cycle_time = extract_cycle_time(pdf_text)
                    print(f"Extracted cycle time: {cycle_time}")
                    results.append({'File_Name': file.filename, 'Cycle_Time': cycle_time, 'Setup_Number': os.path.splitext(file.filename)[0]})

        if excel_data is not None and results:
            for result in results:
                # Find matching row by file name or setup number (assuming setup number is part of the filename)
                match_row = excel_data[excel_data.iloc[:, 0].astype(str).str.lower().str.contains(result['Setup_Number'].lower(), case=False)]
                
                if not match_row.empty:
                    # Update existing row
                    match_row_index = match_row.index[0]
                    excel_data.loc[match_row_index, 'C'] = result['Setup_Number']  # Assuming 'C' is the third column
                    excel_data.loc[match_row_index, 'D'] = result['Cycle_Time'] if result['Cycle_Time'] else 'Not Found'
                else:
                    # Add new row if no match found
                    new_row = pd.DataFrame({
                        excel_data.columns[0]: [result['Setup_Number']],  # 1st column for setup number
                        'C': [result['Setup_Number']],  # 3rd column for setup number if not found
                        'D': [result['Cycle_Time'] if result['Cycle_Time'] else 'Not Found']
                    })
                    excel_data = pd.concat([excel_data, new_row], ignore_index=True)

            output_excel = os.path.join(upload_dir, 'updated_excel.xlsx')
            excel_data.to_excel(output_excel, index=False)
            print(f"Updated Excel file saved as: {output_excel}")

            return jsonify({'message': 'Files processed successfully!', 'output_excel': output_excel})
        else:
            return jsonify({'message': 'No valid Excel file or PDF files were processed.'})
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'message': 'Error processing files.', 'error': str(e)})

if __name__ == '__main__':
    app.run(debug=True)
