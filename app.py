from flask import Flask, request, render_template, send_file
import pandas as pd
from PyPDF2 import PdfReader
import os

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

    # Save files temporarily
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
    excel_file.save(excel_path)

    # Read Excel file
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        return f"Error reading Excel file: {e}"

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

        # Update Excel file with PDF text
        df['PDF_Text'] = pdf_text  # Add a new column with PDF text

    # Save updated Excel file
    updated_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'updated_excel.xlsx')
    df.to_excel(updated_excel_path, index=False)

    # Provide download link
    return send_file(updated_excel_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
