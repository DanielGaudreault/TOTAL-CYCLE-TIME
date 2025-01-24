from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os
from datetime import datetime
import PyPDF2  # For PDF processing

app = Flask(__name__)

# Function to extract text from a PDF file
def extract_text_from_pdf(file_path):
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    return text

# Function to calculate processing time (customize this based on your file structure)
def calculate_processing_time(file_path):
    """
    Example: Calculate processing time based on file content.
    Replace this logic with your actual calculation.
    """
    if file_path.endswith('.pdf'):
        content = extract_text_from_pdf(file_path)
    else:
        with open(file_path, 'r') as file:
            content = file.read()
    
    # Example: Calculate processing time based on content length
    processing_time = len(content)  # Replace with your actual logic
    return processing_time

@app.route('/process', methods=['POST'])
def process_files():
    # Get uploaded Excel file and directory path
    excel_file = request.files['excelFile']
    file_directory = request.form['fileDirectory']

    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Ensure the Excel file has a column named "Filename" (or adjust as needed)
    if "Filename" not in df.columns:
        return jsonify({"error": "The Excel file must contain a 'Filename' column."}), 400

    # Scan files and calculate processing time
    results = []
    for filename in df["Filename"]:
        file_path = os.path.join(file_directory, filename)
        if os.path.exists(file_path):
            processing_time = calculate_processing_time(file_path)
            results.append({"Filename": filename, "Processing Time": processing_time})
        else:
            results.append({"Filename": filename, "Processing Time": "File not found"})

    # Update the Excel file with the results
    results_df = pd.DataFrame(results)
    df = df.merge(results_df, on="Filename", how="left")

    # Save the updated Excel file
    output_file = "updated_output.xlsx"
    df.to_excel(output_file, index=False)

    return jsonify({
        "message": "Processing complete!",
        "outputFile": output_file,
    })

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_from_directory('.', filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
