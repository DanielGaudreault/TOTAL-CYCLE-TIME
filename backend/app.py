from flask import Flask, request, jsonify
import os
import pandas as pd
import pdfplumber

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Function to extract data from a PDF file
def extract_data_from_pdf(pdf_path):
    item_no = None
    setup_no = None
    cycle_time = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                # Example logic to extract data (customize based on your PDF structure)
                if "Item No:" in text:
                    item_no = text.split("Item No:")[1].split()[0]
                if "Setup No:" in text:
                    setup_no = text.split("Setup No:")[1].split()[0]
                if "Cycle Time:" in text:
                    cycle_time = text.split("Cycle Time:")[1].split()[0]

    return item_no, setup_no, cycle_time

# Function to update the Excel sheet
def update_excel_with_pdf_data(excel_path, pdf_folder):
    df = pd.read_excel(excel_path)

    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            item_no, setup_no, cycle_time = extract_data_from_pdf(pdf_path)

            if item_no:
                match_index = df.index[df["Item No."] == item_no].tolist()
                if match_index:
                    df.at[match_index[0], "Setup No."] = setup_no
                    df.at[match_index[0], "Cycle Time"] = cycle_time

    updated_excel_path = os.path.join(app.config["UPLOAD_FOLDER"], "updated_excel.xlsx")
    df.to_excel(updated_excel_path, index=False)
    return updated_excel_path

# Route for file upload and processing
@app.route("/upload", methods=["POST"])
def upload_files():
    if "excel" not in request.files or "pdfs" not in request.files:
        return jsonify({"error": "Missing files"}), 400

    excel_file = request.files["excel"]
    pdf_files = request.files.getlist("pdfs")

    # Save uploaded files
    excel_path = os.path.join(app.config["UPLOAD_FOLDER"], excel_file.filename)
    excel_file.save(excel_path)

    pdf_folder = os.path.join(app.config["UPLOAD_FOLDER"], "pdfs")
    if not os.path.exists(pdf_folder):
        os.makedirs(pdf_folder)

    for pdf_file in pdf_files:
        pdf_file.save(os.path.join(pdf_folder, pdf_file.filename))

    # Process files
    updated_excel_path = update_excel_with_pdf_data(excel_path, pdf_folder)

    return jsonify({"message": "Files processed successfully", "updated_excel": updated_excel_path}), 200

if __name__ == "__main__":
    app.run(debug=True)
