import pandas as pd
import fitz  # PyMuPDF
import os

def extract_text_from_pdf(pdf_path):
    """Extract text from a PDF file."""
    text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text()
    return text

def update_excel_with_pdf_data(excel_path, pdf_folder):
    """Update Excel sheet with data extracted from PDF files."""
    # Load the existing Excel file
    df = pd.read_excel(excel_path)

    # Iterate through all PDF files in the specified folder
    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith('.pdf'):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            pdf_text = extract_text_from_pdf(pdf_path)
            
            # Here, you can process the extracted text and update the DataFrame as needed
            # For demonstration, let's assume we append the PDF text as a new row
            new_row = {'PDF_File': pdf_file, 'Extracted_Text': pdf_text}
            df = df.append(new_row, ignore_index=True)

    # Save the updated DataFrame back to the Excel file
    df.to_excel(excel_path, index=False)

# Example usage
excel_path = 'path/to/your/excel_file.xlsx'
pdf_folder = 'path/to/your/pdf_folder'
update_excel_with_pdf_data(excel_path, pdf_folder)
