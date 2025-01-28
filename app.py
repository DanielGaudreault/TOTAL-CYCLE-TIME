import os
import pandas as pd
import pdfplumber  # For PDF extraction

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
    # Load the Excel file
    df = pd.read_excel(excel_path)

    # Loop through the PDF folder
    for pdf_file in os.listdir(pdf_folder):
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, pdf_file)
            item_no, setup_no, cycle_time = extract_data_from_pdf(pdf_path)

            if item_no:
                # Find the row in the DataFrame that matches the item number
                match_index = df.index[df["Item No."] == item_no].tolist()
                if match_index:
                    # Update the setup number and cycle time
                    df.at[match_index[0], "Setup No."] = setup_no
                    df.at[match_index[0], "Cycle Time"] = cycle_time

    # Save the updated DataFrame back to the Excel file
    df.to_excel(excel_path, index=False)

# Main execution
if __name__ == "__main__":
    excel_file_path = "path/to/your/excel_file.xlsx"  # Replace with your Excel file path
    pdf_folder_path = "path/to/your/pdf_folder"  # Replace with your PDF folder path

    update_excel_with_pdf_data(excel_file_path, pdf_folder_path)
    print("Excel file updated successfully!")
