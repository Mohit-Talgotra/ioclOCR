from pdf_to_excel_pipeline import process_pdf_to_excel

json_path, excel_path = process_pdf_to_excel(
    pdf_path="/home/talgotram/Repos/ioclOCR/uploads/Adobe Scan 13 May 2025.pdf",
    output_folder="processing/images",
    json_output="data.json",
    excel_output="results.xlsx"
)