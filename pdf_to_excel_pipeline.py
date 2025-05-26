import os
import logging
from dotenv import load_dotenv
from pdf_to_json import main as pdf_to_json_main
from json_to_excel import main as json_to_excel_main

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def process_pdf_to_excel(pdf_path, output_folder="extracted_images", json_output=None, excel_output=None):
    if not json_output:
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        json_output = f"{pdf_name}_extracted.json"
    
    if not excel_output:
        excel_output = os.path.splitext(json_output)[0] + ".xlsx"
    
    logger.info(f"Starting complete PDF to Excel pipeline for {pdf_path}")
    logger.info(f"JSON will be saved to: {json_output}")
    logger.info(f"Excel will be saved to: {excel_output}")
    
    try:
        logger.info("Step 1: Converting PDF to structured JSON...")
        pdf_to_json_main(pdf_path, output_folder, json_output)

        logger.info("Step 2: Converting JSON to formatted Excel...")
        json_to_excel_main(json_output, excel_output)
        
        logger.info("Pipeline completed successfully!")
        logger.info(f"Results saved to: {excel_output}")
        
        return json_output, excel_output
    
    except Exception as e:
        logger.error(f"Error in processing pipeline: {e}")
        raise

if __name__ == "__main__":
    import argparse

    load_dotenv()

    parser = argparse.ArgumentParser(description="Process PDF documents to structured Excel format")
    parser.add_argument("pdf_path", help="Path to the PDF file")
    parser.add_argument("--images-folder", default="extracted_images", help="Folder to save extracted images")
    parser.add_argument("--json-output", help="Path to save the intermediate JSON output")
    parser.add_argument("--excel-output", help="Path to save the final Excel output")
    
    args = parser.parse_args()

    process_pdf_to_excel(
        args.pdf_path,
        args.images_folder,
        args.json_output,
        args.excel_output
    )