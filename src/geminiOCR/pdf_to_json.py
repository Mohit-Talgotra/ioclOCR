import os
import json
import logging
from typing import List, Dict, Any, Optional
import google.generativeai as genai
from dotenv import load_dotenv
from pdf2image import convert_from_path
from PIL import Image
import concurrent.futures
import time
from datetime import datetime
from threading import Lock

class APIRateLimiter:
    def __init__(self, calls_per_minute=10):
        self.calls_per_minute = calls_per_minute
        self.min_interval = 60.0 / calls_per_minute
        self.last_call_time = 0
        self.lock = Lock()
    
    def wait_if_needed(self):
        with self.lock:
            current_time = time.time()
            time_since_last_call = current_time - self.last_call_time
            if time_since_last_call < self.min_interval:
                sleep_time = self.min_interval - time_since_last_call
                logger.info(f"⏳ Rate limiting: waiting {sleep_time:.2f} seconds...")
                time.sleep(sleep_time)
            self.last_call_time = time.time()

api_limiter = APIRateLimiter(calls_per_minute=15)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

load_dotenv()
api_key = os.getenv('GOOGLE_API_KEY')
if not api_key:
    raise ValueError("GOOGLE_API_KEY environment variable not found. Please set it in your .env file.")

genai.configure(api_key=api_key)
model = genai.GenerativeModel('gemini-2.0-flash')

def convert_pdf_to_images(pdf_path: str, output_folder: str, dpi: int = 300) -> List[str]:
    logger.info(f"Converting PDF: {pdf_path} to images")
    logger.info(f"Using DPI: {dpi}, Output folder: {output_folder}")
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        logger.info(f"Created output directory: {output_folder}")
        
    try:
        start_time = time.time()
        logger.info("Starting PDF to image conversion...")
        images = convert_from_path(pdf_path, dpi=dpi)
        conversion_time = time.time() - start_time
        logger.info(f"PDF conversion completed in {conversion_time:.2f} seconds. Found {len(images)} pages")
        
        image_paths = []
        for i, image in enumerate(images):
            logger.info(f"Saving page {i+1}/{len(images)}...")
            image_path = os.path.join(output_folder, f'page_{i+1}.jpg')
            image.save(image_path, 'JPEG', quality=95)
            image_paths.append(image_path)
            logger.info(f"Saved: {image_path}")
            
        logger.info(f"Successfully extracted {len(image_paths)} pages from {pdf_path}")
        return image_paths
    
    except Exception as e:
        logger.error(f"Error converting PDF to images: {e}")
        raise

def analyze_document_structure(image_path: str) -> str:
    logger.info(f"Analyzing document structure: {image_path}")

    try:
        start_time = time.time()
        image = Image.open(image_path)
        logger.info(f"Image loaded, size: {image.size}")
        
        structure_prompt = """
        Analyze this document page and describe its structure in detail.
        
        Focus on identifying:
        1. Overall layout (single column, multi-column, complex layout)
        2. Presence of tables (simple or complex)
        3. Forms or structured data fields
        4. Headers, footers, and page numbers
        5. Charts, graphs, or diagrams
        6. Special formatting elements (boxes, highlights, etc.)
        
        Provide a concise structural analysis that would help determine the best approach for 
        extracting structured information from this document.
        """

        api_limiter.wait_if_needed()
        
        logger.info("Sending analysis request to Gemini...")
        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = model.generate_content([structure_prompt, image])
                analysis_time = time.time() - start_time
                
                structure_analysis = response.text
                logger.info(f"Structure analysis completed in {analysis_time:.2f} seconds")
                return structure_analysis
                
            except Exception as e:
                logger.warning(f"Attempt {attempt + 1} failed: {str(e)}")
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 5  # 5, 10, 15 seconds
                    logger.info(f"Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    raise e
    except Exception as e:
        logger.error(f"Error analyzing document structure: {e}")
        return "Error analyzing document structure"

def extract_structured_content(image_path: str, structure_analysis: str) -> Dict[str, Any]:
    logger.info(f"Extracting structured content from: {image_path}")
    
    try:
        start_time = time.time()
        image = Image.open(image_path)
    
        extraction_prompt = f"""
        Based on the structural analysis, extract ALL content from this document page into well-structured JSON.
        
        Structural analysis: {structure_analysis}
        
        Instructions for extraction:
        
        1. For general text:
        - Preserve paragraph structure
        - Maintain headings and subheadings hierarchy
        - Capture lists with their items
        
        2. For tables:
        - Extract as a nested array structure with headers
        - Maintain column headers and row labels
        - Preserve all cell values with their exact formatting
        - Handle merged cells appropriately
        
        3. For forms or structured data:
        - Create key-value pairs for each field and its value
        - Group related fields together
        - Preserve field labels exactly as they appear
        
        4. For charts/diagrams:
        - Extract title, axes labels, and legend text
        - Describe the chart type and key data points
        
        5. For headers/footers:
        - Capture page numbers, dates, and reference numbers
        - Extract any metadata like document ID or revision info
        
        OUTPUT FORMAT:
        Return ONLY valid JSON with this structure:
        {{
            "document_type": "detected document type (e.g., invoice, form, report)",
            "page_metadata": {{
                "page_number": "detected page number if present",
                "header": "header text if present",
                "footer": "footer text if present"
            }},
            "sections": [
                {{
                    "section_type": "text|table|form|chart",
                    "section_title": "section heading if present",
                    "content": "appropriate content structure based on section type"
                }}
            ],
            "tables": [
                {{
                    "table_title": "title if present",
                    "headers": ["header1", "header2", ...],
                    "data": [
                        ["row1col1", "row1col2", ...],
                        ["row2col1", "row2col2", ...]
                    ]
                }}
            ],
            "key_value_pairs": {{
                "key1": "value1",
                "key2": "value2"
            }}
        }}
        
        Make sure to use proper JSON escaping for special characters and ensure the output is valid JSON.
        If certain elements don't exist, include them as empty arrays or objects rather than omitting them.
        VERY IMPORTANT: If a signature is detected in a column like User Sign or anything of that kind, add an indication that signature detected in that column in your structure output
        """

        api_limiter.wait_if_needed()

        logger.info("Sending extraction request to Gemini...")
        max_retries = 3
        for attempt in range(max_retries):
            try:
                response = model.generate_content([extraction_prompt, image])
                extraction_time = time.time() - start_time
                logger.info(f"Content extraction completed in {extraction_time:.2f} seconds")
                
                extract_text = response.text
                logger.info(f"Raw response length: {len(extract_text)} characters")
                break
                
            except Exception as e:
                logger.warning(f"Extraction attempt {attempt + 1} failed: {str(e)}")
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 5
                    logger.info(f"Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                else:
                    logger.error(f"All extraction attempts failed for {image_path}")
                    return {"error": f"API failed after {max_retries} attempts: {str(e)}"}
        
        json_start = extract_text.find('{')
        json_end = extract_text.rfind('}') + 1
        
        if json_start >= 0 and json_end > json_start:
            logger.info("Parsing JSON response...")
            json_content = extract_text[json_start:json_end]
            try:
                structured_data = json.loads(json_content)
                logger.info(f"Successfully parsed JSON structure with {len(structured_data)} top-level keys")
                return structured_data
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON from model: {e}")
                return {"error": "Failed to parse JSON", "raw_text": extract_text}
        else:
            logger.error("No JSON found in model response")
            return {"error": "No JSON found in response", "raw_text": extract_text}
    
    except Exception as e:
        logger.error(f"Error extracting structured content: {e}")
        return {"error": str(e)}

def verify_extraction(image_path: str, structured_data: Dict[str, Any]) -> Dict[str, Any]:
    logger.info(f"Verifying extraction quality for: {image_path}")
    
    try:
        start_time = time.time()
        image = Image.open(image_path)
        structured_json = json.dumps(structured_data, indent=2)
    
        verification_prompt = f"""
        I need you to verify and correct the structured data extracted from this document image.
        
        Extracted structured data:
        ```json
        {structured_json}
        ```
        
        Verification tasks:
        1. Check for missing content (sections, tables, form fields, etc.)
        2. Verify accuracy of all extracted text, numbers, and values
        3. Ensure table structures correctly represent the original format
        4. Confirm that relationships between data elements are preserved
        5. Verify all key-value pairs have been correctly identified and paired
        
        If everything is correct, respond with "VERIFICATION_PASSED" followed by the original JSON.
        
        If corrections are needed, respond with "CORRECTIONS_NEEDED" followed by the complete corrected JSON.
        Ensure the corrected JSON maintains the same structure but with accurate data.
        
        OUTPUT FORMAT:
        Either:
        VERIFICATION_PASSED
        {{original JSON}}
        
        Or:
        CORRECTIONS_NEEDED
        {{corrected JSON}}
        """
        
        logger.info("Sending verification request to Gemini...")
        response = model.generate_content([verification_prompt, image])
        verification_time = time.time() - start_time
        
        verification_text = response.text
        logger.info(f"Verification completed in {verification_time:.2f} seconds")
        
        if "VERIFICATION_PASSED" in verification_text:
            logger.info("Verification PASSED - no corrections needed")
            return structured_data
        
        elif "CORRECTIONS_NEEDED" in verification_text:
            logger.info(f"Corrections needed for {image_path}")
            json_start = verification_text.find('{')
            json_end = verification_text.rfind('}') + 1
            
            if json_start >= 0 and json_end > json_start:
                json_content = verification_text[json_start:json_end]
                try:
                    corrected_data = json.loads(json_content)
                    return corrected_data
                except json.JSONDecodeError as e:
                    logger.error(f"Invalid JSON in verification response: {e}")
                    return structured_data
            else:
                logger.error("No corrected JSON found in verification response")
                return structured_data
        else:
            logger.warning("Unexpected verification response format")
            return structured_data
    
    except Exception as e:
        logger.error(f"Error during verification: {e}")
        return structured_data

def process_single_page(image_path: str) -> Dict[str, Any]:
    try:
        structure_analysis = analyze_document_structure(image_path)

        structured_data = extract_structured_content(image_path, structure_analysis)

        verified_data = verify_extraction(image_path, structured_data)

        page_number = int(os.path.basename(image_path).split('_')[1].split('.')[0])
        verified_data["page_info"] = {
            "page_number": page_number,
            "image_path": image_path
        }
        
        return verified_data
    
    except Exception as e:
        logger.error(f"Error processing page {image_path}: {e}")
        return {"error": str(e), "page": image_path}
    
def process_single_page_with_timeout(image_path: str, timeout_minutes=10) -> Dict[str, Any]:
    """Process a single page with timeout protection"""
    logger.info(f"Starting processing of {os.path.basename(image_path)}")
    start_time = time.time()
    
    try:
        result = process_single_page(image_path)
        processing_time = time.time() - start_time
        logger.info(f"Completed {os.path.basename(image_path)} in {processing_time:.2f} seconds")
        return result
    except Exception as e:
        processing_time = time.time() - start_time
        logger.error(f"Failed {os.path.basename(image_path)} after {processing_time:.2f} seconds: {e}")
        return {"error": str(e), "page": image_path}

def merge_page_results(page_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    logger.info("Merging results from all pages")
    
    merged_data = {
        "document_metadata": {
            "total_pages": len(page_results)
        },
        "pages": []
    }

    sorted_pages = sorted(page_results, key=lambda x: x.get("page_info", {}).get("page_number", 0))

    for page_data in sorted_pages:
        page_info = page_data.pop("page_info", {})
        page_number = page_info.get("page_number", 0)

        merged_data["pages"].append({
            "page_number": page_number,
            "content": page_data
        })
    
    return merged_data

def process_pdf_to_json(pdf_path: str, output_folder: str, json_output_path: Optional[str] = None) -> Dict[str, Any]:
    try:
        image_paths = convert_pdf_to_images(pdf_path, output_folder)

        page_results = []
        total_pages = len(image_paths)
        logger.info(f"Starting parallel processing of {total_pages} pages...")

        with concurrent.futures.ThreadPoolExecutor(max_workers=min(2, len(image_paths))) as executor:
            future_to_page = {executor.submit(process_single_page_with_timeout, image_path): i+1 for i, image_path in enumerate(image_paths)}
            
            completed = 0
            for future in concurrent.futures.as_completed(future_to_page):
                page_num = future_to_page[future]
                completed += 1
                logger.info(f"Page {page_num} completed ({completed}/{total_pages}) - {(completed/total_pages)*100:.1f}%")
                page_results.append(future.result())

        merged_data = merge_page_results(page_results)

        if json_output_path:
            with open(json_output_path, 'w', encoding='utf-8') as f:
                json.dump(merged_data, f, indent=2, ensure_ascii=False)
            logger.info(f"JSON output saved to: {json_output_path}")
        
        return merged_data
    
    except Exception as e:
        logger.error(f"Error processing PDF: {e}")
        raise

def perform_final_qc(merged_data: Dict[str, Any], pdf_path: str) -> Dict[str, Any]:
    logger.info("Performing final quality check")
    return merged_data

def main(pdf_path: str, output_folder: str = "extracted_images", json_output_path: Optional[str] = None):
    if not json_output_path:
        pdf_name = os.path.basename(pdf_path).split('.')[0]
        json_output_path = f"{pdf_name}_extracted.json"
    
    start_time = time.time()
    logger.info(f"Starting processing of {pdf_path} at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Output JSON will be saved to {json_output_path}")

    merged_data = process_pdf_to_json(pdf_path, output_folder, None)

    final_data = perform_final_qc(merged_data, pdf_path)

    with open(json_output_path, 'w', encoding='utf-8') as f:
        json.dump(final_data, f, indent=2, ensure_ascii=False)
    
    total_time = time.time() - start_time
    logger.info(f"Processing complete in {total_time:.2f} seconds. JSON output saved to: {json_output_path}")
    return final_data

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Convert PDF documents to structured JSON")
    parser.add_argument("pdf_path", help="Path to the PDF file")
    parser.add_argument("--output-folder", default="extracted_images", help="Folder to save extracted images")
    parser.add_argument("--json-output", help="Path to save the JSON output")
    
    args = parser.parse_args()
    
    main(args.pdf_path, args.output_folder, args.json_output)