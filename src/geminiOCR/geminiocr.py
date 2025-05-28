from dotenv import load_dotenv
from pdf_to_json import main

load_dotenv()

pdf_path = "Adobe Scan 13 May 2025.pdf"

output_folder = "extracted_images"
json_output_path = "extracted_data.json"

result = main(pdf_path, output_folder, json_output_path)

print(f"Successfully processed {pdf_path}")
print(f"Extracted {len(result['pages'])} pages of content")
print(f"JSON output saved to: {json_output_path}")

if result["pages"]:
    first_page = result["pages"][0]["content"]

    if "document_type" in first_page:
        print(f"Document type: {first_page['document_type']}")

    if "key_value_pairs" in first_page and first_page["key_value_pairs"]:
        print("Found key-value pairs:")
        for key, value in first_page["key_value_pairs"].items():
            print(f"  {key}: {value}")