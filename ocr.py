import os
import json
import cv2
import pytesseract
from pdf2image import convert_from_path
import easyocr
import pix2text

def pdf_to_images(pdf_path, output_dir, dpi=300):
    """Convert PDF to high-resolution images"""
    os.makedirs(output_dir, exist_ok=True)
    images = convert_from_path(pdf_path, dpi=dpi)
    image_paths = []

    for i, img in enumerate(images):
        path = os.path.join(output_dir, f"page_{i+1}.png")
        img.save(path)
        image_paths.append(path)
        
    return image_paths

def detect_form_fields(image_path):
    """Detect form fields using contour detection and line detection"""
    image = cv2.imread(image_path)
    original = image.copy()
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    thresh = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                  cv2.THRESH_BINARY_INV, 11, 2)

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    horizontal_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
    
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    vertical_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)

    form_structure = cv2.add(horizontal_lines, vertical_lines)

    contours, _ = cv2.findContours(form_structure, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    min_area = 5000
    form_fields = []
    
    for contour in contours:
        area = cv2.contourArea(contour)
        if area > min_area:
            x, y, w, h = cv2.boundingRect(contour)
            form_fields.append((x, y, w, h))

    debug_image = original.copy()
    for x, y, w, h in form_fields:
        cv2.rectangle(debug_image, (x, y), (x+w, y+h), (0, 255, 0), 2)
    
    debug_path = os.path.join(os.path.dirname(image_path), "debug_fields.png")
    cv2.imwrite(debug_path, debug_image)

    h, w = image.shape[:2]
    
    default_fields = {
        "name": (int(0.1*w), int(0.05*h), int(0.5*w), int(0.05*h)),
        "roll_number": (int(0.1*w), int(0.1*h), int(0.5*w), int(0.05*h)),
        "answers": (int(0.1*w), int(0.15*h), int(0.8*w), int(0.6*h)),
        "signature": (int(0.7*w), int(0.8*h), int(0.2*w), int(0.1*h))
    }

    field_regions = {}

    if len(form_fields) >= 4:
        form_fields.sort(key=lambda f: f[1])

        field_regions["name"] = form_fields[0]
        field_regions["roll_number"] = form_fields[1]

        largest_field = max(form_fields[2:-1], key=lambda f: f[2]*f[3]) if len(form_fields) > 3 else form_fields[2]
        field_regions["answers"] = largest_field

        field_regions["signature"] = form_fields[-1]
    else:
        field_regions = default_fields
    
    return field_regions, original

def extract_and_enhance_fields(image, field_regions, image_path):
    """Extract and enhance field images for better OCR"""
    field_images = {}
    enhanced_dir = os.path.join(os.path.dirname(image_path), "enhanced_fields")
    os.makedirs(enhanced_dir, exist_ok=True)
    
    for field_name, (x, y, w, h) in field_regions.items():
        field_img = image[y:y+h, x:x+w]
        gray = cv2.cvtColor(field_img, cv2.COLOR_BGR2GRAY)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)

        denoised = cv2.fastNlMeansDenoising(enhanced, h=10)

        enhanced_path = os.path.join(enhanced_dir, f"{field_name}.png")
        cv2.imwrite(enhanced_path, denoised)

        field_images[field_name] = {
            "path": enhanced_path,
            "image": denoised
        }
    
    return field_images

def recognize_text_multi_engine(field_images):
    """Use multiple OCR engines and combine results"""
    reader = easyocr.Reader(['en'])
    p2t = pix2text.Pix2Text()
    
    field_results = {}
    
    for field_name, data in field_images.items():
        image_path = data["path"]
        image = data["image"]
        
        results = []
        
        # Engine 1: Tesseract (optimized for handwriting)
        try:
            custom_config = r'--oem 1 --psm 6 -c preserve_interword_spaces=1'
            if field_name == "signature":
                custom_config = r'--oem 1 --psm 13'  # Special config for signature
            
            tesseract_text = pytesseract.image_to_string(image, config=custom_config)
            results.append(tesseract_text.strip())
        except Exception as e:
            print(f"Tesseract error on {field_name}: {e}")

        # Engine 2: EasyOCR (optimized for handwriting)
        try:
            easyocr_result = reader.readtext(image_path)
            easyocr_text = " ".join([res[1] for res in easyocr_result])
            results.append(easyocr_text.strip())
        except Exception as e:
            print(f"EasyOCR error on {field_name}: {e}")

        # Engine 3: Pix2Text (optimized for handwriting)
        try:
            p2t_result = p2t.recognize(image_path)
            p2t_text = p2t_result.get('text', '')
            results.append(p2t_text.strip())
        except Exception as e:
            print(f"Pix2Text error on {field_name}: {e}")

        valid_results = [r for r in results if r and len(r) > 1]
        if valid_results:
            if field_name == "name" or field_name == "roll_number":
                field_results[field_name] = min(valid_results, key=len)
            elif field_name == "signature":
                field_results[field_name] = min(valid_results, key=len)
            else:
                field_results[field_name] = max(valid_results, key=len)
        else:
            field_results[field_name] = "[No text detected]"
    
    return field_results

def post_process_results(results):
    """Apply field-specific post-processing and validation"""
    processed = {}
    
    for field, text in results.items():
        cleaned = text.replace("\\", "").replace("|", "").replace("\"", "")

        if field == "name":
            words = cleaned.split()
            cleaned = " ".join([w.capitalize() for w in words if w])
        
        elif field == "roll_number":
            import re
            digits = re.sub(r'[^0-9\-\/]', '', cleaned)
            if digits:
                cleaned = digits
        
        elif field == "signature":
            words = cleaned.split()
            if words:
                cleaned = words[0]
                if len(words) > 1:
                    cleaned += " " + words[1]
        
        processed[field] = cleaned
    
    return processed

def run_improved_ocr_pipeline(pdf_path, output_json):
    """Run the complete improved OCR pipeline"""
    print(f"Processing PDF: {pdf_path}")

    image_dir = os.path.join(os.path.dirname(pdf_path), "extracted_images")
    image_paths = pdf_to_images(pdf_path, image_dir, dpi=400)  # Higher DPI
    
    all_results = {}
    
    for image_path in image_paths:
        page_name = os.path.basename(image_path)
        print(f"Processing page: {page_name}")

        field_regions, original_image = detect_form_fields(image_path)

        field_images = extract_and_enhance_fields(original_image, field_regions, image_path)

        raw_results = recognize_text_multi_engine(field_images)

        processed_results = post_process_results(raw_results)
        
        all_results[page_name] = processed_results

    with open(output_json, "w") as f:
        json.dump(all_results, f, indent=4)
    
    print(f"OCR results saved to: {output_json}")
    return all_results

if __name__ == "__main__":
    run_improved_ocr_pipeline("Adobe Scan 13 May 2025.pdf", "improved_ocr_results.json")