import os
import json
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import easyocr
import pix2text

def pdf_to_images(pdf_path, output_dir, dpi=300):
    
    os.makedirs(output_dir, exist_ok=True)
    images = convert_from_path(pdf_path, dpi=dpi)
    image_paths = []

    for i, img in enumerate(images):
        
        path = os.path.join(output_dir, f"page_{i+1}.png")
        img.save(path)
        image_paths.append(path)
        
    return image_paths


def detect_table(image):
    
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    horizontal_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)

    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    vertical_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)

    table_mask = cv2.add(horizontal_lines, vertical_lines)

    contours, _ = cv2.findContours(table_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if not contours:
        return None, None
    
    largest_contour = max(contours, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(largest_contour)

    table_img = image[y:y+h, x:x+w]
    
    return table_img, (x, y, w, h)


def detect_table_cells(table_img):
    
    gray = cv2.cvtColor(table_img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (table_img.shape[1]//10, 1))
    horizontal_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)

    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, table_img.shape[0]//10))
    vertical_lines = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)

    grid_mask = cv2.add(horizontal_lines, vertical_lines)

    contours, _ = cv2.findContours(grid_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    min_cell_area = 500
    
    cell_boxes = []
    
    for c in contours:
        area = cv2.contourArea(c)
        if area > min_cell_area:
            x, y, w, h = cv2.boundingRect(c)
            cell_boxes.append((x, y, w, h))

    cell_boxes = sorted(cell_boxes, key=lambda b: (b[1], b[0]))

    rows = []
    current_row = []
    row_threshold = 10
    
    for i, box in enumerate(cell_boxes):
        if i == 0:
            current_row.append(box)
            continue

        prev_box = cell_boxes[i-1]
        
        if abs(box[1] - prev_box[1]) < row_threshold:
            current_row.append(box)
        else:
            rows.append(current_row)
            current_row = [box]

    if current_row:
        rows.append(current_row)

    for r in range(len(rows)):
        rows[r] = sorted(rows[r], key=lambda b: b[0])

    return rows


def extract_cells_from_table(image, table_bbox, rows):
    
    x_offset, y_offset, _, _ = table_bbox
    cells = []

    for row in rows:
        row_cells = []
        for (x, y, w, h) in row:
            cell_img = image[y_offset + y:y_offset + y + h, x_offset + x:x_offset + x + w]
            row_cells.append(cell_img)
        cells.append(row_cells)

    return cells


def extract_and_enhance_cells(cells, base_output_dir, page_num):
    enhanced_cells = {}
    page_dir = os.path.join(base_output_dir, f"page_{page_num}")
    os.makedirs(page_dir, exist_ok=True)

    for row_i, row in enumerate(cells):
        for col_i, cell_img in enumerate(row):
            
            gray = cv2.cvtColor(cell_img, cv2.COLOR_BGR2GRAY)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(gray)
            denoised = cv2.fastNlMeansDenoising(enhanced, h=10)

            cell_path = os.path.join(page_dir, f"cell_r{row_i+1}_c{col_i+1}.png")
            cv2.imwrite(cell_path, denoised)

            enhanced_cells[(row_i, col_i)] = {
                "path": cell_path,
                "image": denoised
            }

    return enhanced_cells


def recognize_text_multi_engine_cells(enhanced_cells):
    
    reader = easyocr.Reader(['en'])
    p2t = pix2text.Pix2Text()

    results = {}

    for (r, c), data in enhanced_cells.items():
        image_path = data["path"]
        image = data["image"]

        ocr_results = []

        try:
            tesseract_text = pytesseract.image_to_string(image, config=r'--oem 1 --psm 6 -c preserve_interword_spaces=1')
            ocr_results.append(tesseract_text.strip())
        except Exception:
            pass

        try:
            easyocr_result = reader.readtext(image_path)
            easyocr_text = " ".join([res[1] for res in easyocr_result])
            ocr_results.append(easyocr_text.strip())
        except Exception:
            pass

        try:
            p2t_result = p2t.recognize(image_path)
            p2t_text = p2t_result.get('text', '')
            ocr_results.append(p2t_text.strip())
        except Exception:
            pass

        valid = [r for r in ocr_results if r and len(r) > 1]

        if valid:
            results[(r, c)] = max(valid, key=len)
        else:
            results[(r, c)] = "[No text detected]"

    return results


def run_improved_ocr_pipeline(pdf_path, output_json):

    print(f"Processing PDF: {pdf_path}")

    image_dir = os.path.join(os.path.dirname(pdf_path), "extracted_images")
    image_paths = pdf_to_images(pdf_path, image_dir, dpi=400)

    all_results = {}

    for idx, image_path in enumerate(image_paths):
        page_num = idx + 1

        if page_num > 2:
            continue

        print(f"Processing page {page_num}: {os.path.basename(image_path)}")

        image = cv2.imread(image_path)

        table_img, table_bbox = detect_table(image)

        if table_img is None:
            print(f"No table detected on page {page_num}. Skipping.")
            continue

        rows = detect_table_cells(table_img)

        if not rows:
            print(f"No table cells detected on page {page_num}. Skipping.")
            continue

        cells = extract_cells_from_table(image, table_bbox, rows)

        enhanced_cells = extract_and_enhance_cells(cells, image_dir, page_num)

        ocr_results = recognize_text_multi_engine_cells(enhanced_cells)

        page_result = []
        max_col_count = max(len(row) for row in rows)

        for r in range(len(rows)):
            row_texts = []
            for c in range(max_col_count):
                text = ocr_results.get((r, c), "")
                row_texts.append(text)
            page_result.append(row_texts)

        all_results[f"page_{page_num}"] = page_result

    with open(output_json, "w") as f:
        json.dump(all_results, f, indent=4)

    print(f"OCR results saved to: {output_json}")

    return all_results


if __name__ == "__main__":
    run_improved_ocr_pipeline("Adobe Scan 13 May 2025.pdf", "improved_ocr_results.json")