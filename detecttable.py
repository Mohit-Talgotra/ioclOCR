import cv2
import os
from pdf2image import convert_from_path

def pdf_to_images(pdf_path, output_dir, dpi=300):
    
    os.makedirs(output_dir, exist_ok=True)
    images = convert_from_path(pdf_path, dpi=dpi)
    image_paths = []

    for i, img in enumerate(images):
        path = os.path.join(output_dir, f"page_{i+1}.png")
        img.save(path)
        image_paths.append(path)
        
    return image_paths

def detect_full_table_outline(image_path, save_debug=True):
    
    image = cv2.imread(image_path)
    orig = image.copy()
    padded = cv2.copyMakeBorder(image, 20, 20, 20, 20, cv2.BORDER_CONSTANT, value=[255, 255, 255])
    gray = cv2.cvtColor(padded, cv2.COLOR_BGR2GRAY)

    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    _, thresh = cv2.threshold(blurred, 200, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 25))
    closed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)

    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if not contours:
        return None

    largest_contour = max(contours, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(largest_contour)

    if save_debug:
        debug_image = orig.copy()
        cv2.rectangle(debug_image, (x-20, y-20), (x + w - 20, y + h - 20), (0, 0, 255), 3)
        cv2.imwrite(image_path.replace(".png", "_table_outline.png"), debug_image)

    return (x-20, y-20, w, h)


if __name__ == "__main__":
    
    pdf_path = "Adobe Scan 13 May 2025.pdf"
    output_dir = "debug_table_detection"
    image_paths = pdf_to_images(pdf_path, output_dir, dpi=400)

    for i, path in enumerate(image_paths[:2]):
        bbox = detect_full_table_outline(path)
        if bbox:
            print(f"Page {i+1} table bbox: {bbox}")
        else:
            print(f"Page {i+1}: No table detected")