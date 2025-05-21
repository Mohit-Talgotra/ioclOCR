import cv2
import os

os.makedirs("table_rows", exist_ok=True)

img = cv2.imread("extracted_images/page_2.png", cv2.IMREAD_GRAYSCALE)
_, binary = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY_INV)

horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
detected_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)

contours, _ = cv2.findContours(detected_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

rows = []
for cnt in contours:
    x, y, w, h = cv2.boundingRect(cnt)
    rows.append((y, y + h))

rows = sorted(rows)

original = cv2.imread("extracted_images/page_2.png")

for idx, (y1, y2) in enumerate(rows):
    row_img = original[y1:y2, :]
    cv2.imwrite(f"table_rows/row_{idx}.png", row_img)