
import cv2
import pandas as pd
import os
import json
from paddleocr import PaddleOCR

def simple_table_extraction(image_path, output_dir='./output'):
    os.makedirs(output_dir, exist_ok=True)

    ocr = PaddleOCR(
        use_angle_cls=True, 
        lang='en',
        use_gpu=False,
        show_log=False,
        draw_img_save_dir=None,
        vis_font_path=None
    )

    img = cv2.imread(image_path)
    if img is None:
        raise ValueError(f"Could not load image from {image_path}")
    
    try:
        result = ocr.ocr(img, cls=True)
        
        if not result or not result[0]:
            print("No text detected in the image")
            return None

        text_info = []
        for line in result[0]:
            box = line[0]
            text = line[1][0]
            confidence = line[1][1]

            center_x = sum([p[0] for p in box]) / 4
            center_y = sum([p[1] for p in box]) / 4
            
            text_info.append({
                'text': text,
                'confidence': confidence,
                'center_x': center_x,
                'center_y': center_y,
                'box': box
            })

        text_info.sort(key=lambda x: x['center_y'])
        
        rows = []
        current_row = [text_info[0]]
        row_height_threshold = 20

        for i in range(1, len(text_info)):
            if abs(text_info[i]['center_y'] - current_row[0]['center_y']) < row_height_threshold:
                current_row.append(text_info[i])
            else:
                current_row.sort(key=lambda x: x['center_x'])
                rows.append(current_row)
                current_row = [text_info[i]]

        if current_row:
            current_row.sort(key=lambda x: x['center_x'])
            rows.append(current_row)

        table_data = []
        for row in rows:
            row_data = [cell['text'] for cell in row]
            table_data.append(row_data)

        max_cols = max(len(row) for row in table_data)

        padded_data = [row + [''] * (max_cols - len(row)) for row in table_data]

        if len(padded_data) > 1:
            try:
                df = pd.DataFrame(padded_data[1:], columns=padded_data[0])
            except:
                df = pd.DataFrame(padded_data)

            csv_path = os.path.join(output_dir, "extracted_table.csv")
            df.to_csv(csv_path, index=False)
            print(f"Table saved to CSV: {csv_path}")

            txt_path = os.path.join(output_dir, "extracted_table.txt")
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                if isinstance(df, pd.DataFrame) and not df.empty:
                    if df.columns is not None and not all(pd.isna(df.columns)):
                        txt_file.write("\t".join(str(col) for col in df.columns) + "\n")

                    txt_file.write("-" * 80 + "\n")
                    
                    for _, row in df.iterrows():
                        txt_file.write("\t".join(str(cell) for cell in row) + "\n")
                else:
                    txt_file.write("No valid table data to write")
            
            print(f"Table saved to TXT: {txt_path}")

            json_path = os.path.join(output_dir, "extracted_table.json")

            json_data = {
                "table_data": {
                    "headers": df.columns.tolist() if isinstance(df, pd.DataFrame) else [],
                    "rows": df.values.tolist() if isinstance(df, pd.DataFrame) else []
                },
                "raw_data": {
                    "text_blocks": [{
                        "text": item["text"],
                        "confidence": float(item["confidence"]),
                        "position": {
                            "center_x": float(item["center_x"]),
                            "center_y": float(item["center_y"]),
                            "box": [[float(p[0]), float(p[1])] for p in item["box"]]
                        }
                    } for item in text_info]
                }
            }
            
            with open(json_path, 'w', encoding='utf-8') as json_file:
                json.dump(json_data, json_file, indent=2, ensure_ascii=False)
            
            print(f"Table saved to JSON: {json_path}")

            print("\nExtracted table content:")
            print(df)
            
            return df, json_data
        else:
            print("Not enough data to form a table")
            return None
            
    except Exception as e:
        print(f"Error: {str(e)}")
        return None

def main():
    image_path = "/home/talgotram/Repos/ioclOCR/extracted_images/page_1.png"
    output_dir = "./table_output"
    
    try:
        simple_table_extraction(image_path, output_dir)
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()