import cv2
import pandas as pd
import os
import json
from paddleocr import PaddleOCR

def improved_table_extraction(image_path, output_dir='./output'):
    os.makedirs(output_dir, exist_ok=True)

    ocr = PaddleOCR(
        use_angle_cls=True, 
        lang='en',
        use_gpu=True,
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
            text = line[1][0].strip()
            confidence = line[1][1]

            if not text:
                continue

            x_coords = [p[0] for p in box]
            y_coords = [p[1] for p in box]
            
            min_x = min(x_coords)
            max_x = max(x_coords)
            min_y = min(y_coords)
            max_y = max(y_coords)
            
            width = max_x - min_x
            height = max_y - min_y
            
            center_x = (min_x + max_x) / 2
            center_y = (min_y + max_y) / 2
            
            text_info.append({
                'text': text,
                'confidence': confidence,
                'center_x': center_x,
                'center_y': center_y,
                'min_x': min_x,
                'max_x': max_x,
                'min_y': min_y,
                'max_y': max_y,
                'width': width,
                'height': height,
                'box': box
            })

        if not text_info:
            print("No valid text blocks detected")
            return None

        text_info.sort(key=lambda x: x['center_y'])

        rows = []
        current_row = [text_info[0]]
        
        heights = [item['height'] for item in text_info]
        avg_height = sum(heights) / len(heights)
        row_height_threshold = avg_height * 0.8  # Adjust based on text density
        
        for i in range(1, len(text_info)):
            current_block = text_info[i]
            reference_block = current_row[0]
            
            y_overlap = min(current_block['max_y'], reference_block['max_y']) - max(current_block['min_y'], reference_block['min_y'])
            
            if y_overlap > 0 or abs(current_block['center_y'] - reference_block['center_y']) < row_height_threshold:
                current_row.append(current_block)
            else:
                current_row.sort(key=lambda x: x['center_x'])
                rows.append(current_row)
                current_row = [current_block]

        if current_row:
            current_row.sort(key=lambda x: x['center_x'])
            rows.append(current_row)

        all_centers_x = [block['center_x'] for row in rows for block in row]
        
        column_centers = []
        if len(all_centers_x) > 10:
            sorted_x = sorted(all_centers_x)
            gaps = [(sorted_x[i+1] - sorted_x[i], i) for i in range(len(sorted_x)-1)]
            gaps.sort(reverse=True)
            
            num_columns = min(7, len(gaps) // 3 + 2)  # Estimate number of columns
            
            separators = sorted([sorted_x[gap[1]] for gap in gaps[:num_columns-1]])
            
            column_centers = [sorted_x[0] / 2]  # Start
            for i in range(len(separators)):
                mid_point = (separators[i] + (separators[i+1] if i+1 < len(separators) else sorted_x[-1])) / 2
                column_centers.append(mid_point)
        
        if not column_centers or len(column_centers) < 3:
            num_columns = max(len(row) for row in rows)
            
            min_x = min(block['min_x'] for row in rows for block in row)
            max_x = max(block['max_x'] for row in rows for block in row)
            
            column_width = (max_x - min_x) / num_columns
            column_centers = [min_x + column_width * (i + 0.5) for i in range(num_columns)]

        table_matrix = []
        for row in rows:
            row_data = [''] * len(column_centers)
            
            for block in row:
                distances = [abs(block['center_x'] - center) for center in column_centers]
                closest_col = distances.index(min(distances))
                
                if row_data[closest_col]:
                    row_data[closest_col] += ' ' + block['text']
                else:
                    row_data[closest_col] = block['text']
            
            table_matrix.append(row_data)

        
        headers = []
        header_row_idx = 0
        
        if table_matrix and any(cell for cell in table_matrix[0] if cell.lower() in ['s. no', 's.no', 'sno', 'sl. no', 'serial no']):
            headers = table_matrix[0]
        else:
            non_empty_cols = sum(1 for cell in table_matrix[0] if cell.strip())
            if non_empty_cols >= min(3, len(table_matrix[0])):
                headers = table_matrix[0]
            else:
                headers = [f"Column {i+1}" for i in range(len(column_centers))]
                header_row_idx = -1  # No header row to skip
        
        headers = [h.strip() or f"Column {i+1}" for i, h in enumerate(headers)]
        
        if header_row_idx >= 0:
            data_rows = table_matrix[header_row_idx+1:]
        else:
            data_rows = table_matrix
            
        max_cols = len(headers)
        normalized_rows = []
        for row in data_rows:
            if not any(cell.strip() for cell in row):
                continue
                
            if len(row) < max_cols:
                row = row + [''] * (max_cols - len(row))
            elif len(row) > max_cols:
                row = row[:max_cols]
                
            normalized_rows.append(row)

        if not normalized_rows:
            print("No valid table data extracted")
            return None

        df = pd.DataFrame(normalized_rows, columns=headers)
        
        cleaned_df = df.copy()
        
        i = 0
        while i < len(cleaned_df) - 1:
            current_row = cleaned_df.iloc[i]
            next_row = cleaned_df.iloc[i+1]
            
            current_non_empty = current_row.astype(str).str.strip().ne('').sum()
            next_non_empty = next_row.astype(str).str.strip().ne('').sum()
            
            if next_non_empty <= max_cols // 2 and current_non_empty > next_non_empty:
                for col in cleaned_df.columns:
                    if pd.notna(next_row[col]) and next_row[col].strip():
                        if pd.notna(current_row[col]) and current_row[col].strip():
                            cleaned_df.at[i, col] = f"{current_row[col]} {next_row[col]}"
                        else:
                            cleaned_df.at[i, col] = next_row[col]
                
                cleaned_df = cleaned_df.drop(i+1)
                cleaned_df = cleaned_df.reset_index(drop=True)
            else:
                i += 1
        
        if 'S. No' in cleaned_df.columns or any(col for col in cleaned_df.columns if 'no' in col.lower()):
            sno_col = next((col for col in cleaned_df.columns if 'no' in col.lower()), cleaned_df.columns[0])
            
            cleaned_df[sno_col] = pd.to_numeric(cleaned_df[sno_col], errors='coerce')
            
            mask = cleaned_df[sno_col].isna()
            if mask.any():
                valid_sns = cleaned_df.loc[~mask, sno_col].dropna()
                if not valid_sns.empty:
                    last_sn = valid_sns.iloc[-1]
                    counter = last_sn + 1
                    
                    for idx in cleaned_df.index[mask]:
                        cleaned_df.at[idx, sno_col] = counter
                        counter += 1
        
        csv_path = os.path.join(output_dir, "extracted_table.csv")
        cleaned_df.to_csv(csv_path, index=False)
        print(f"Table saved to CSV: {csv_path}")

        txt_path = os.path.join(output_dir, "extracted_table.txt")
        with open(txt_path, 'w', encoding='utf-8') as txt_file:
            txt_file.write("\t".join(cleaned_df.columns) + "\n")
            txt_file.write("-" * 80 + "\n")
            
            for _, row in cleaned_df.iterrows():
                txt_file.write("\t".join(str(cell) for cell in row) + "\n")
        
        print(f"Table saved to TXT: {txt_path}")

        json_path = os.path.join(output_dir, "extracted_table.json")
        json_data = {
            "table_data": {
                "headers": cleaned_df.columns.tolist(),
                "rows": cleaned_df.values.tolist()
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
        print(cleaned_df)
        
        return cleaned_df, json_data
    
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def main():
    image_path = "/home/talgotram/Repos/ioclOCR/output/images/page_1.jpg"
    output_dir = "./table_output"
    
    try:
        improved_table_extraction(image_path, output_dir)
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()