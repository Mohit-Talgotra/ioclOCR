# PDF to Excel Converter (VBA)

This project is a *VBA-based automation script* that parses structured data from a JSON (converted from PDF) and writes it into Excel in an organized format.

⚠ *Note:* This project assumes that the PDF has already been converted into a structured JSON format using external tools (e.g., OCR + layout parser).

---

## 📁 Project Structure

- pdf2excel.bas: Core VBA module that reads and parses JSON into Excel.
- data.json: (Expected) Input JSON file containing extracted data from the PDF.
- Workbook.xlsm: Excel macro-enabled workbook that runs the conversion script.

---

## ✅ Features

- Reads structured JSON data exported from scanned PDFs.
- Handles nested objects like pages, tables, and cells.
- Writes table content into Excel in row-wise format.
- Designed for batch or large document processing.
- Optimized for performance with screen updating/calculation toggles.

---

## 🛠 How It Works

1. Use external tools to:
   - Convert the PDF to image/text.
   - Run OCR and layout detection.
   - Export the results as JSON (structured with pages and tables).

2. Open Workbook.xlsm in Excel.

3. Import the pdf2excel.bas module into the VBA editor:
   - Press ALT + F11 → File → Import File → Select pdf2excel.bas.

4. Run the macro:
   - Press ALT + F8 → Select ParsePDFJsonToExcel → Click Run.

---

## ⚡ Performance Optimizations

- Screen updating and Excel recalculation are temporarily disabled for faster execution.
- Avoids repetitive cell writes by using sequential indexing.
- Designed for ~150-page JSON files (batch processing supported).

---

## 📝 Prerequisites

- Windows Excel with macros enabled (.xlsm support).
- A structured data.json file placed in the same folder as the workbook.
- JSON parser module included (e.g., [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)).

---

## 🔄 Future Improvements

- Add GUI for selecting JSON file.
- Add support for image extraction.
- Auto-detect and create new sheets per table or section.
- Handle edge cases like merged cells or rotated text.

---

## 🧠 Credits

- Developed using native VBA and Excel features.
- JSON parsing powered by the open-source [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) library.

---

## 📜 License

MIT License – free to use, modify, and distribute with attribution.

---

## 📧 Contact

For questions, suggestions, or contributions, feel free to reach out via GitHub Issues or email.
