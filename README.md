```markdown
# 📄 PDF to Excel Converter

This repository offers two complementary tools to convert PDFs to structured Excel files:

1. **🔧 Excel VBA Add-in** for parsing structured JSON and writing to Excel.  
2. **🌐 Web-based OCR Pipeline** for converting scanned PDFs/images into structured JSON using OCR.

---

## 📁 Project Structure

```
.
├── addin/                   # Excel VBA Add-in with custom ribbon
│   ├── modules/             # Placeholder for future modules
│   └── ribbon/
│       ├── addin.xml        # Ribbon customization
│       └── PDF2Excel_Final.xlam # Add-in file for Excel
│
├── app/                     # Flask web app to serve the OCR frontend
│   ├── static/              # Static assets (CSS/JS)
│   ├── templates/           # HTML templates
│   ├── __init__.py
│   ├── app.py               # Main app entry
│   └── app2.py              # Alternate entry (for testing/multi-version support)
│
├── src/
│   └── geminiOCR/           # Core OCR processing logic
│       ├── json_to_excel.py
│       ├── main.py
│       ├── pdf_to_excel_pipeline.py
│       ├── pdf_to_json.py
│       └── trials/          # Experimental/test scripts
│
├── uploads/                 # Uploaded files (PDFs/images)
├── output/                  # Processed output (JSON/Excel)
├── processing/              # Intermediate steps/data
├── testing/                 # Testing scripts or sample data
├── .env                     # Environment variables
├── requirements.txt         # Python dependencies
└── README.md
```

---

## 🔧 Excel VBA Add-In

### 📌 Overview

This add-in parses structured JSON files (output from OCR pipeline) and writes them into Excel in a tabular format. Ideal for large documents and batch processing.

### ✅ Features

- Ribbon integration for one-click conversion.
- Supports nested structures: pages, tables, cells.
- Optimized for large documents (~150 pages).
- No dependencies outside Excel (VBA + [VBA-JSON](https://github.com/VBA-tools/VBA-JSON)).

### 🚀 How to Use

1. **Install the Add-In:**
   - Open Excel → File → Options → Add-ins → Manage: Excel Add-ins → Browse → Select `PDF2Excel_Final.xlam`.

2. **Enable the Custom Ribbon:**
   - Ribbon appears with a button to run the macro.

3. **Run Macro:**
   - Macro reads `data.json` and populates Excel.

> 💡 Ensure `data.json` is in the same folder as the workbook.

---

## 🌐 Web-Based OCR Pipeline (Flask + Python)

### 📌 Overview

This is the backend that extracts structured content (tables, text) from PDFs/images using OCR + layout parsing.

### ✅ Features

- PDF/Image upload via web interface.
- OCR and layout parsing using Gemini or similar pipelines.
- Outputs JSON for downstream processing (like the Excel macro).
- Modular and extensible.

### 🛠 How to Run

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Run the web app:

   ```bash
   cd app
   python app.py
   ```

3. Upload a scanned PDF/image.

4. View/download the extracted JSON and Excel files.

---

## 🔄 Flow Diagram

```
[PDF/Image] 
     ↓ (OCR + Layout Parsing)
[Structured JSON] 
     ↓ (Excel Add-in)
[Excel Table Output]
```

---

## ⚡ Performance Tips

- Excel macro disables screen updating and auto-calculation during runtime.
- Sequential writing avoids slow cell-by-cell operations.
- OCR pipeline uses intermediate caching (in `processing/` folder).

---

## 📝 Prerequisites

| Component    | Requirement                                  |
|--------------|----------------------------------------------|
| Excel        | Windows with macro support (.xlsm or .xlam)  |
| Python       | ≥ 3.7                                         |
| OCR backend  | Tesseract, EasyOCR, or Gemini-based module   |
| Browser      | For accessing Flask interface                |

---

## 🔮 Future Enhancements

- GUI in Excel for JSON file selection.
- Auto-sheet generation for multiple tables.
- Support for rotated text and merged cells.
- Role-based dashboard for uploads/results.

---

## 🧠 Credits

- **VBA Parser**: Built using native VBA and [VBA-JSON](https://github.com/VBA-tools/VBA-JSON).
- **OCR Pipeline**: Modular architecture inspired by Gemini/Pix2Text layout parsers.

---

## 📜 License

MIT License – Free to use, modify, and distribute with proper attribution.

---

## 📧 Contact

For support, bugs, or feature requests, open an issue or reach out via email/GitHub.
```