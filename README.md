# ioclOCR

**ioclOCR** is a Python-based Optical Character Recognition (OCR) tool designed to extract text from images, particularly tailored for processing documents related to the Indian Oil Corporation Limited (IOCL).

## Features

- Utilizes Tesseract OCR for accurate text extraction
- Preprocessing capabilities to enhance image quality before OCR
- Batch processing support for multiple images
- Customizable parameters to suit various document types

## Installation

### 1. Clone the repository
```bash
git clone https://github.com/Mohit-Talgotra/ioclOCR.git
cd ioclOCR
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Install Tesseract OCR

#### On Ubuntu
```bash
sudo apt-get install tesseract-ocr
```

#### On Windows
- Download and install from the official Tesseract OCR GitHub page
- Ensure the Tesseract executable is added to your system PATH.