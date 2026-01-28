# Image2Word

**Img2Word Converter** is a modern OCR desktop application that converts images into formatted Word documents.  
It features a dark GUI, live document preview, and supports bold headers and subheadings.

---

## Features

- Dark and modern UI using CustomTkinter
- Converts images (JPG, PNG) to Word (.docx)
- Preserves text formatting: headers, subheadings, paragraphs
- Live preview of converted text within the app
- Progress indicator during OCR processing
- Single-click conversion

---
## Deployment
Test out the Tesseract application using the link: https://huggingface.co/spaces/AhmedSenpai/ocr-converter

Test out the Gemini application using the link: https://huggingface.co/spaces/AhmedSenpai/Image2Word

---

## Running from Source

If you want to run from source:

For Tesseract
```bash
git clone https://github.com/AhmedZ2003/Image2Word-Converter.git
pip install -r requirements.txt
python ocr_tesseract_app.py
```
For Gemini
```bash
git clone https://github.com/AhmedZ2003/Image2Word-Converter.git
pip install -r requirements.txt
python ocr_gemini_app.py
```
