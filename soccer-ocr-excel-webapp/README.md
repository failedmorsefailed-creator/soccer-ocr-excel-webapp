# Soccer OCR to Excel Web App

A Flask web application that lets you upload an Excel template and a soccer schedule image.  
It runs OCR to extract match info from the image and fills the Excel template, preserving formulas and formatting.

## Usage

1. Upload your Excel template (.xlsx).
2. Upload a schedule image (jpg/png).
3. Preview the extracted data.
4. Confirm and download the filled Excel sheet.

## Deploy

- Designed for easy deployment to [Render.com](https://render.com/) and mobile-friendly use.
- Requires Tesseract OCR (installed via Aptfile).

---

**Files:**
- `app.py` — Main Flask app
- `requirements.txt` — Python requirements
- `Aptfile` — OS packages for Render (Tesseract)
- `Procfile` — Gunicorn start command
- `templates/` — HTML templates
