import pdfplumber
from pdf2image import convert_from_path
import pytesseract

def extract_text_from_pdf(file_path, poppler_path):
    with pdfplumber.open(file_path) as pdf:
        text = "".join(page.extract_text() or "" for page in pdf.pages)

    try:
        images = convert_from_path(file_path, dpi=300, poppler_path=poppler_path)
        image_text = "".join(pytesseract.image_to_string(image) for image in images)
    except Exception as e:
        print(f"Error processing images: {e}")
        image_text = ""

    return text + "\n" + image_text