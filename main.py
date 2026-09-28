import os
import pdf2image
import pytesseract
import pandas as pd

# ==============================================================================
# CONFIGURATION FOR VS CODE
# ==============================================================================
# 1. Windows Tesseract Path (Uncomment and adjust if on Windows):
# Before:
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# After (Uncommented):
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# 2. Windows Poppler Path (Uncomment and adjust if Poppler is not in System PATH):
# POPPLER_PATH = r"C:\Users\Jerikoy\Downloads\Release-26.09.0-0.zip\poppler-26.09.0\Library\bin" 
POPPLER_PATH = r'C:\Users\Jerikoy\Downloads\Release-26.09.0-0\poppler-26.09.0\Library\bin'

def ocr_pdf_to_excel(pdf_path, excel_path):
    print(f"Opening {pdf_path}...")
    
    if not os.path.exists(pdf_path):
        print(f"Error: The file '{pdf_path}' was not found in the current folder.")
        return

    try:
        # Convert PDF to images
        if POPPLER_PATH:
            images = pdf2image.convert_from_path(pdf_path, poppler_path=POPPLER_PATH)
        else:
            images = pdf2image.convert_from_path(pdf_path)
            
    except Exception as e:
        print(f"\n[Poppler Error] Failed to read PDF: {e}")
        print("Ensure Poppler is installed and POPPLER_PATH is set correctly.")
        return

    all_data = []

    for page_num, img in enumerate(images, start=1):
        print(f"Processing page {page_num} of {len(images)}...")
        
        try:
            # Extract text via Tesseract OCR
            text = pytesseract.image_to_string(img)
            
            # Split into lines for structured tabular export
            for line in text.split('\n'):
                clean_line = line.strip()
                if clean_line:
                    all_data.append({
                        "Page": page_num,
                        "Extracted Text": clean_line
                    })
        except Exception as e:
            print(f"\n[Tesseract Error] Failed on page {page_num}: {e}")
            print("Ensure Tesseract-OCR is installed and tesseract_cmd path is correct.")
            return

    # Export to Excel
    if all_data:
        df = pd.DataFrame(all_data)
        df.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"\nSuccess! Output saved to: {excel_path}")
    else:
        print("\nNo readable text found in PDF.")

if __name__ == "__main__":
    # Ensure 'document.pdf' exists in your VS Code workspace folder
    ocr_pdf_to_excel("Daabay.resumeee.pdf", "output.xlsx")