import ocrmypdf
import pandas as pd
import fitz # Assuming PyMuPDF is already installed
import os

def ocr_my_pdf(rutapdf, rutasalida):
    error_log = {}
    try:
        # Perform OCR and create new PDF with extracted text
        output_file = ocrmypdf.ocr(
        rutapdf, rutasalida, output_type="pdf", skip_text=True, deskew=True
        )

        extraction_pdfs = {}
        pages_df = pd.DataFrame(columns=["text"])
        doc = fitz.open(output_file) # Open the newly created OCRed PDF
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            pages_df = pd.concat(
                [pages_df, pd.DataFrame([{"text": page.get_text("text")}])],
                ignore_index=True,
        )
        extraction_pdfs[output_file] = pages_df
        return extraction_pdfs
    except Exception as e:
        error_log[rutapdf] = str(e) # Convert exception to string for better logging
    

