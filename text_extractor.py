# text_extractor.py

import os
from pdfminer.high_level import extract_text as extract_pdf_text
import docx2txt
import logging

# Suppress pdfminer warnings
logging.getLogger('pdfminer').setLevel(logging.WARNING)

def extract_text(file_path):
    """
    Extracts text from a file (PDF or DOCX).
    """
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()

    if ext == '.pdf':
        logging.info("Extracting text from PDF file.")
        return extract_text_from_pdf(file_path)
    elif ext == '.docx':
        logging.info("Extracting text from DOCX file.")
        return extract_text_from_docx(file_path)
    else:
        raise ValueError(f"Unsupported file extension: {ext}")

def extract_text_from_pdf(file_path):
    """
    Extracts text from a PDF file using pdfminer.six.
    """
    try:
        text = extract_pdf_text(file_path)
        return text
    except Exception as e:
        logging.error(f"Error extracting text from PDF: {e}")
        raise

def extract_text_from_docx(file_path):
    """
    Extracts text from a DOCX file using docx2txt.
    """
    try:
        text = docx2txt.process(file_path)
        return text
    except Exception as e:
        logging.error(f"Error extracting text from DOCX: {e}")
        raise
