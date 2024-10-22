# file_handler.py

import os

ALLOWED_EXTENSIONS = ['.pdf', '.docx']

def validate_file(file_path):
    if not os.path.isfile(file_path):
        raise FileNotFoundError("File does not exist.")
    ext = os.path.splitext(file_path)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise ValueError(f"Unsupported file type: {ext}")
