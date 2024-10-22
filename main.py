# main.py

import logging
from text_extractor import extract_text
from data_extractor import extract_cv_data
from formatter import format_data
from document_generator import create_document
from file_handler import validate_file
import os

def main(file_path, output_directory='Documents/Processed'):
    """
    Main function to process the CV file.

    Parameters:
    - file_path (str): Path to the uploaded CV file.
    - output_directory (str): Directory to save the processed document.

    Returns:
    - output_path (str): Path to the generated document.
    """
    try:
        # Set up logging with DEBUG level for detailed information
        logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

        # Validate the input file
        validate_file(file_path)
        logging.info("File validation completed.")

        # Extract text from the file
        text = extract_text(file_path)
        logging.info("Text extraction completed.")

        # Extract data using OpenAI API
        raw_data = extract_cv_data(text)
        logging.debug(f"Raw data extracted: {raw_data}")  # Debug log for raw data

        logging.info("Data extraction completed.")

        # Format the data
        data = format_data(raw_data)
        logging.debug(f"Formatted data: {data}")  # Debug log for formatted data
        logging.info("Data formatting completed.")

        # Construct the output path
        applicant_name = data.get('ApplicantName', 'output').replace(" ", "_")
        output_filename = f"{applicant_name}_CV.docx"
        output_path = os.path.join(output_directory, output_filename)

        # Generate the standardized Word document
        create_document(data, output_path=output_path)
        logging.info("Document generation completed.")

        logging.info("CV processed successfully.")
        return output_path

    except Exception as e:
        logging.error(f"An error occurred: {e}")
        raise
