# main.py

import logging
import os
import re

from text_extractor import extract_text
from data_extractor import extract_cv_data
from formatter import format_data
from document_generator import create_document
from file_handler import validate_file


def _extract_experience_lines_verbatim(full_text: str):
    """
    Return the lines from the Experience section *verbatim* (no model involvement).
    We look for a start heading ("Experience" variants) and stop at the next major heading.
    """
    experience_headings = {"experience", "professional experience", "[experience]"}
    stop_headings = {
        "education", "certifications", "skills", "summary",
        "[education]", "[certifications]", "[skills]", "[summary]"
    }

    # Keep original line breaks to avoid collapsing bullets
    lines = [ln.strip() for ln in (full_text or "").splitlines()]

    # Find the start of the experience section
    start = None
    for i, ln in enumerate(lines):
        t = re.sub(r'[\[\]]', '', ln).strip().lower()
        if t in experience_headings:
            start = i + 1
            break
    if start is None:
        return []

    # Find the end (next major section)
    end = len(lines)
    for j in range(start, len(lines)):
        t = re.sub(r'[\[\]]', '', lines[j]).strip().lower()
        if t in stop_headings:
            end = j
            break

    # Keep non-empty lines only
    return [ln for ln in lines[start:end] if ln]


def _structure_experience_from_lines(exp_lines):
    """
    Convert verbatim lines into a minimal structure expected by the document generator,
    without rewriting or shortening any bullet text.

    Heuristic:
      - A line starting with a bullet marker (•, -, *) is a responsibility.
      - Any non-bullet line starts (or updates) the current role header (stored in "Position").
      - We DO NOT attempt to split 'Position / Company / Duration' from the header to avoid
        accidental abbreviation; the header is kept as-is in Position.
    """
    experience_struct = []
    current = {"Position": "", "Company": "", "Duration": "", "Responsibilities": []}

    bullet_re = re.compile(r'^\s*[•\-\*]\s+')

    for ln in exp_lines:
        if bullet_re.match(ln):
            # Responsibility bullet: strip the marker, keep the text verbatim
            text = bullet_re.sub('', ln).strip()
            current["Responsibilities"].append(text)
        else:
            # New header (role line). If we already have an item accumulating, push it.
            # Keep the header verbatim (store in Position to avoid any lossy parsing).
            if any(v for v in current.values()):
                experience_struct.append(current)
            current = {"Position": ln.strip(), "Company": "", "Duration": "", "Responsibilities": []}

    if any(v for v in current.values()):
        experience_struct.append(current)

    return experience_struct


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

        # Extract text from the file (preserves line breaks for bullets)
        text = extract_text(file_path)
        logging.info("Text extraction completed.")

        # === CHANGE: deterministically capture Experience *before* LLM to avoid abbreviation
        exp_lines_verbatim = _extract_experience_lines_verbatim(text)
        exp_struct_verbatim = _structure_experience_from_lines(exp_lines_verbatim)
        logging.debug(f"Verbatim Experience (lines): {exp_lines_verbatim}")
        logging.debug(f"Verbatim Experience (structured): {exp_struct_verbatim}")

        # Extract the rest of the data using the LLM (basic fields, summary, skills, education, etc.)
        raw_data = extract_cv_data(text)
        logging.debug(f"Raw data extracted (pre-override): {raw_data}")

        # === CHANGE: override LLM-produced Experience with verbatim structured version
        if exp_struct_verbatim:
            raw_data["Experience"] = exp_struct_verbatim
            logging.debug("Overwrote LLM 'Experience' with verbatim experience from source document.")

        logging.info("Data extraction completed.")

        # Format the data (formatter keeps bullet text; order is preserved unless you enabled sorting there)
        data = format_data(raw_data)
        logging.debug(f"Formatted data: {data}")
        logging.info("Data formatting completed.")

        # Construct the output path
        applicant_name = data.get('ApplicantName', 'output').replace(" ", "_")
        output_filename = f"{applicant_name}_CV.docx"
        output_path = os.path.join(output_directory, output_filename)

        # Ensure output directory exists
        os.makedirs(output_directory, exist_ok=True)

        # Generate the standardized Word document
        create_document(data, output_path=output_path)
        logging.info("Document generation completed.")

        logging.info("CV processed successfully.")
        return output_path

    except Exception as e:
        logging.error(f"An error occurred: {e}", exc_info=True)
        raise
