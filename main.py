# main.py

import logging
import os
import re

from text_extractor import extract_text
from data_extractor import extract_cv_data
from formatter import format_data
from document_generator import create_document
from file_handler import validate_file
from experience_parser import extract_experience_lines  # robust slice w/ markers & fallbacks


def _mark_sections(text: str) -> str:
    """
    Insert the same explicit section markers used by data_extractor so our verbatim
    extractor can slice *exactly* the Experience block the LLM sees.
    """
    if not text:
        return ""

    sections = ["Summary", "Skills", "Experience", "Education", "Certifications"]
    marked = text

    for sec in sections:
        # Add markers only when a standalone heading line appears (various variants)
        # Examples: "Experience", "[Experience]", "WORK EXPERIENCE", "Professional Experience"
        pattern = rf'(^|\n)\s*(\[{sec}\]|{sec}|(?i:{sec})|(?i:work\s+experience)|(?i:employment\s+history)|(?i:professional\s+experience))\s*(\n|$)'
        marked = re.sub(pattern, lambda m: f'\n=== {sec} ===\n', marked, flags=re.IGNORECASE)

    return marked


def _structure_experience_from_lines(exp_lines):
    """
    Convert verbatim lines into a minimal structure expected by the document generator,
    without rewriting or shortening any bullet text.

    Heuristic:
      - Bullet lines start with common markers: •, -, *, –, —, ·, ▪, ◦, numbers like '1.' etc.
      - The first non-bullet line after a blank or at the start is treated as a role header (Position).
      - We keep the header verbatim inside 'Position'. We do NOT attempt to split Company/Duration.
      - If we never see a header, we put all bullets under a single item with an empty Position.
    """
    experience_struct = []
    current = {"Position": "", "Company": "", "Duration": "", "Responsibilities": []}

    bullet_re = re.compile(r'^\s*(?:[•\-\*\u2013\u2014\u00B7\u2219\u25AA\u25E6]|\d+[\.\)]|[A-Za-z]\))\s+')

    seen_any_header = False
    pending_header = None

    def flush_current():
        nonlocal current, experience_struct
        if current["Position"] or current["Responsibilities"]:
            experience_struct.append(current)
        current = {"Position": "", "Company": "", "Duration": "", "Responsibilities": []}

    prev_blank = True  # treat the very start as a boundary

    for raw in exp_lines:
        ln = raw.rstrip()

        if not ln:
            prev_blank = True
            continue

        if bullet_re.match(ln):
            # Responsibility bullet: strip marker, keep verbatim remainder
            text = bullet_re.sub('', ln).strip()
            current["Responsibilities"].append(text)
            prev_blank = False
            continue

        # Non-bullet line — likely a role header
        # Start a new role if we already have content
        if current["Position"] or current["Responsibilities"]:
            flush_current()

        current["Position"] = ln.strip()
        seen_any_header = True
        prev_blank = False

    flush_current()

    # Fallback: if we never found a header but there are lines,
    # put everything as responsibilities under a single (empty-position) role.
    if not seen_any_header and not experience_struct:
        if exp_lines:
            return [{
                "Position": "",
                "Company": "",
                "Duration": "",
                "Responsibilities": [re.sub(bullet_re, '', x).strip() for x in exp_lines if x.strip()]
            }]

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

        # === NEW: mark sections so both the LLM and verbatim path have aligned boundaries
        marked_text = _mark_sections(text)

        # === deterministically capture Experience *before* or alongside LLM to avoid abbreviation
        # Prefer marker-based slice; fall back to heading-based slice if markers not present.
        exp_lines_verbatim = extract_experience_lines(marked_text)
        if not exp_lines_verbatim:
            # Fallback to unmarked text (broad heading variants)
            exp_lines_verbatim = extract_experience_lines(text)

        exp_struct_verbatim = _structure_experience_from_lines(exp_lines_verbatim)
        logging.debug(f"Verbatim Experience (lines): {exp_lines_verbatim}")
        logging.debug(f"Verbatim Experience (structured): {exp_struct_verbatim}")

        # Extract the rest of the data using the LLM (basic fields, summary, skills, education, etc.)
        raw_data = extract_cv_data(marked_text)
        logging.debug(f"Raw data extracted (pre-override): {raw_data}")

        # === IMPORTANT: always override LLM-produced Experience with verbatim structured version
        # If our struct is empty but we *do* have lines, wrap them as a single role.
        if not exp_struct_verbatim and exp_lines_verbatim:
            exp_struct_verbatim = [{
                "Position": "",
                "Company": "",
                "Duration": "",
                "Responsibilities": exp_lines_verbatim[:]  # keep every line verbatim
            }]

        if exp_struct_verbatim:
            raw_data["Experience"] = exp_struct_verbatim
            logging.debug("Overwrote LLM 'Experience' with verbatim experience from source document.")

        logging.info("Data extraction completed.")

        # Format the data (formatter will sort roles; bullets remain verbatim)
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
