# main.py

import logging
import os
import re

from text_extractor import extract_text
from data_extractor import extract_cv_data
from formatter import format_data
from document_generator import create_document
from file_handler import validate_file
from experience_parser import extract_experience_lines  # robust slice


def _mark_sections(text: str) -> str:
    """
    Insert explicit section markers, carefully:
      - Summary, Skills, Education, Certifications: only match the exact heading or [Heading].
      - Experience: match broader synonyms.
    """
    if not text:
        return ""

    marked = text

    # Exact headings (and bracketed variants)
    def mark_exact(sec):
        pattern = rf'(^|\n)\s*(\[{sec}\]|{sec})\s*(\n|$)'
        return re.sub(pattern, f'\n=== {sec} ===\n', marked, flags=re.IGNORECASE)

    # Start by marking exact, non-Experience sections
    for sec in ("Summary", "Skills", "Education", "Certifications"):
        marked = re.sub(
            rf'(^|\n)\s*(\[{sec}\]|{sec})\s*(\n|$)',
            f'\n=== {sec} ===\n',
            marked,
            flags=re.IGNORECASE
        )

    # Experience with broad synonyms (ONLY for Experience)
    exp_variants = [
        r'\[?\s*Experience\s*\]?',
        r'Professional\s+Experience',
        r'Work\s+Experience',
        r'Employment\s+History',
        r'Career\s+History',
        r'Relevant\s+Experience',
    ]
    exp_pat = r'(^|\n)\s*(?:' + '|'.join(exp_variants) + r')\s*(\n|$)'
    marked = re.sub(exp_pat, '\n=== Experience ===\n', marked, flags=re.IGNORECASE)

    return marked


def _structure_experience_from_lines(exp_lines):
    """
    Convert verbatim lines into the minimal structure expected by the document generator,
    without rewriting or shortening any bullet text.

    Heuristic:
      - Common bullet markers: •, -, *, –, —, ·, ▪, ◦, '1.' etc.
      - Non-bullet line => role header (kept verbatim as Position).
      - We do NOT set Company/Duration here (so the writer won't recompose a title).
    """
    experience_struct = []
    current = {"Position": "", "Company": "", "Duration": "", "Responsibilities": []}

    bullet_re = re.compile(r'^\s*(?:[•\-\*\u2013\u2014\u00B7\u2219\u25AA\u25E6]|\d+[\.\)]|[A-Za-z]\))\s+')

    def flush_current():
        nonlocal current, experience_struct
        if current["Position"] or current["Responsibilities"]:
            experience_struct.append(current)
        current = {"Position": "", "Company": "", "Duration": "", "Responsibilities": []}

    for raw in exp_lines:
        ln = raw.rstrip()
        if not ln:
            continue
        if bullet_re.match(ln):
            text = bullet_re.sub('', ln).strip()
            current["Responsibilities"].append(text)
        else:
            if current["Position"] or current["Responsibilities"]:
                flush_current()
            current["Position"] = ln.strip()

    flush_current()
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
        logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

        # Validate the input file
        validate_file(file_path)
        logging.info("File validation completed.")

        # Extract raw text (we preserve single newlines in text_extractor.normalize_text)
        text = extract_text(file_path)
        logging.info("Text extraction completed.")

        # Align markers so verbatim slice and LLM see the same boundaries
        marked_text = _mark_sections(text)

        # Deterministically capture Experience BEFORE calling the LLM
        exp_lines_verbatim = extract_experience_lines(marked_text) or extract_experience_lines(text)
        exp_struct_verbatim = _structure_experience_from_lines(exp_lines_verbatim)

        logging.debug(f"Verbatim Experience lines count: {len(exp_lines_verbatim)}")
        logging.debug(f"Verbatim Experience structured items: {len(exp_struct_verbatim)}")

        # Extract remaining data via LLM (basic fields, summary, skills, education, certs)
        raw_data = extract_cv_data(marked_text)
        logging.debug(f"Raw data extracted (pre-override): {raw_data.keys()}")

        # HARD OVERRIDE: never allow LLM Experience to leak through
        raw_data["Experience"] = exp_struct_verbatim  # even if empty, we prefer empty over summarised

        logging.info("Data extraction completed.")

        # Format (formatter will sort roles; it must NOT rewrite bullets)
        data = format_data(raw_data)
        logging.debug("Data formatted.")

        # Construct the output path
        applicant_name = data.get('ApplicantName', 'output').replace(" ", "_")
        output_filename = f"{applicant_name}_CV.docx"
        output_path = os.path.join(output_directory, output_filename)
        os.makedirs(output_directory, exist_ok=True)

        # Generate the standardized Word document
        create_document(data, output_path=output_path)
        logging.info("Document generation completed.")
        return output_path

    except Exception as e:
        logging.error(f"An error occurred: {e}", exc_info=True)
        raise
