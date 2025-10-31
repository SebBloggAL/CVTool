# main.py

import logging
import os
import re

from text_extractor import extract_text
from data_extractor import extract_cv_data
from formatter import format_data
from document_generator import create_document
from file_handler import validate_file
from experience_parser import extract_experience_lines
from section_config import SECTION_SYNONYMS

# --- Regex for dates and bullets used in structuring ---
MONTH = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*"
DATE_TOKEN = rf"(?:{MONTH}\s+\d{{4}}|\d{{1,2}}/\d{{4}}|\d{{4}})"
PRESENT = r"(?:Present|Current|Now)"
RANGE_SEP = r"[–—-]"  # en/em/hyphen
DURATION_LINE_RE = re.compile(rf"^\s*{DATE_TOKEN}\s*{RANGE_SEP}\s*(?:{PRESENT}|{DATE_TOKEN})(?:.*)?$", re.IGNORECASE)
PAREN_DURATION_RE = re.compile(rf"\((\s*{DATE_TOKEN}\s*{RANGE_SEP}\s*(?:{PRESENT}|{DATE_TOKEN})\s*)\)", re.IGNORECASE)
HEADER_DASH_RE = re.compile(r".+\s[–—-]\s.+")
BULLET_RE = re.compile(r'^\s*(?:[•\-\*\u2013\u2014\u00B7\u2219\u25AA\u25E6]|\d+[\.\)]|[A-Za-z]\))\s+')

# Security clearance fallback
CLEARANCE_RE = re.compile(r'\b(DV|SC|CTC|BPSS)\s*(?:cleared|clearance)?\b', re.IGNORECASE)

def _mark_sections(text: str) -> str:
    """
    Insert explicit section markers using SECTION_SYNONYMS so LLM and slicers see aligned boundaries.
    """
    if not text:
        return ""
    marked = text
    for sec, variants in SECTION_SYNONYMS.items():
        pat = r'(^|\n)\s*(?:' + '|'.join(variants) + r')\s*(\n|$)'
        marked = re.sub(pat, f"\n=== {sec} ===\n", marked, flags=re.IGNORECASE)
    return marked

def _strip_bullet(s): return BULLET_RE.sub("", s).strip()

def _is_header(s):
    s_stripped = s.strip()
    if HEADER_DASH_RE.match(s_stripped):
        return True
    if s_stripped.lower().startswith("earlier "):
        return True
    if PAREN_DURATION_RE.search(s_stripped):
        return True
    t = re.sub(r'[\[\]]', '', s_stripped).strip().lower()
    if t in {"technical skills", "skills", "education", "certifications", "summary"}:
        return False
    return False

def _structure_experience_from_lines(exp_lines):
    """
    Group header + (optional) duration + following lines as responsibilities. Keep wording verbatim.
    If header contains a parenthesised date window, remove it from Position and store in Duration.
    """
    items, i, n = [], 0, len(exp_lines)
    while i < n:
        line = exp_lines[i].rstrip()
        if not line:
            i += 1
            continue

        if _is_header(line):
            position_text = line.strip()

            # duration inside header?
            duration_text = ""
            m = PAREN_DURATION_RE.search(position_text)
            if m:
                duration_text = m.group(1).strip()
                position_text = PAREN_DURATION_RE.sub("", position_text).strip()

            # or as next line?
            used_next_for_duration = False
            if not duration_text and (i + 1) < n and DURATION_LINE_RE.match(exp_lines[i + 1].strip()):
                duration_text = exp_lines[i + 1].strip()
                used_next_for_duration = True

            item = {
                "Position": position_text,
                "Company": "",
                "Duration": duration_text,
                "Responsibilities": [],
            }
            i += 2 if used_next_for_duration else 1

            while i < n:
                peek = exp_lines[i].strip()
                if not peek:
                    i += 1
                    continue
                if _is_header(peek) or DURATION_LINE_RE.match(peek):
                    break
                sec_name = re.sub(r'[\[\]]', '', peek).strip().lower()
                if sec_name in {"technical skills", "skills", "education", "certifications", "summary"}:
                    break
                item["Responsibilities"].append(_strip_bullet(peek))
                i += 1

            items.append(item)
            continue

        # stray duration without header: skip
        if DURATION_LINE_RE.match(line):
            i += 1
            continue

        i += 1

    return items

def _fallback_clearance(raw_text: str) -> str | None:
    m = CLEARANCE_RE.search(raw_text or "")
    if not m:
        return None
    code = m.group(1).upper()
    return {"DV": "DV Cleared", "SC": "SC Cleared", "CTC": "CTC Cleared", "BPSS": "BPSS"}.get(code, f"{code} Cleared")

def main(file_path, output_directory='Documents/Processed'):
    """
    Main function to process the CV file.
    """
    try:
        logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

        # Validate the input file
        validate_file(file_path)
        logging.info("File validation completed.")

        # Extract raw text (single newlines preserved by text_extractor.normalize_text)
        text = extract_text(file_path)
        logging.info("Text extraction completed.")

        # Insert markers with synonyms
        marked_text = _mark_sections(text)

        # Deterministic, verbatim capture of Experience
        exp_lines = extract_experience_lines(marked_text) or extract_experience_lines(text)
        logging.debug(f"Verbatim Experience lines count: {len(exp_lines)}")
        if exp_lines:
            logging.debug("First few Experience lines: " + " | ".join(exp_lines[:5]))

        exp_struct = _structure_experience_from_lines(exp_lines)
        logging.debug(f"Verbatim Experience structured items: {len(exp_struct)}")

        # LLM for non-experience fields only
        raw_data = extract_cv_data(marked_text)
        logging.debug(f"Raw data extracted (pre-override) keys: {list(raw_data.keys())}")

        # HARD OVERRIDE: ensure verbatim Experience wins (even if empty)
        raw_data["Experience"] = exp_struct

        # Fallback for SecurityClearance
        if not raw_data.get("SecurityClearance") or raw_data["SecurityClearance"].strip().lower() == "not specified":
            sc = _fallback_clearance(text)
            if sc:
                raw_data["SecurityClearance"] = sc

        logging.info("Data extraction completed.")

        # Format (formatter will sort roles by end date; bullets remain verbatim)
        data = format_data(raw_data)
        logging.debug(f"Formatted data prepared.")

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
