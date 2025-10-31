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
    Insert explicit section markers with sensible synonyms so that both the LLM
    and our verbatim slicer see aligned boundaries.

    Mappings:
      - Summary: "Summary", "[Summary]", "Profile", "Professional Summary"
      - Skills:  "Skills", "[Skills]", "Technical Skills", "Core Skills", "Key Skills"
      - Experience: "Experience" + broad variants incl. "Career Summary"
      - Education / Certifications: exact names or bracketed variants
    """
    if not text:
        return ""

    marked = text

    synonyms = {
        "Summary": [
            r"\[?\s*Summary\s*\]?",
            r"Profile",
            r"Professional\s+Summary",
        ],
        "Skills": [
            r"\[?\s*Skills\s*\]?",
            r"Technical\s+Skills",
            r"Core\s+Skills",
            r"Key\s+Skills",
        ],
        "Experience": [
            r"\[?\s*Experience\s*\]?",
            r"Professional\s+Experience",
            r"Work\s+Experience",
            r"Employment\s+History",
            r"Career\s+History",
            r"Relevant\s+Experience",
            r"Career\s+Summary",  # <-- critical for your CVs
        ],
        "Education": [
            r"\[?\s*Education\s*\]?",
        ],
        "Certifications": [
            r"\[?\s*Certifications\s*\]?",
            r"Qualifications",
            r"Certificates",
        ],
    }

    # Apply all mappings; each pattern must be on its own line or delimited by newlines.
    for sec, variants in synonyms.items():
        pat = r'(^|\n)\s*(?:' + '|'.join(variants) + r')\s*(\n|$)'
        marked = re.sub(pat, lambda m, s=sec: f"\n=== {s} ===\n", marked, flags=re.IGNORECASE)

    return marked


# --- Grouping helpers for structuring Experience ---------------------------------

# Date token patterns: "Sep 2012", "September 2012", "09/2012", "2012"
MONTH = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*"
DATE_WORD = rf"(?:{MONTH}\s+\d{{4}}|\d{{1,2}}/\d{{4}}|\d{{4}})"
PRESENT_WORD = r"(?:Present|Current|Now)"
RANGE_SEP = r"[–—-]"  # en dash / em dash / hyphen

# Standalone duration line: "Sep 2012 – May 2014", "Apr 2022 – Present", "2004 – 2012"
DURATION_LINE_RE = re.compile(
    rf"^\s*{DATE_WORD}\s*{RANGE_SEP}\s*(?:{PRESENT_WORD}|{DATE_WORD})(?:.*)?$",
    re.IGNORECASE,
)

# Header with duration in parentheses: "Earlier Career (1991 – 2004)"
PAREN_DURATION_RE = re.compile(
    rf"\((\s*{DATE_WORD}\s*{RANGE_SEP}\s*(?:{PRESENT_WORD}|{DATE_WORD})\s*)\)",
    re.IGNORECASE,
)

# Company/role style header with dash: "BDR Thermea (BAXI) – Business Intelligence Manager"
HEADER_DASH_RE = re.compile(r".+\s[–—-]\s.+")

# Bullet markers we will strip from verbatim responsibility lines if present
BULLET_RE = re.compile(r'^\s*(?:[•\-\*\u2013\u2014\u00B7\u2219\u25AA\u25E6]|\d+[\.\)]|[A-Za-z]\))\s+')


def _structure_experience_from_lines(exp_lines):
    """
    Convert verbatim lines into a structured list of roles without changing wording.

    Rules:
      - A header line is either:
          * line with a spaced dash between phrases (Company – Role), or
          * line like "Earlier Career (1991 – 2004)", or
          * "Earlier ..." standalone headings we want to keep.
      - A duration line matches DURATION_LINE_RE (e.g., "Apr 2022 – Present ...").
      - We group: [HEADER] [optional DURATION-LINE or duration-in-parentheses on header]
        then treat all following non-header lines as responsibilities until the next header.
      - Each responsibility line is kept verbatim (bullet marker stripped if present).
    """
    items = []
    i = 0
    n = len(exp_lines)

    def strip_bullet(s):
        return BULLET_RE.sub("", s).strip()

    def is_header(s):
        s_stripped = s.strip()
        if HEADER_DASH_RE.match(s_stripped):
            return True
        if s_stripped.lower().startswith("earlier "):
            return True
        if PAREN_DURATION_RE.search(s_stripped):
            return True
        # Avoid misclassifying "Technical Skills" as a role header here
        t = re.sub(r'[\[\]]', '', s_stripped).strip().lower()
        if t in {"technical skills", "skills", "education", "certifications", "summary"}:
            return False
        return False

    while i < n:
        line = exp_lines[i].rstrip()
        if not line:
            i += 1
            continue

        if is_header(line):
            # Start a new role
            position_text = line.strip()

            # 1) Duration inside header?
            duration_text = ""
            m = PAREN_DURATION_RE.search(position_text)
            if m:
                duration_text = m.group(1).strip()

            # 2) Or duration on the next line?
            used_next_for_duration = False
            if not duration_text and (i + 1) < n and DURATION_LINE_RE.match(exp_lines[i + 1].strip()):
                duration_text = exp_lines[i + 1].strip()
                used_next_for_duration = True

            # Create the role item
            item = {
                "Position": position_text,     # keep verbatim; do not split Company/Role
                "Company": "",
                "Duration": duration_text,
                "Responsibilities": [],
            }

            # Advance past header (+ optional duration line)
            i += 2 if used_next_for_duration else 1

            # Collect responsibilities until next header/section-like line
            while i < n:
                peek = exp_lines[i].strip()
                if not peek:
                    i += 1
                    continue
                # Stop if the next line starts a new header/role
                if is_header(peek) or DURATION_LINE_RE.match(peek):
                    break
                # Stop if we accidentally ran into a section heading
                sec_name = re.sub(r'[\[\]]', '', peek).strip().lower()
                if sec_name in {"technical skills", "skills", "education", "certifications", "summary"}:
                    break
                item["Responsibilities"].append(strip_bullet(peek))
                i += 1

            items.append(item)
            continue

        # If the current line looks like a duration but we don't have a header, skip (cannot anchor).
        if DURATION_LINE_RE.match(line):
            i += 1
            continue

        # Otherwise, not a header — skip forward
        i += 1

    return items


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

        # 1) Insert markers with synonyms (Career Summary -> Experience; Technical Skills -> Skills)
        marked_text = _mark_sections(text)

        # 2) Deterministic, verbatim capture of Experience
        exp_lines = extract_experience_lines(marked_text) or extract_experience_lines(text)
        logging.debug(f"Verbatim Experience lines count: {len(exp_lines)}")
        if exp_lines:
            logging.debug("First few Experience lines: " + " | ".join(exp_lines[:5]))

        exp_struct = _structure_experience_from_lines(exp_lines)
        logging.debug(f"Verbatim Experience structured items: {len(exp_struct)}")

        # 3) LLM for non-experience fields only
        raw_data = extract_cv_data(marked_text)
        logging.debug(f"Raw data extracted (pre-override): dict_keys({list(raw_data.keys())})")

        # 4) HARD OVERRIDE: ensure verbatim Experience wins (even if empty)
        raw_data["Experience"] = exp_struct

        logging.info("Data extraction completed.")

        # 5) Format (your formatter will sort roles by end date; bullets remain verbatim)
        data = format_data(raw_data)
        logging.debug(f"Formatted data: {data}")
        logging.info("Data formatting completed.")

        # 6) Output
        applicant_name = data.get('ApplicantName', 'output').replace(" ", "_")
        output_filename = f"{applicant_name}_CV.docx"
        output_path = os.path.join(output_directory, output_filename)
        os.makedirs(output_directory, exist_ok=True)

        create_document(data, output_path=output_path)
        logging.info(f"Document saved to {output_path}")
        logging.info("Document generation completed.")
        return output_path

    except Exception as e:
        logging.error(f"An error occurred: {e}", exc_info=True)
        raise
