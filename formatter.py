# formatter.py

import re
import logging
from datetime import datetime
from typing import List, Dict, Any

# Reuse date parsing from document_generator for consistency
from document_generator import parse_end_date

# Parentheses-aware comma splitting for skills
PAREN_AWARE_COMMA = re.compile(r',(?![^()]*\))')

# Date window detection anywhere in a string
MONTH = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*"
DATE_TOKEN = rf"(?:{MONTH}\s+\d{{4}}|\d{{1,2}}/\d{{4}}|\d{{4}})"
PRESENT = r"(?:Present|Current|Now)"
RANGE_SEP = r"[–—-]"
DUR_WINDOW_RE = re.compile(rf"({DATE_TOKEN})\s*{RANGE_SEP}\s*({PRESENT}|{DATE_TOKEN})", re.IGNORECASE)

def format_data(raw_data):
    """
    Formats the raw data extracted from the CV to match the template requirements.
    """
    security_clearance = raw_data.get("SecurityClearance", "Not specified") or "Not specified"

    formatted_data = {
        "ApplicantName": raw_data.get("ApplicantName", "Name not provided"),
        "Role": raw_data.get("Role", "Role not specified"),
        "SecurityClearance": security_clearance,
        "Summary": raw_data.get("Summary", "Summary not provided"),
        "Skills": format_skills(raw_data.get("Skills", [])),
        "Experience": format_experience(raw_data.get("Experience", [])),  # sorted newest->oldest
        "Education": format_education(raw_data.get("Education", [])),
        "Certifications": format_certifications(raw_data.get("Certifications", [])),
        # Optional additional content if you decide to include it later:
        "Additional": raw_data.get("Additional", []),
    }

    logging.debug(f"Formatted data: keys={list(formatted_data.keys())}")
    return formatted_data


def format_skills(skills_data):
    """
    Formats the skills data into a list of individual skills without splitting inside parentheses.
    """
    if isinstance(skills_data, list):
        return [s.strip() for s in skills_data if isinstance(s, str) and s.strip()]

    if isinstance(skills_data, str):
        out = []
        for raw_line in skills_data.splitlines():
            line = (raw_line or "").strip()
            if not line:
                continue
            # Keep category lines as-is (e.g., "Cloud Platforms: AWS, Azure")
            if ':' in line:
                out.append(line)
                continue
            # Otherwise split by commas that are NOT inside parentheses
            for token in PAREN_AWARE_COMMA.split(line):
                tok = token.strip()
                if tok:
                    out.append(tok)
        return out

    logging.warning("Unexpected format for skills data.")
    return []


def format_experience(experience_data):
    """
    Keep responsibilities verbatim; sort roles by parsed end date (newest first).
    """
    items: List[Dict[str, Any]] = []

    if isinstance(experience_data, list):
        for it in experience_data:
            if not isinstance(it, dict):
                logging.warning(f"Unexpected experience item type: {type(it)}")
                continue
            items.append({
                "Position": it.get("Position", ""),
                "Company": it.get("Company", ""),
                "Duration": it.get("Duration", ""),
                "Responsibilities": it.get("Responsibilities", []),  # DO NOT rewrite
            })

    elif isinstance(experience_data, dict):
        items.append({
            "Position": experience_data.get("Position", ""),
            "Company": experience_data.get("Company", ""),
            "Duration": experience_data.get("Duration", ""),
            "Responsibilities": experience_data.get("Responsibilities", []),
        })
    else:
        logging.warning("Unexpected format for experience data.")

    try:
        items = sort_experiences(items)
    except Exception as e:
        logging.warning(f"Could not sort experiences: {e}")

    return items


def format_education(education_data):
    """
    Formats the education data as a single string with entries separated by two line breaks.
    """
    formatted = ""
    if isinstance(education_data, list):
        for item in education_data:
            if not isinstance(item, dict):
                continue
            degree = (item.get("Degree", "") or "").strip()
            institution = (item.get("Institution", "") or "").strip()
            duration = (item.get("Duration", "") or "").strip()

            entry = degree
            if institution and institution.lower() != "not specified":
                entry += f" at {institution}"
            if duration and duration.lower() != "not specified":
                entry += f" ({duration})"

            entry = entry.strip()
            if entry:
                formatted += f"{entry}\n\n"

    elif isinstance(education_data, dict):
        for k, v in education_data.items():
            formatted += f"{k}\n"
            if isinstance(v, list):
                for detail in v:
                    formatted += f"- {detail}\n"
            elif isinstance(v, str):
                formatted += f"- {v}\n"
            formatted += "\n"

    elif isinstance(education_data, str):
        formatted += education_data

    return formatted.strip()


def format_certifications(cert_data):
    """
    Formats the certifications data as a list of strings.
    """
    if isinstance(cert_data, list):
        return [c.strip() for c in cert_data if isinstance(c, str) and c.strip()]
    if isinstance(cert_data, str):
        return [s.strip() for s in re.split(r'[\n,]', cert_data) if s.strip()]
    logging.warning("Unexpected format for certifications data.")
    return []


# ---- Sorting helpers ----

def _extract_date_window(text: str) -> str | None:
    if not isinstance(text, str) or not text.strip():
        return None
    m = DUR_WINDOW_RE.search(text)
    if not m:
        return None
    start, end = m.group(1), m.group(2)
    if re.match(PRESENT, end, re.IGNORECASE):
        end = "Present"
    return f"{start} – {end}"

def _best_effort_end_dt(item: Dict[str, Any]) -> datetime:
    dur = (item.get("Duration") or "").strip()
    win = _extract_date_window(dur)
    if win:
        try:
            return parse_end_date(win)
        except Exception:
            pass

    pos = (item.get("Position") or "").strip()
    win = _extract_date_window(pos)
    if win:
        try:
            return parse_end_date(win)
        except Exception:
            pass

    return datetime.min

def sort_experiences(experience_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    indexed: List[tuple[int, datetime, Dict[str, Any]]] = []
    for idx, item in enumerate(experience_data):
        end_dt = _best_effort_end_dt(item)
        indexed.append((idx, end_dt, item))

    # Sort by end date DESC, tie-break by original order (stable)
    indexed.sort(key=lambda t: (t[1], -t[0]), reverse=True)

    # Debug order
    debug_view = [
        {"i": i, "end_dt": ed.isoformat() if isinstance(ed, datetime) else str(ed),
         "pos": it.get("Position", "")[:80], "dur": it.get("Duration", "")[:80]}
        for (i, ed, it) in indexed
    ]
    logging.debug(f"sort_experiences order: {debug_view}")

    return [t[2] for t in indexed]
