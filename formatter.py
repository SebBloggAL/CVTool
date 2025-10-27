# formatter.py

import re
import logging
from datetime import datetime
from typing import List, Dict, Any

# Reuse the date parsing already defined in document_generator to avoid duplication.
# (document_generator does NOT import formatter, so this won't create a circular import.)
from document_generator import parse_end_date


def format_data(raw_data):
    """
    Formats the raw data extracted from the CV to match the template requirements.
    """
    # Handle Security Clearances
    security_clearance = raw_data.get("SecurityClearance", "Not specified")
    if not security_clearance:
        security_clearance = "Not specified"

    formatted_data = {
        "ApplicantName": raw_data.get("ApplicantName", "Name not provided"),
        "Role": raw_data.get("Role", "Role not specified"),
        "SecurityClearance": security_clearance,
        "Summary": raw_data.get("Summary", "Summary not provided"),
        "Skills": format_skills(raw_data.get("Skills", [])),
        "Experience": format_experience(raw_data.get("Experience", [])),  # sorted newest->oldest
        "Education": format_education(raw_data.get("Education", [])),
        "Certifications": format_certifications(raw_data.get("Certifications", []))
    }

    logging.debug(f"Formatted data: {formatted_data}")
    return formatted_data


def format_skills(skills_data):
    """
    Formats the skills data into a list of individual skills.
    """
    if isinstance(skills_data, list):
        return [skill.strip() for skill in skills_data if isinstance(skill, str) and skill.strip()]
    elif isinstance(skills_data, str):
        # Split on commas or newlines
        return [skill.strip() for skill in re.split(r'[\n,]', skills_data) if skill.strip()]
    else:
        logging.warning("Unexpected format for skills data.")
        return []


def format_experience(experience_data):
    """
    Keep responsibilities verbatim; sort roles by parsed end date (newest first).
    Does NOT paraphrase/abbreviate any text.
    """
    formatted_experiences: List[Dict[str, Any]] = []

    if isinstance(experience_data, list):
        for item in experience_data:
            if not isinstance(item, dict):
                logging.warning(f"Unexpected experience item type: {type(item)}")
                continue
            formatted_item = {
                "Position": item.get("Position", ""),
                "Company": item.get("Company", ""),
                "Duration": item.get("Duration", ""),
                # IMPORTANT: keep bullets exactly as provided (no summarisation)
                "Responsibilities": item.get("Responsibilities", []),
            }
            formatted_experiences.append(formatted_item)

    elif isinstance(experience_data, dict):
        formatted_experiences.append({
            "Position": experience_data.get("Position", ""),
            "Company": experience_data.get("Company", ""),
            "Duration": experience_data.get("Duration", ""),
            "Responsibilities": experience_data.get("Responsibilities", []),
        })
    else:
        logging.warning("Unexpected format for experience data.")

    # Sort newest -> oldest, but preserve original order for ties/unknown dates
    try:
        formatted_experiences = sort_experiences(formatted_experiences)
    except Exception as e:
        logging.warning(f"Could not sort experiences: {e}")

    return formatted_experiences


def format_education(education_data):
    """
    Formats the education data as a single string with entries separated by two line breaks.
    """
    formatted_education = ""
    if isinstance(education_data, list):
        for item in education_data:
            if not isinstance(item, dict):
                continue
            degree = item.get("Degree", "")
            institution = item.get("Institution", "")
            duration = item.get("Duration", "")

            entry = f"{degree}".strip()
            if institution and institution.lower() != "not specified":
                entry += f" at {institution}"
            if duration and duration.lower() != "not specified":
                entry += f" ({duration})"

            entry = entry.strip()
            if entry:
                formatted_education += f"{entry}\n\n"

    elif isinstance(education_data, dict):
        # Fallback: flatten a dict (unlikely if prompt is followed)
        for key, value in education_data.items():
            formatted_education += f"{key}\n"
            if isinstance(value, list):
                for detail in value:
                    formatted_education += f"- {detail}\n"
            elif isinstance(value, str):
                formatted_education += f"- {value}\n"
            formatted_education += "\n"

    elif isinstance(education_data, str):
        formatted_education += education_data

    return formatted_education.strip()


def format_certifications(cert_data):
    """
    Formats the certifications data as a list of strings.
    """
    if isinstance(cert_data, list):
        return [item.strip() for item in cert_data if isinstance(item, str) and item.strip()]
    elif isinstance(cert_data, str):
        # Split on commas or newlines
        return [s.strip() for s in re.split(r'[\n,]', cert_data) if s.strip()]
    else:
        logging.warning("Unexpected format for certifications data.")
        return []


def sort_experiences(experience_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Sorts a list of experience objects by their parsed end date, newest first,
    using document_generator.parse_end_date. Preserves original order for ties or
    invalid/unknown dates.
    """
    # Capture original index to keep sort stable on ties
    indexed = []
    for idx, item in enumerate(experience_data):
        duration = item.get("Duration", "")
        try:
            end_dt = parse_end_date(duration) if duration is not None else datetime.min
        except Exception:
            end_dt = datetime.min
        indexed.append((idx, end_dt, item))

    # Sort by end date DESC, then by original index ASC
    indexed.sort(key=lambda tup: (tup[1], -tup[0]), reverse=True)

    # Rebuild list in sorted order
    return [it[2] for it in indexed]
