#formatter.py

import re
import logging

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
        "Experience": format_experience(raw_data.get("Experience", [])),
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
        return [skill.strip() for skill in skills_data if skill.strip()]
    elif isinstance(skills_data, str):
        # Split on commas or newlines
        return [skill.strip() for skill in re.split(r'[\n,]', skills_data) if skill.strip()]
    else:
        logging.warning("Unexpected format for skills data.")
        return []


def format_experience(experience_data):
    """
    Formats the experience data into a list of dictionaries, each with Position, Company, Duration, Responsibilities.
    """
    formatted_experiences = []
    if isinstance(experience_data, list):
        for item in experience_data:
            formatted_item = {
                "Position": item.get("Position", ""),
                "Company": item.get("Company", ""),
                "Duration": item.get("Duration", ""),
                "Responsibilities": item.get("Responsibilities", [])
            }
            formatted_experiences.append(formatted_item)
    elif isinstance(experience_data, dict):
        # Single object → wrap in list
        formatted_item = {
            "Position": experience_data.get("Position", ""),
            "Company": experience_data.get("Company", ""),
            "Duration": experience_data.get("Duration", ""),
            "Responsibilities": experience_data.get("Responsibilities", [])
        }
        formatted_experiences.append(formatted_item)
    else:
        logging.warning("Unexpected format for experience data.")
    return formatted_experiences


def format_education(education_data):
    """
    Formats the education data as a single string with entries separated by two line breaks.
    """
    formatted_education = ""
    if isinstance(education_data, list):
        for item in education_data:
            degree = item.get("Degree", "")
            institution = item.get("Institution", "")
            duration = item.get("Duration", "")

            entry = f"{degree}"
            if institution and institution.lower() != "not specified":
                entry += f" at {institution}"
            if duration and duration.lower() != "not specified":
                entry += f" ({duration})"

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
        return [item.strip() for item in cert_data if item.strip()]
    elif isinstance(cert_data, str):
        # Split on commas or newlines
        return [s.strip() for s in re.split(r'[\n,]', cert_data) if s.strip()]
    else:
        logging.warning("Unexpected format for certifications data.")
        return []
