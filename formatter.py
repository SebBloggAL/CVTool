# formatter.py

import logging

def format_data(raw_data):
    """
    Formats the raw data extracted from the CV to match the template requirements.
    """
    # Retrieve Security Clearance and handle null or empty values
    security_clearance = raw_data.get("Security Clearance", "Not specified")
    if not security_clearance:
        security_clearance = "Not specified"
    
    # Use .get() with default values to handle missing fields
    formatted_data = {
        "ApplicantName": raw_data.get("Applicant's Name", "Name not provided"),
        "Role": raw_data.get("Role", "Role not specified"),
        "SecurityClearance": security_clearance,
        "Summary": raw_data.get("Summary", "Summary not provided"),
        "Skills": format_skills(raw_data.get("Skills", [])),
        "Experience": format_experience(raw_data.get("Experience", [])),
        "Education": format_education(raw_data.get("Education", []))
    }
    logging.debug(f"Formatted data: {formatted_data}")
    return formatted_data

def format_skills(skills_data):
    """
    Formats the skills data, returning it as a list of skills.
    """
    if isinstance(skills_data, list):
        return [skill.strip() for skill in skills_data if skill.strip()]
    elif isinstance(skills_data, str):
        # Split the string by commas or newlines and strip whitespace
        skills_list = [skill.strip() for skill in skills_data.replace('\n', ',').split(',') if skill.strip()]
        return skills_list
    else:
        logging.warning("Unexpected format for skills data.")
        return []



def format_experience(experience_data):
    """
    Formats the experience data, returning it as a list of dictionaries.
    Each dictionary contains the keys: Position, Company, Duration, Responsibilities.
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
        # If experience_data is a single dictionary, wrap it in a list
        formatted_item = {
            "Position": experience_data.get("Position", ""),
            "Company": experience_data.get("Company", ""),
            "Duration": experience_data.get("Duration", ""),
            "Responsibilities": experience_data.get("Responsibilities", [])
        }
        formatted_experiences.append(formatted_item)
    else:
        # Handle other cases or return an empty list
        logging.warning("Unexpected format for experience data.")
    return formatted_experiences


def format_education(education_data):
    """
    Formats the education data, only including Institution and Duration if they are specified.
    """
    formatted_education = ""
    if isinstance(education_data, list):
        for item in education_data:
            degree = item.get("Degree", "")
            institution = item.get("Institution", "")
            duration = item.get("Duration", "")
            
            # Initialize the entry with Degree
            entry = f"{degree}"
            
            # Append Institution if it's specified and not "Not specified"
            if institution and institution.lower() != "not specified":
                entry += f" at {institution}"
            
            # Append Duration if it's specified and not "Not specified"
            if duration and duration.lower() != "not specified":
                entry += f" ({duration})"
            
            # Add the formatted entry to the overall formatted education string
            formatted_education += f"{entry}\n\n"
    elif isinstance(education_data, dict):
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

