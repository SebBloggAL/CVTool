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

def format_skills(skills_list):
    if isinstance(skills_list, list):
        return ', '.join(skills_list)
    return skills_list  # In case it's already a string

def format_experience(experience_data):
    formatted_experience = ""
    if isinstance(experience_data, list):
        for item in experience_data:
            position = item.get("Position", "")
            company = item.get("Company", "")
            duration = item.get("Duration", "")
            responsibilities = item.get("Responsibilities", [])
            formatted_experience += f"{position} at {company} ({duration})\n"
            if isinstance(responsibilities, list):
                for responsibility in responsibilities:
                    formatted_experience += f"{responsibility}\n"  
            elif isinstance(responsibilities, str):
                formatted_experience += f"{responsibilities}\n"  
            formatted_experience += "\n"
    elif isinstance(experience_data, dict):
        for key, value in experience_data.items():
            formatted_experience += f"{key}\n"
            if isinstance(value, list):
                for responsibility in value:
                    formatted_experience += f"- {responsibility}\n"
            elif isinstance(value, str):
                formatted_experience += f"- {value}\n"
            formatted_experience += "\n"
    elif isinstance(experience_data, str):
        formatted_experience += experience_data
    return formatted_experience.strip()

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

