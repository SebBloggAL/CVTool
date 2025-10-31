# section_config.py
SECTION_SYNONYMS = {
    "Summary": [
        r"\[?\s*Summary\s*\]?",
        r"Profile",
        r"Professional\s+Summary",
        r"Summary\s+Profile",
        r"Career\s+Objective"
    ],
    "Skills": [
        r"\[?\s*Skills\s*\]?",
        r"Technical\s+Skills",
        r"Core\s+Skills",
        r"Key\s+Skills",
        r"Skills\s*/\s*Capabilities"
    ],
    "Experience": [
        r"\[?\s*Experience\s*\]?",
        r"Professional\s+Experience",
        r"Work\s+Experience",
        r"Employment\s+History",
        r"Career\s+History",
        r"Relevant\s+Experience",
        r"Recent\s+Work\s+Experience",
        r"Career\s+Summary"           # Emmanuel & Simon pattern
    ],
    "Education": [ r"\[?\s*Education\s*\]?" ],
    "Certifications": [
        r"\[?\s*Certifications\s*\]?",
        r"Qualifications",
        r"Certificates",
        r"Additional\s+Experience\s*&\s+Qualifications"  # capture extra section
    ],
    "Interests": [ r"Interests" ],
    "Additional": [
        r"Additional\s+Experience\s*&\s+Qualifications",
        r"Additional\s+Information"
    ]
}
STOP_HEADINGS = {
    # include ‘Technical Skills’ and other sections so we don’t bleed into Skills
    "Skills", "Technical Skills", "Education", "Certifications", "Summary",
    "Projects", "Publications", "Interests", "Additional Experience & Qualifications"
}
