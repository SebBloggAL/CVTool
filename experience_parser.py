# experience_parser.py
import re

EXPERIENCE_HEADINGS = { "experience", "professional experience", "[experience]" }
STOP_HEADINGS = { "education", "certifications", "skills", "summary",
                  "[education]", "[certifications]", "[skills]", "[summary]" }

def extract_experience_lines(full_text: str):
    """
    Returns lines from the Experience section verbatim (no summarisation).
    """
    lines = [ln.strip() for ln in full_text.splitlines()]
    # find start
    start = None
    for i, ln in enumerate(lines):
        t = re.sub(r'[\[\]]', '', ln).strip().lower()
        if t in EXPERIENCE_HEADINGS:
            start = i + 1
            break
    if start is None:
        return []

    # find end
    end = len(lines)
    for j in range(start, len(lines)):
        t = re.sub(r'[\[\]]', '', lines[j]).strip().lower()
        if t in STOP_HEADINGS:
            end = j
            break

    # Keep non-empty lines; do not strip bullets
    kept = [ln for ln in lines[start:end] if ln]
    return kept
