# experience_parser.py
import re

# Broaden heading variants substantially
EXPERIENCE_HEADINGS = {
    "experience",
    "professional experience",
    "work experience",
    "employment history",
    "career history",
    "relevant experience",
    "[experience]",
}
STOP_HEADINGS = {
    "education",
    "certifications",
    "skills",
    "summary",
    "projects",
    "publications",
    "[education]",
    "[certifications]",
    "[skills]",
    "[summary]",
    "[projects]",
    "[publications]",
}

MARKER_START = "=== Experience ==="
# We'll stop at the next "=== <Section> ===" marker if present.


def _slice_between_markers(text: str):
    """Return lines between '=== Experience ===' and the next '=== ... ===' marker, if present."""
    if not text:
        return []

    start_idx = text.find(MARKER_START)
    if start_idx == -1:
        return []

    # Find the next marker after the Experience one
    next_marker = None
    for m in re.finditer(r"===\s+[A-Za-z ]+\s+===", text):
        if m.start() > start_idx:
            next_marker = m.start()
            break

    if next_marker is None:
        chunk = text[start_idx + len(MARKER_START):]
    else:
        chunk = text[start_idx + len(MARKER_START): next_marker]

    lines = [ln.strip() for ln in chunk.splitlines() if ln.strip()]
    return lines


def _slice_by_headings(text: str):
    """Fallback: use broad set of headings (case/brace-insensitive)."""
    if not text:
        return []

    lines = [ln.strip() for ln in text.splitlines()]

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

    return [ln for ln in lines[start:end] if ln]


def extract_experience_lines(full_text: str):
    """
    Returns lines from the Experience section verbatim (no summarisation).

    Strategy:
      1) Prefer marker-based slice if '=== Experience ===' exists (aligned to LLM prompt).
      2) Otherwise, fall back to heading-based slice with broad variants.
    """
    lines = _slice_between_markers(full_text)
    if lines:
        return lines
    return _slice_by_headings(full_text)
