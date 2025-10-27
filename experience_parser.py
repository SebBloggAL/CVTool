# experience_parser.py
import re
import logging

# Treat "Career Summary" as the Experience section heading too.
EXPERIENCE_HEADINGS = {
    "experience",
    "professional experience",
    "work experience",
    "employment history",
    "career history",
    "relevant experience",
    "career summary",            # <-- NEW
    "[experience]",
}

# Make sure we stop before skills of any kind (incl. "Technical Skills")
STOP_HEADINGS = {
    "education",
    "certifications",
    "skills",
    "technical skills",          # <-- NEW
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
MARKER_PATTERN = re.compile(r"===\s+[A-Za-z ]+\s+===")


def _first_stop_index(lines):
    """Return index of first line that looks like a stop heading, else None."""
    for idx, ln in enumerate(lines):
        t = re.sub(r'[\[\]]', '', ln).strip().lower()
        if t in STOP_HEADINGS:
            return idx
    return None


def _slice_between_markers(text: str):
    """
    Return lines between '=== Experience ===' and the next '=== ... ===' marker.
    Additionally, if a STOP heading (e.g., 'Technical Skills') appears before the next marker,
    stop there to avoid pulling Skills into Experience.
    """
    if not text:
        return []

    start_idx = text.find(MARKER_START)
    if start_idx == -1:
        return []

    # Next explicit marker after Experience
    next_marker = None
    for m in MARKER_PATTERN.finditer(text):
        if m.start() > start_idx:
            next_marker = m.start()
            break

    chunk = text[start_idx + len(MARKER_START):] if next_marker is None else text[start_idx + len(MARKER_START): next_marker]
    lines = [ln.strip() for ln in chunk.splitlines() if ln.strip()]

    # Early stop inside the chunk if we see a STOP heading like "Technical Skills"
    stop_at = _first_stop_index(lines)
    if stop_at is not None:
        lines = lines[:stop_at]

    return lines


def _slice_by_headings(text: str):
    """Fallback: detect by broad headings without markers."""
    if not text:
        return []

    lines = [ln.strip() for ln in text.splitlines()]

    # Find start (now includes "career summary")
    start = None
    for i, ln in enumerate(lines):
        t = re.sub(r'[\[\]]', '', ln).strip().lower()
        if t in EXPERIENCE_HEADINGS:
            start = i + 1
            break
    if start is None:
        return []

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
      1) Prefer marker-based slice if '=== Experience ===' exists.
      2) Otherwise, fall back to heading-based slice with broad variants.
    """
    lines = _slice_between_markers(full_text)
    if lines:
        logging.debug(f"[experience_parser] Marker-based Experience lines: {len(lines)}")
        return lines

    lines = _slice_by_headings(full_text)
    logging.debug(f"[experience_parser] Heading-based Experience lines: {len(lines)}")
    return lines
