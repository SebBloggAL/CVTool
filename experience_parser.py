# experience_parser.py
import re
import logging
from section_config import SECTION_SYNONYMS, STOP_HEADINGS

MARKER = "=== Experience ==="
MARKER_RE = re.compile(r"===\s+[A-Za-z ]+\s+===")

def _slice_between_markers(text: str):
    if not text:
        return []
    start = text.find(MARKER)
    if start == -1:
        return []
    # find next marker
    next_pos = None
    for m in MARKER_RE.finditer(text):
        if m.start() > start:
            next_pos = m.start()
            break
    chunk = text[start + len(MARKER):] if next_pos is None else text[start + len(MARKER): next_pos]
    lines = [ln.strip() for ln in chunk.splitlines() if ln.strip()]
    # Early stop if we see a STOP heading inside
    for idx, ln in enumerate(lines):
        t = re.sub(r'[\[\]]', '', ln).strip().lower()
        if t in {s.lower() for s in STOP_HEADINGS}:
            lines = lines[:idx]
            break
    return lines

def _slice_by_headings(text: str):
    if not text:
        return []
    lines = [ln.strip() for ln in text.splitlines()]
    exp_variants = {v for v in SECTION_SYNONYMS.get("Experience", [])}
    # Find start
    start = None
    for i, ln in enumerate(lines):
        t = re.sub(r'[\[\]]', '', ln).strip()
        for pat in exp_variants:
            if re.fullmatch(pat, t, flags=re.IGNORECASE):
                start = i + 1
                break
        if start is not None:
            break
    if start is None:
        return []
    # Find end
    end = len(lines)
    stop_set = {s.lower() for s in STOP_HEADINGS}
    for j in range(start, len(lines)):
        t = re.sub(r'[\[\]]', '', lines[j]).strip().lower()
        if t in stop_set:
            end = j
            break
    return [ln for ln in lines[start:end] if ln]

def extract_experience_lines(text: str):
    """
    Returns lines from the Experience section verbatim (no summarisation).
    """
    lines = _slice_between_markers(text)
    if lines:
        logging.debug(f"[experience_parser] Marker-based Experience lines: {len(lines)}")
        return lines
    lines = _slice_by_headings(text)
    logging.debug(f"[experience_parser] Heading-based Experience lines: {len(lines)}")
    return lines
