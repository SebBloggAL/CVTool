import openai
from config import OPENAI_API_KEY  # Import the API key
from datetime import datetime
import re
import json
import logging
import threading

# Set the OpenAI API key
openai.api_key = OPENAI_API_KEY

# Initialize a lock for thread-safe file writing
file_lock = threading.Lock()

def extract_cv_data(text):
    """
    Uses the OpenAI API to extract structured data from the CV text.
    We first inject explicit section markers so the LLM output is more reliable.
    """
    # 1) Insert consistent section markers
    sections = ["Summary", "Skills", "Experience", "Education"]
    for sec in sections:
        # match a heading on its own line, case-insensitive
        text = re.sub(
            rf'\n\s*{sec}\s*\n',
            f'\n=== {sec} ===\n',
            text,
            flags=re.IGNORECASE
        )

    # 2) Now extract JSON in two passes
    data_basic = extract_basic_info(text)
    data_experience = extract_experience_education(text)

    # Combine the results
    return {**data_basic, **data_experience}


def extract_basic_info(text):
    """
    Extracts basic information from the CV text.
    """
    prompt_basic = f"""
 You are an AI assistant that extracts specific information from resumes.

 IMPORTANT INSTRUCTIONS:
 1. Output ONLY one valid JSON object — no explanations or extra text.
 2. Do NOT wrap in code fences (` or ```).
 3. Must be fully valid JSON:
    - Balanced braces {{ }}.
    - Double quotes around keys and string values.
    - No trailing commas.
    - Every key:value pair must be separated by a comma.
 4. Use the exact text from the CV for each field.

 EXAMPLE STRUCTURE:
 {{
   "ApplicantName": "Jane Doe",
   "Role": "DevOps Engineer",
   "SecurityClearance": "TopSecret",
   "Summary": "Cloud infrastructure specialist…",
   "Skills": ["Terraform", "Kubernetes"]
 }}

 TASK:
 Extract the following from the CV text, copying the text exactly as it appears:
 - ApplicantName
 - Role
 - SecurityClearance
 - Summary
 - Skills

 CV TEXT:
 {text}
 """

    response_text = call_openai_api(
        prompt_basic,
        max_tokens=1500,
        call_type="Basic Info Extraction"
    )
    return parse_json_response(response_text)


def extract_experience_education(text):
    """
    Extracts experience and education from the CV text.
    """
    prompt_experience = f"""
You are an AI assistant extracting from a CV with explicit markers.

OUTPUT a single JSON object — no explanations, no code fences.

Use everything between "=== Experience ===" and "=== Education ===" as "Experience".
Use everything after "=== Education ===" as "Education".

STRUCTURE:
{{
  "Experience": [ {{ 
      "Position": "...", 
      "Company": "...", 
      "Duration": "...", 
      "Responsibilities": [...], 
      "TechnologiesUsed": "..." 
  }} ],
  "Education": [ {{ 
      "Degree": "...", 
      "Institution": "..." 
  }} ]
}}

CV TEXT:
{text}
"""

    response_text = call_openai_api(
        prompt_experience,
        max_tokens=3000,
        call_type="Experience & Education Extraction"
    )
    return parse_json_response(response_text)


def call_openai_api(prompt, max_tokens, call_type=""):
    """
    Calls OpenAI and logs the raw response.
    """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=max_tokens,
            temperature=0
        )
        response_text = response.choices[0].message["content"].strip()
        logging.debug(f"=== {call_type} RAW RESPONSE ===\n{response_text}\n")

        # Timestamped logging
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry = f"\n\n=== {call_type} [{timestamp}] ===\n{response_text}"
        with file_lock, open("assistant_response.txt", "a", encoding="utf-8") as f:
            f.write(entry)

        return response_text.replace('`', '')

    except Exception as e:
        logging.error(f"OpenAI API error in {call_type}: {e}", exc_info=True)
        raise


def parse_json_response(response_text):
    """
    Parses the assistant's response text as JSON, with a fallback extractor.
    """
    try:
        return json.loads(response_text)
    except json.JSONDecodeError:
        json_str = extract_json(response_text)
        return json.loads(json_str)


def extract_json(text):
    """
    Grabs the first JSON object in the text.
    """
    # Strip code fences and whitespace
    clean = text.strip().strip('`')
    start = clean.find('{')
    end = clean.rfind('}')
    if start == -1 or end == -1:
        logging.error("No JSON object found.")
        raise ValueError("Invalid JSON response.")
    return clean[start : end+1]


def clean_json_string(json_str):
    """
    Performs regex clean-up on a JSON-ish string.
    """
    import re

    # Straighten quotes and remove trailing commas
    json_str = (
        json_str
        .replace("'", '"')
        .replace("“", '"').replace("”", '"')
        .replace("–", "-").replace("—", "-")
    )
    json_str = re.sub(r',\s*(\}|\])', r'\1', json_str)
    json_str = re.sub(r'\s+', ' ', json_str).strip()
    return json_str
