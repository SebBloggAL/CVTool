import openai
from config import OPENAI_API_KEY
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
    """
    try:
        # First API call for basic information
        data_basic = extract_basic_info(text)

        # Second API call for experience and education
        data_experience = extract_experience_education(text)

        # Combine the results
        data = {**data_basic, **data_experience}
        return data

    except json.JSONDecodeError as json_error:
        logging.error(f"JSON decoding failed: {json_error}")
        logging.error(f"Response text was: {json_error.doc}")
        raise ValueError("Failed to decode JSON from OpenAI response.")
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        raise

def extract_basic_info(text):
    """
    Extracts basic information from the CV text.
    """
    prompt_basic = f"""
You are an AI assistant that extracts specific information from resumes.

**IMPORTANT INSTRUCTIONS**:

- **OUTPUT ONLY THE JSON OBJECT**: Do not include any explanations, notes, or extra text. Do not add any markdown formatting.
- **NO CODE BLOCKS**: Do not wrap the JSON in backticks or code blocks.
- **EXACT MATCHING**: Copy the text exactly as it appears in the CV. Do not paraphrase or summarize.
- **VALID JSON**: Ensure the JSON is valid and parsable by standard JSON parsers. Use double quotes for keys and string values.

**Task**:

Extract the following information from the CV text provided and return it in valid JSON format. For each item, copy the text exactly as it appears.

- **ApplicantName**: Copy the exact name as it appears.
- **Role**: Copy the exact role as stated.
- **SecurityClearance**: Copy the exact security clearance details.
- **Summary**: Copy the summary exactly as it appears.
- **Skills**: Copy the skills exactly as listed.

**CV Text**:
{text}
"""

    response_text = call_openai_api(prompt_basic, max_tokens=1500, call_type="Basic Info Extraction")
    data_basic = parse_json_response(response_text)
    return data_basic

def extract_experience_education(text):
    """
    Extracts experience and education information from the CV text.
    """
    prompt_experience = f"""
You are an AI assistant that extracts specific information from resumes.

**IMPORTANT INSTRUCTIONS**:

- **OUTPUT ONLY THE JSON OBJECT**: Do not include any explanations, notes, or extra text. Do not add any markdown formatting.
- **NO CODE BLOCKS**: Do not wrap the JSON in backticks or code blocks.
- **EXACT MATCHING**: Copy the text exactly as it appears in the CV. Do not paraphrase or summarize.
- **VALID JSON**: Ensure the JSON is valid and parsable by standard JSON parsers. Use double quotes for keys and string values.

**Task**:

Extract the following information from the CV text provided and return it in valid JSON format. For each item, copy the text exactly as it appears.

- **Experience** (as a list of work experiences, each with):
  - **Position**: Copy the exact job title.
  - **Company**: Copy the exact company name.
  - **Duration**: Copy the exact duration.
  - **Responsibilities**: Copy the exact responsibilities and achievements.
  - **TechnologiesUsed**: Copy the exact technologies mentioned.

- **Education** (as a list of educational qualifications, each with):
  - **Degree**: Copy the exact degree title.
  - **Institution**: Copy the exact institution name.

**CV Text**:
{text}
"""

    response_text = call_openai_api(prompt_experience, max_tokens=3000, call_type="Experience & Education Extraction")
    data_experience = parse_json_response(response_text)
    return data_experience

def call_openai_api(prompt, max_tokens, call_type=""):
    """
    Calls the OpenAI API with the given prompt and returns the assistant's response text.
    """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",  # Use your valid model name
            messages=[
                {"role": "user", "content": prompt}
            ],
            max_tokens=max_tokens,
            temperature=0
        )

        response_text = response.choices[0].message["content"].strip()
        logging.debug(f"Assistant's raw response:\n{response_text}\n")

        # Get current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Prepare the log entry
        if call_type:
            log_entry = f"\n\n=== {call_type} === [{timestamp}]\n{response_text}"
        else:
            log_entry = f"\n\n=== {timestamp} ===\n{response_text}"

        # Acquire the lock before writing to the file
        with file_lock:
            with open("assistant_response.txt", "a", encoding="utf-8") as f:
                f.write(log_entry)

        return response_text

    except openai.error.OpenAIError as e:
        logging.error(f"OpenAI API error: {e}")
        raise

def parse_json_response(response_text):
    """
    Parses the assistant's response text as JSON.
    """
    try:
        # Clean the response
        json_str = clean_assistant_response(response_text)
        # Parse the JSON
        data = json.loads(json_str)
        return data
    except json.JSONDecodeError as e:
        logging.error(f"JSON decoding failed: {e}")
        logging.error(f"Problematic JSON string: {json_str}")
        raise ValueError("Failed to decode JSON from assistant's response.")

def clean_assistant_response(response_text):
    """
    Cleans the assistant's response by removing any extraneous text before or after the JSON object.
    """
    # Find the first occurrence of '{' and the last occurrence of '}'
    start_idx = response_text.find('{')
    end_idx = response_text.rfind('}')
    if start_idx == -1 or end_idx == -1:
        logging.error("No JSON object found in the assistant's response.")
        raise ValueError("No JSON object found in the assistant's response.")
    # Extract the JSON substring
    json_str = response_text[start_idx:end_idx+1]
    return json_str
