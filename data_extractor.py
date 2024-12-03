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

**IMPORTANT INSTRUCTIONS:**

- **OUTPUT ONLY THE JSON OBJECT:** Do not include any explanations, notes, or extra text.
- **NO CODE BLOCKS:** Do not wrap the JSON in backticks or code blocks.
- **VALID JSON ONLY:** Ensure the JSON is valid and can be parsed by standard JSON parsers.
- **USE DOUBLE QUOTES:** Use double quotes for all keys and string values.
- **NO EXTRA SPACES:** Avoid unnecessary whitespace within the JSON.
- **EXACT TEXT:** Copy the text exactly as it appears in the CV for each field.

**TASK:**

Extract the following information from the CV text provided and return it in valid JSON format. For each item, copy the text exactly as it appears in the CV.

- **ApplicantName**: Copy the exact name of the applicant as it appears in the CV.
- **Role**: Copy the exact role as stated.
- **SecurityClearance**: Copy the exact security clearance details.
- **Summary**: Copy the summary exactly as it appears.
- **Skills**: Copy the skills exactly as listed.

**CV TEXT:**
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

**IMPORTANT INSTRUCTIONS:**

- **OUTPUT ONLY THE JSON OBJECT:** Do not include any explanations, notes, or extra text.
- **NO CODE BLOCKS:** Do not wrap the JSON in backticks or code blocks.
- **VALID JSON ONLY:** Ensure the JSON is valid and can be parsed by standard JSON parsers.
- **USE DOUBLE QUOTES:** Use double quotes for all keys and string values.
- **NO EXTRA SPACES:** Avoid unnecessary whitespace within the JSON.
- **EXACT TEXT:** Copy the text exactly as it appears in the CV for each field.

**TASK:**

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

**CV TEXT:**
{text}
"""


    response_text = call_openai_api(prompt_experience, max_tokens=3000, call_type="Experience & Education Extraction")
    data_experience = parse_json_response(response_text)
    return data_experience


def call_openai_api(prompt, max_tokens, call_type=""):
    """
    Calls the OpenAI API with the given prompt and returns the assistant's response text.
    
    Parameters:
    - prompt (str): The prompt to send to the OpenAI API.
    - max_tokens (int): The maximum number of tokens to generate.
    - call_type (str): An optional identifier for the type of call (e.g., 'Basic Info', 'Experience').
    """
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",  # Use your valid model name
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

        # Remove any backticks from the response
        response_text = response_text.replace('`', '')

        return response_text

    except openai.error.RateLimitError as e:
        logging.error(f"Rate limit exceeded: {e}")
        raise
    except openai.error.InvalidRequestError as e:
        logging.error(f"Invalid request: {e}")
        raise
    except openai.error.AuthenticationError as e:
        logging.error(f"Authentication error: {e}")
        raise
    except openai.error.OpenAIError as e:
        logging.error(f"OpenAI API error: {e}")
        raise

def parse_json_response(response_text):
    """
    Parses the assistant's response text as JSON.
    """
    try:
        data = json.loads(response_text)
        return data
    except json.JSONDecodeError:
        # Attempt to extract JSON using regex
        json_str = extract_json(response_text)
        try:
            data = json.loads(json_str)
            return data
        except json.JSONDecodeError as e:
            logging.error(f"JSON decoding failed: {e}")
            logging.error(f"Problematic JSON string: {json_str}")
            raise

def extract_json(text):
    """
    Extracts the JSON object from the text, cleans it, and returns it as a dictionary.
    """
    import json

    # Remove any markdown code block markers or leading/trailing whitespace
    text = text.strip().strip('`').strip()

    # Use a regular expression to find the first JSON object in the text
    json_match = re.search(r'\{.*\}', text, re.DOTALL)
    if not json_match:
        logging.error("No JSON object found in the assistant's response.")
        raise ValueError("Extracted string is not valid JSON.")

    json_str = json_match.group(0)

    # Clean the JSON string
    json_str = clean_json_string(json_str)

    # Attempt to parse the JSON string
    try:
        data = json.loads(json_str)
        return data
    except json.JSONDecodeError as e:
        logging.error(f"JSON decoding failed: {e}")
        logging.error(f"Problematic JSON string: {json_str}")
        raise ValueError("Extracted string is not valid JSON.")

def clean_json_string(json_str):
    """
    Cleans common issues in a JSON string to make it parseable.
    """
    # Replace single quotes with double quotes
    json_str = json_str.replace("'", '"')

    # Remove any trailing commas before closing braces/brackets
    json_str = re.sub(r',\s*(\}|])', r'\1', json_str)

    # Remove newlines and tabs
    json_str = json_str.replace('\n', '').replace('\t', '')

    # Ensure proper escaping of double quotes inside strings
    json_str = re.sub(r'\\(?=")', r'', json_str)

    # Remove any backslashes not used for escaping
    json_str = json_str.replace('\\', '')

    # Remove extra whitespace between keys and colons
    json_str = re.sub(r'\s*:\s*', ':', json_str)

    # Remove extra whitespace after commas
    json_str = re.sub(r',\s*', ',', json_str)

    return json_str
