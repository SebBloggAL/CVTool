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
