import os

# Retrieve the OpenAI API key from environment variables
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY")  # Ensure correct case-sensitivity

# Optionally, handle the case where the API key is not set
if not OPENAI_API_KEY:
    raise ValueError("OpenAI API key not found. Please set openai_api_key in your environment.")
