# config.py

import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Retrieve the OpenAI API key from environment variables
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Optionally, you can handle the case where the API key is not set
if OPENAI_API_KEY is None:
    raise ValueError("OpenAI API key not found. Please set OPENAI_API_KEY in your .env file.")
