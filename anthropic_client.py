import os
from dotenv import load_dotenv
import anthropic

# Load .env with override to ensure fresh values
load_dotenv(override=True)

api_key = os.getenv("ANTHROPIC_API_KEY")

if not api_key:
    raise ValueError("ANTHROPIC_API_KEY not found in .env file")

# Strip any accidental whitespace or quotes
api_key = api_key.strip().strip('"').strip("'")

client = anthropic.Anthropic(api_key=api_key)