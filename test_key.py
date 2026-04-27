import os
from dotenv import load_dotenv
import anthropic

# Force load from the exact path
load_dotenv(override=True)

api_key = os.getenv("ANTHROPIC_API_KEY")
print(f"Key exists: {api_key is not None}")
print(f"Key length: {len(api_key) if api_key else 0}")
print(f"Key starts with: {api_key[:20] if api_key else 'NOTHING'}...")
print(f"Contains quotes: {'\"' in api_key if api_key else False}")
print(f"Contains spaces: {' ' in api_key if api_key else False}")

# Try a minimal API call
try:
    client = anthropic.Anthropic(api_key=api_key)
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=10,
        messages=[{"role": "user", "content": "Say hi"}]
    )
    print("✅ SUCCESS! Response:", resp.content[0].text)
except Exception as e:
    print(f"❌ FAILED: {e}")