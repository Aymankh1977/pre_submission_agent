from anthropic_client import client

SYSTEM = """You are a senior scientist and writer. Expand the user's notes into 
publishable scientific text. Use real citations when possible; if not, give a placeholder 
and explain how to find it. After the text, list all suggested citations with justification.
Be transparent about your choices."""

def generate(user_prompt):
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=3000,
        system=SYSTEM,
        messages=[{"role": "user", "content": user_prompt}]
    )
    return resp.content[0].text