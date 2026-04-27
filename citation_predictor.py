from anthropic_client import client

def suggest_citations(claim, field):
    prompt = f"""Given the claim: "{claim}" in the field of {field},
suggest 2-3 highly relevant, real academic citations (full reference), 
and explain why each supports the claim. If you cannot recall a real paper, 
state clearly you are using a placeholder and suggest search terms."""
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}]
    )
    return resp.content[0].text