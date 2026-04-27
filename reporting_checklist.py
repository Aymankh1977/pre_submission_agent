from anthropic_client import client

def check_guideline(manuscript, guideline_name):
    """Prompts Claude to assess adherence to a reporting guideline."""
    prompt = f"""Check if the following manuscript adheres to the {guideline_name} reporting guideline.
Output a checklist with each item marked as Yes/No/Partial, and explain the reasoning for each.
Manuscript:
{manuscript[:8000]}"""
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2500,
        messages=[{"role": "user", "content": prompt}]
    )
    return resp.content[0].text