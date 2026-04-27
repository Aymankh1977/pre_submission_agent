from anthropic_client import client

def polish_section(manuscript, section_name):
    """Rewrite a given section to improve clarity and style."""
    prompt = f"""Polish the following section from a scientific manuscript to be more 
clear, concise, and engaging for a top journal. Preserve all scientific meaning.
After the rewritten text, list the changes you made and why.

Section ({section_name}):
{manuscript}"""
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}]
    )
    return resp.content[0].text