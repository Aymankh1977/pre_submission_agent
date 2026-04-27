from anthropic_client import client

def generate_cover_letter(title, abstract, highlights, journal_name):
    prompt = f"""Draft a cover letter for the editor of {journal_name}.
Manuscript title: {title}
Abstract: {abstract}
Key highlights provided by author: {highlights}
The letter should be formal, persuasive, and include a statement of novelty.
After the letter, explain the reasoning behind the phrasing."""
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}]
    )
    return resp.content[0].text