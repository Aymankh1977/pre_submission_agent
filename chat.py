from anthropic_client import client

SYSTEM_TEMPLATE = """You are an expert scientific editor. The full manuscript is below.
Answer questions, suggest improvements, rewrite sections, propose citations.
Always explain your reasoning and reference exact parts of the text.
Manuscript:
---
{manuscript}
---"""

def chat(manuscript, user_msg, history=None):
    messages = history.copy() if history else []
    messages.append({"role": "user", "content": user_msg})
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4000,
        system=SYSTEM_TEMPLATE.format(manuscript=manuscript),
        messages=messages
    )
    return resp.content[0].text