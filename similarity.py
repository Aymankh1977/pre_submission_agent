import requests
from manuscript_utils import get_embedding, cosine_sim
from anthropic_client import client  # Use shared client

def search_semantic_scholar(query, limit=50):
    url = "https://api.semanticscholar.org/graph/v1/paper/search"
    params = {"query": query, "limit": limit, "fields": "title,abstract,year,externalIds,url"}
    r = requests.get(url, params=params)
    if r.status_code == 200:
        return r.json().get('data', [])
    return []

def find_similar(manuscript_title, manuscript_abstract, top_k=10):
    query = manuscript_title
    candidates = search_semantic_scholar(query, limit=50)
    if not candidates:
        return []
    man_emb = get_embedding(manuscript_title + ' ' + (manuscript_abstract or ''))
    scored = []
    for paper in candidates:
        if not paper.get('abstract'):
            continue
        paper_txt = paper['title'] + ' ' + paper['abstract']
        sim = cosine_sim(man_emb, get_embedding(paper_txt))
        scored.append((sim, paper))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[:top_k]

def overlap_report(title, abstract, papers):
    """Generate an editorial overlap analysis using Claude."""
    paper_list = "\n".join(
        f"- Title: {p['title']}\n  Abstract: {p.get('abstract','')}"
        for _, p in papers
    )
    prompt = f"""Assess novelty of the manuscript:
Title: {title}
Abstract: {abstract}

Potentially similar papers:
{paper_list}

For each paper: explain overlap and differences. Rate rejection risk (Low/Medium/High).
Provide overall recommendation and reasoning. Be transparent and specific."""

    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}]
    )
    return resp.content[0].text