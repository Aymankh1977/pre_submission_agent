import pymupdf
from docx import Document
import re
from sentence_transformers import SentenceTransformer

model = SentenceTransformer('sentence-transformers/allenai-specter')

def extract_text(uploaded_file):
    if uploaded_file.name.endswith('.pdf'):
        doc = pymupdf.open(stream=uploaded_file.read(), filetype='pdf')
        return '\n'.join(page.get_text() for page in doc)
    elif uploaded_file.name.endswith('.docx'):
        doc = Document(uploaded_file)
        return '\n'.join(para.text for para in doc.paragraphs)
    else:
        raise ValueError('Only PDF or DOCX')

def extract_title_abstract(text):
    """Heuristic to get title (first non-empty line) and abstract."""
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    title = lines[0] if lines else 'Untitled'
    abstract = ''
    for i, line in enumerate(lines):
        if re.search(r'(?i)^\s*abstract\s*$', line):
            abstract = ' '.join(lines[i+1:i+10])
            break
    # Fallback: take first 500 chars after title if no abstract marker
    if not abstract and len(lines) > 1:
        abstract = ' '.join(lines[1:])[:500]
    return title, abstract

def get_embedding(txt):
    return model.encode(txt)

def cosine_sim(a, b):
    from sklearn.metrics.pairwise import cosine_similarity
    return cosine_similarity([a], [b])[0][0]