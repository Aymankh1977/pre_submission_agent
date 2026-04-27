from manuscript_utils import get_embedding, cosine_sim
import numpy as np

JOURNALS = {
    "Nature": "Nature is a weekly international journal publishing the finest peer-reviewed research in all fields of science...",
    "Science": "Science publishes significant original scientific research...",
    "Cell": "Cell publishes findings of unusual significance in any area of experimental biology...",
    # ... add more
}

def score_journals(manuscript_abstract):
    man_emb = get_embedding(manuscript_abstract)
    scores = {}
    for name, aim in JOURNALS.items():
        j_emb = get_embedding(aim)
        scores[name] = cosine_sim(man_emb, j_emb)
    return sorted(scores.items(), key=lambda x: x[1], reverse=True)
