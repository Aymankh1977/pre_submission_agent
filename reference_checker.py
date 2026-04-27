import requests, re

CROSSREF_API = "https://api.crossref.org/works/"
RETRACTION_DB = "https://api.labs.crossref.org/queries/retraction/"

def extract_dois(text):
    """Find all DOIs in the manuscript."""
    doi_pattern = r'\b10\.\d{4,}/[^\s]+'
    return list(set(re.findall(doi_pattern, text)))

def validate_doi(doi):
    """Return True if DOI is valid and not retracted."""
    try:
        r = requests.get(CROSSREF_API + doi)
        valid = r.status_code == 200
        # Check retraction status (simplified)
        retracted = False
        # In real code, query retraction database. Here a placeholder.
        return {"doi": doi, "valid": valid, "retracted": retracted}
    except:
        return {"doi": doi, "valid": False, "retracted": False}

def reference_report(text):
    """Validate all DOIs in text and return a report."""
    dois = extract_dois(text)
    results = [validate_doi(d) for d in dois]
    return results