import streamlit as st
import os
import base64
import json
import hashlib
import datetime
import re
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from dotenv import load_dotenv
from anthropic import Anthropic, NotFoundError
from pypdf import PdfReader
import requests
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

try:
    from duckduckgo_search import DDGS
    DDG_AVAILABLE = True
except ImportError:
    DDG_AVAILABLE = False

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DentEdTech™ | Pre-Submission Agent",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── LOAD ENVIRONMENT & CLIENTS ───────────────────────────────────────────────
load_dotenv(override=True)
api_key = os.getenv("ANTHROPIC_API_KEY")
if not api_key:
    try:
        api_key = st.secrets.get("ANTHROPIC_API_KEY")
    except Exception:
        pass
if not api_key:
    st.error("🚨 ANTHROPIC_API_KEY is missing. Add it to `.env` or Streamlit Secrets.")
    st.stop()

client = Anthropic(api_key=api_key.strip().strip('"').strip("'"))

# Load embedding model for semantic similarity
@st.cache_resource
def load_embedding_model():
    return SentenceTransformer('sentence-transformers/allenai-specter')

embedding_model = load_embedding_model()

# ─── MODELS ───────────────────────────────────────────────────────────────────
PRIMARY_MODEL  = "claude-opus-4-5"
FALLBACK_MODEL = "claude-sonnet-4-5"
CHAT_MODEL     = "claude-haiku-4-5-20251001"

# ─── ACCESS CONTROL ───────────────────────────────────────────────────────────
VALID_ACCESS_CODES = {
    "DEMO-ACCESS":          "Demo / Internal Use",
    "DENTEDTECH-ADMIN":     "DentEdTech Admin",
    "DENTEDTECH-OWNER":     "DentEdTech Owner",
    "MEDED-PILOT-2025":     "Medical Education (ISI)",
    "ACADMED-PILOT-2025":   "Academic Medicine (ISI)",
    "MT-PILOT-2025":        "Medical Teacher (ISI)",
    "AHSE-PILOT-2025":      "Advances in Health Sciences Education (ISI)",
    "JDR-PILOT-2025":       "Journal of Dental Research (ISI)",
    "CDOE-PILOT-2025":      "Community Dentistry and Oral Epidemiology (ISI)",
    "EJOS-PILOT-2025":      "European Journal of Oral Sciences (ISI)",
    "JDS-PILOT-2025":       "Journal of Dental Sciences (ISI)",
    "BMC-PILOT-2025":       "BMC Medical Education (Scopus)",
    "JGME-PILOT-2025":      "JGME Journal of Graduate Medical Education (Scopus)",
    "TLM-PILOT-2025":       "Teaching and Learning in Medicine (Scopus)",
    "IJME-PILOT-2025":      "International Journal of Medical Education (Scopus)",
    "GMS-PILOT-2025":       "GMS Journal for Medical Education (Scopus)",
    "EFH-PILOT-2025":       "Education for Health (Scopus)",
    "JEEHP-PILOT-2025":     "Journal of Educational Evaluation for Health Professions (Scopus)",
    "MEP-PILOT-2025":       "MedEdPublish (Scopus)",
    "DENTJ-PILOT-2025":     "Dentistry Journal MDPI (Scopus)",
    "JDE-PILOT-2025":       "Journal of Dental Education (JDE)",
    "EJDE-PILOT-2025":      "European Journal of Dental Education (EJDE)",
    "BDJ-PILOT-2025":       "British Dental Journal (BDJ)",
    "JDENT-PILOT-2025":     "Journal of Dentistry",
    "ADEE-PILOT-2025":      "Dental Education Today (ADEE)",
    "JDHE-PILOT-2025":      "Journal of Dental Hygiene Education",
}

ADMIN_CODES = {"DEMO-ACCESS", "DENTEDTECH-ADMIN", "DENTEDTECH-OWNER"}
SESSION_ANALYSIS_CAP = 5

# ─── JOURNALS ─────────────────────────────────────────────────────────────────
JOURNALS = [
    "Medical Education (ISI)",
    "Academic Medicine (ISI)",
    "Medical Teacher (ISI)",
    "Advances in Health Sciences Education (ISI)",
    "Journal of Dental Research (ISI)",
    "Community Dentistry and Oral Epidemiology (ISI)",
    "European Journal of Oral Sciences (ISI)",
    "Journal of Dental Sciences (ISI)",
    "BMC Medical Education (Scopus)",
    "JGME – Journal of Graduate Medical Education (Scopus)",
    "Teaching and Learning in Medicine (Scopus)",
    "International Journal of Medical Education – IJME (Scopus)",
    "GMS Journal for Medical Education (Scopus)",
    "Education for Health (Scopus)",
    "Journal of Educational Evaluation for Health Professions (Scopus)",
    "MedEdPublish (Scopus)",
    "Dentistry Journal – MDPI (Scopus)",
    "Journal of Dental Education (JDE)",
    "European Journal of Dental Education (EJDE)",
    "British Dental Journal (BDJ)",
    "Journal of Dentistry",
    "Dental Education Today (ADEE)",
    "Journal of Dental Hygiene Education",
]

# Journal aims & scope for fit scoring
JOURNAL_AIMS = {
    "Medical Education (ISI)": "Medical Education seeks to be the pre-eminent journal in the field of education for health care professionals, and publishes material of the highest quality, reflecting worldwide or provocative issues and perspectives. The journal welcomes high quality papers on all aspects of health professional education including undergraduate, postgraduate and continuing education.",
    "Academic Medicine (ISI)": "Academic Medicine serves as an international forum for the exchange of ideas and information about policy, issues, and research concerning academic medicine, including strengthening the quality of medical education and training, enhancing the search for biomedical knowledge, advancing research in health services, and integrating education and research into the provision of effective health care.",
    "Journal of Dental Research (ISI)": "The Journal of Dental Research is dedicated to the dissemination of new knowledge and information on all sciences relevant to dentistry and to the oral cavity and associated structures in health and disease.",
    # Add more journal aims as needed
}

# ─── REVIEW CRITERIA ──────────────────────────────────────────────────────────
REVIEW_CRITERIA = {
    "research_question": "Research question clarity & PICO/SPIDER framing",
    "methodology":       "Methodology rigor & reproducibility",
    "consort_srqr":      "CONSORT / SRQR / COREQ guideline adherence",
    "kirkpatrick":       "Kirkpatrick level outcomes achieved",
    "citations":         "Citation currency, completeness & in-text accuracy",
    "statistics":        "Statistical / qualitative data analysis soundness",
    "ethics":            "Ethical considerations & positionality",
    "golden_thread":     "Golden thread coherence (RQ → method → results → conclusion)",
}

# ─── SESSION STATE ─────────────────────────────────────────────────────────────
APP_VERSION = "3.0.0"

defaults = {
    "access_granted":        False,
    "access_partner":        "",
    "is_admin":              False,
    "consent_given":         False,
    "analyses_this_session": 0,
    "pdf_base64":            None,
    "pdf_name":              "",
    "pdf_hash":              "",
    "pdf_text":              "",
    "report":                None,
    "raw_report":            "",
    "chat_history":          [],
    "model_used":            "",
    "session_start":         None,
    "upload_count":          0,
    "similarity_report":     None,
    "raw_similarity":        "",
    "search_results":        [],
    "agreement_log":         [],
    "feedback_given":        False,
    "feedback_submitted":    False,
    "_app_version":          APP_VERSION,
    # New feature states
    "ref_validation_report": None,
    "journal_fit_scores":    None,
    "generated_content":     None,
    "polished_section":      None,
    "cover_letter":          None,
    "citation_suggestions":  None,
    "checklist_report":      None,
}

if st.session_state.get("_app_version") != APP_VERSION:
    for k in list(st.session_state.keys()):
        del st.session_state[k]

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

if st.session_state.session_start is None:
    st.session_state.session_start = datetime.datetime.utcnow()

# ─── CUSTOM STYLING ───────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,600;1,400&family=Source+Serif+4:ital,opsz,wght@0,8..60,300;0,8..60,400;1,8..60,300&family=JetBrains+Mono:wght@400;500&display=swap');

:root {
    --navy: #0f2535; --navy2: #1a3a4a; --teal: #1d6b52; --teal2: #2d8f6f;
    --gold: #b5903a; --ivory: #f8f5ef; --ivory2: #f0ebe0; --stone: #e8e2d6;
    --text: #1a1612; --text2: #4a4540; --text3: #8a847a; --white: #ffffff;
    --success: #1d6b52; --warn: #8a5a00; --danger: #7a2020;
}

.stApp { background: var(--ivory) !important; font-family: 'Source Serif 4', Georgia, serif !important; }
#MainMenu, footer { visibility: hidden !important; }
.stDeployButton { display: none !important; }
header[data-testid="stHeader"] { background: transparent !important; height: 0 !important; min-height: 0 !important; }

.main .block-container { padding: 2rem 2.5rem 3rem !important; max-width: 1200px !important; }

[data-testid="stSidebar"] {
    background: #2c4a5a !important;
    border-right: 1px solid rgba(255,255,255,0.08) !important;
}
[data-testid="stSidebar"] p, [data-testid="stSidebar"] span:not([data-baseweb]),
[data-testid="stSidebar"] .stMarkdown p {
    color: #c8d8e0 !important; font-family: 'Source Serif 4', Georgia, serif !important;
    font-size: 0.84rem !important; line-height: 1.6 !important;
}
[data-testid="stSidebar"] h1 {
    font-family: 'Playfair Display', Georgia, serif !important;
    font-size: 1.35rem !important; font-weight: 600 !important;
    color: #eef4f8 !important; letter-spacing: 0.01em !important; margin-bottom: 0 !important;
}
[data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
    font-family: 'Source Serif 4', serif !important; font-size: 0.72rem !important;
    font-weight: 400 !important; color: #7a9aaa !important;
    letter-spacing: 0.14em !important; text-transform: uppercase !important;
    margin: 0.8rem 0 0.4rem !important;
}

h1, h2, h3 {
    font-family: 'Playfair Display', Georgia, serif !important;
    color: var(--navy) !important; letter-spacing: -0.01em !important;
}
h1 { font-size: 2rem !important; font-weight: 600 !important; }
h2 { font-size: 1.35rem !important; font-weight: 600 !important; }
h3 { font-size: 1.1rem !important; font-weight: 400 !important; font-style: italic !important; }

p, li, label { font-family: 'Source Serif 4', Georgia, serif !important; }

.stTabs [data-baseweb="tab-list"] { background: transparent !important; border-bottom: 1px solid var(--stone) !important; gap: 0 !important; }
.stTabs [data-baseweb="tab"] {
    font-family: 'Source Serif 4', serif !important; font-size: 0.78rem !important;
    font-weight: 400 !important; letter-spacing: 0.08em !important;
    text-transform: uppercase !important; color: var(--text3) !important;
    padding: 0.6rem 1.2rem !important; border-bottom: 2px solid transparent !important;
    background: transparent !important;
}
.stTabs [aria-selected="true"] {
    color: var(--navy) !important; border-bottom: 2px solid var(--gold) !important; font-weight: 400 !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 1.5rem !important; }

.stButton > button {
    font-family: 'Source Serif 4', serif !important; font-size: 0.85rem !important;
    letter-spacing: 0.06em !important; border-radius: 3px !important;
    border: 1px solid var(--stone) !important; background: var(--white) !important;
    color: var(--navy2) !important; padding: 0.5rem 1.2rem !important;
    transition: all 0.18s ease !important;
}
.stButton > button:hover {
    background: var(--ivory2) !important; border-color: var(--navy2) !important;
}

.stDownloadButton > button {
    font-family: 'Source Serif 4', serif !important; font-size: 0.82rem !important;
    letter-spacing: 0.05em !important; border-radius: 3px !important;
    background: var(--navy) !important; color: #e8e0d0 !important;
    border: none !important; padding: 0.5rem 1.2rem !important;
}
.stDownloadButton > button:hover { background: var(--navy2) !important; }

.stAlert {
    border-radius: 3px !important; font-family: 'Source Serif 4', serif !important;
    font-size: 0.87rem !important; border-left-width: 3px !important;
}
[data-testid="stInfoMessage"] { background: #f0f4f8 !important; border-left-color: var(--navy2) !important; }
[data-testid="stWarningMessage"] { background: #fdf6e3 !important; border-left-color: var(--gold) !important; }
[data-testid="stSuccessMessage"] { background: #f0f8f4 !important; border-left-color: var(--teal) !important; }
[data-testid="stErrorMessage"] { background: #fdf0f0 !important; border-left-color: var(--danger) !important; }

.stMarkdown p {
    font-family: 'Source Serif 4', Georgia, serif !important;
    font-size: 0.92rem !important; line-height: 1.75 !important; color: var(--text2) !important;
}
.stMarkdown blockquote {
    border-left: 3px solid var(--gold) !important; background: var(--ivory2) !important;
    padding: 0.8rem 1.2rem !important; border-radius: 0 3px 3px 0 !important;
    font-style: italic !important;
}
.stMarkdown code {
    font-family: 'JetBrains Mono', monospace !important; font-size: 0.78rem !important;
    background: var(--ivory2) !important; color: var(--navy2) !important;
    padding: 0.1em 0.4em !important; border-radius: 2px !important;
}

hr { border: none !important; border-top: 1px solid var(--stone) !important; margin: 1.2rem 0 !important; }

::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: var(--ivory2); }
::-webkit-scrollbar-thumb { background: var(--stone); border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: var(--text3); }
</style>
""", unsafe_allow_html=True)

# ─── GATE 1: ACCESS CODE ──────────────────────────────────────────────────────
if not st.session_state.access_granted:
    st.markdown(
        """
        <div style="max-width:480px;margin:6rem auto 2rem;text-align:center;">
          <div style="font-family:'Playfair Display',Georgia,serif;
                      font-size:0.7rem;letter-spacing:0.3em;text-transform:uppercase;
                      color:#8a9aaa;margin-bottom:1.5rem;">
            Health Professions Education
          </div>
          <h1 style="font-family:'Playfair Display',Georgia,serif;
                     font-size:2.6rem;font-weight:600;color:#0f2535;
                     margin:0 0 0.3rem;letter-spacing:-0.02em;">
            DentEdTech™
          </h1>
          <div style="font-family:'Source Serif 4',Georgia,serif;
                      font-size:1rem;font-weight:300;font-style:italic;
                      color:#6a7a8a;margin-bottom:2.5rem;">
            Pre-Submission Intelligence Platform
          </div>
          <div style="width:40px;height:1px;background:#b5903a;margin:0 auto 2.5rem;"></div>
          <p style="font-family:'Source Serif 4',Georgia,serif;
                    font-size:0.88rem;color:#4a4540;line-height:1.8;margin-bottom:2rem;">
            This platform is available to <strong>approved pilot partners</strong> only.<br>
            Please enter your institutional access code to continue.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        code_input = st.text_input(
            "Access code",
            placeholder="Enter your access code",
            type="password",
            label_visibility="collapsed",
        )
        if st.button("Enter Platform →", use_container_width=True):
            code = code_input.strip().upper()
            if code in VALID_ACCESS_CODES:
                st.session_state.access_granted = True
                st.session_state.access_partner = VALID_ACCESS_CODES[code]
                st.session_state.is_admin       = code in ADMIN_CODES
                st.rerun()
            else:
                st.error("Invalid access code.")
        st.markdown(
            "<div style='text-align:center;margin-top:1.5rem;font-family:Source Serif 4,serif;"
            "font-size:0.78rem;font-style:italic;color:#8a847a;'>"
            "Not a pilot partner? <a href='mailto:ayman.khalifah@manchester.ac.uk' "
            "style='color:#1a3a4a;'>Contact us</a> to discuss a partnership.</div>",
            unsafe_allow_html=True,
        )
    st.stop()

# ─── GATE 2: CONSENT ──────────────────────────────────────────────────────────
if not st.session_state.consent_given:
    st.markdown(
        f"""
        <div style="max-width:600px;margin:5rem auto 2rem;">
          <h1 style="font-family:'Playfair Display',Georgia,serif;font-size:2.2rem;
                     font-weight:600;color:#0f2535;text-align:center;">
            DentEdTech™
          </h1>
          <div style="font-family:'Source Serif 4',Georgia,serif;font-size:0.9rem;
                      font-style:italic;color:#6a7a8a;text-align:center;margin-bottom:2rem;">
            Pre-Submission Intelligence Platform
          </div>
          <div style="background:#f8f5ef;border:1px solid #e8e2d6;border-radius:4px;
                      padding:1.5rem 1.8rem;margin-bottom:1.5rem;">
            <div style="font-family:'Source Serif 4',serif;font-size:0.75rem;
                        letter-spacing:0.12em;text-transform:uppercase;color:#1d6b52;
                        margin-bottom:0.8rem;">
              ✓ Access granted &mdash; {st.session_state.access_partner}
            </div>
            <div style="font-family:'Playfair Display',serif;font-size:1rem;
                        font-weight:600;color:#0f2535;margin-bottom:1rem;">
              Data &amp; Confidentiality Notice
            </div>
            <ul style="font-family:'Source Serif 4',Georgia,serif;font-size:0.86rem;
                       line-height:1.9;color:#4a4540;padding-left:1.2rem;margin:0;">
              <li>Manuscripts are transmitted to <strong>Anthropic's API</strong> for analysis.
                  Anthropic does <strong>not</strong> train models on API data.</li>
              <li>Documents are held <strong>in memory only</strong> — never written to disk.</li>
              <li>Your session clears automatically when you close the browser tab.</li>
              <li>Do <strong>not</strong> upload manuscripts containing patient-identifiable
                  data or material under confidentiality agreement.</li>
            </ul>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        confirmed = st.checkbox(
            "I have read the data notice above and confirm the manuscript I will upload "
            "does not contain restricted or patient-identifiable data."
        )
        if st.button("Accept & Continue →", disabled=not confirmed, use_container_width=True):
            st.session_state.consent_given = True
            st.rerun()
    st.stop()

# ─── GATE 3: SESSION CAP ──────────────────────────────────────────────────────
def check_session_cap() -> bool:
    if st.session_state.is_admin:
        return True
    return st.session_state.analyses_this_session < SESSION_ANALYSIS_CAP

def show_cap_warning():
    if st.session_state.is_admin:
        return
    remaining = SESSION_ANALYSIS_CAP - st.session_state.analyses_this_session
    if remaining <= 1:
        st.warning(
            f"⚠️ You have **{remaining} analysis remaining** in this session. "
            f"Contact [DentEdTech™](mailto:ayman.khalifah@manchester.ac.uk) for extended access."
        )

# ─── HELPERS ──────────────────────────────────────────────────────────────────
def encode_pdf(uploaded_file) -> tuple[str, str, str]:
    raw = uploaded_file.read()
    b64 = base64.standard_b64encode(raw).decode("utf-8")
    sha = hashlib.sha256(raw).hexdigest()
    try:
        reader = PdfReader(BytesIO(raw))
        text = "\n".join(p.extract_text() or "" for p in reader.pages)
    except Exception:
        text = ""
    return b64, text, sha

def clear_session_data():
    sensitive_keys = [
        "pdf_base64", "pdf_name", "pdf_hash", "pdf_text",
        "report", "raw_report", "chat_history", "model_used",
        "similarity_report", "raw_similarity", "search_results",
        "feedback_given", "ref_validation_report", "journal_fit_scores",
        "generated_content", "polished_section", "cover_letter",
        "citation_suggestions", "checklist_report",
    ]
    for k in sensitive_keys:
        st.session_state[k] = defaults[k]
    st.session_state.upload_count = 0

def extract_title_abstract(text: str) -> tuple[str, str]:
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    title = lines[0] if lines else 'Untitled'
    abstract = ''
    for i, line in enumerate(lines):
        if re.search(r'(?i)^\s*abstract\s*$', line):
            abstract = ' '.join(lines[i+1:i+10])
            break
    if not abstract and len(lines) > 1:
        abstract = ' '.join(lines[1:])[:500]
    return title, abstract

def get_embedding(txt: str):
    return embedding_model.encode(txt)

def cosine_sim(a, b):
    return cosine_similarity([a], [b])[0][0]

def call_api_with_pdf(system: str, user_prompt: str, model: str, max_tok: int = 4096) -> str:
    response = client.messages.create(
        model=model,
        max_tokens=max_tok,
        system=system,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": st.session_state.pdf_base64,
                    },
                },
                {"type": "text", "text": user_prompt},
            ],
        }],
    )
    return response.content[0].text

def call_api_with_text(system: str, user_prompt: str, model: str, max_tok: int = 4096) -> str:
    text = st.session_state.pdf_text[:100_000]
    full_prompt = f"MANUSCRIPT TEXT:\n{text}\n\n{user_prompt}"
    response = client.messages.create(
        model=model,
        max_tokens=max_tok,
        system=system,
        messages=[{"role": "user", "content": full_prompt}],
    )
    return response.content[0].text

def parse_json(raw: str) -> dict | None:
    if not raw:
        return None
    cleaned = raw.strip()
    if "```" in cleaned:
        for part in cleaned.split("```"):
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            if part.startswith("{"):
                cleaned = part
                break
    try:
        start = cleaned.index("{")
    except ValueError:
        return None
    depth, end = 0, start
    for i, ch in enumerate(cleaned[start:], start):
        if ch == "{": depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                end = i + 1
                break
    candidate = cleaned[start:end]
    if depth != 0:
        open_count = candidate.count("{") - candidate.count("}")
        open_arr   = candidate.count("[") - candidate.count("]")
        candidate  = candidate.rstrip().rstrip(",").rstrip()
        candidate += "]" * max(0, open_arr) + "}" * max(0, open_count)
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        try:
            s = cleaned.index("{")
            e = cleaned.rindex("}") + 1
            return json.loads(cleaned[s:e])
        except Exception:
            return None

def render_badge(text: str, colour_key: str, palette: dict) -> str:
    bg, fg = palette.get(colour_key.lower().strip(), ("#e2e3e5", "#383d41"))
    return (
        f'<span style="background:{bg};color:{fg};padding:4px 14px;'
        f'border-radius:20px;font-weight:600;font-size:0.9rem;">{text}</span>'
    )

VERDICT_COLOURS = {
    "accept": ("#d4edda", "#155724"), "minor": ("#fff3cd", "#856404"),
    "major": ("#f8d7da", "#721c24"), "reject": ("#f8d7da", "#491217"),
}
CONFIDENCE_COLOURS = {
    "high": ("#d4edda", "#155724"), "moderate": ("#fff3cd", "#856404"),
    "low": ("#f8d7da", "#721c24"),
}
RISK_COLOURS = {
    "low": ("#d4edda", "#155724"), "moderate": ("#fff3cd", "#856404"),
    "high": ("#f8d7da", "#721c24"), "very high": ("#f5c6cb", "#491217"),
}

def verdict_key(v: str) -> str:
    vl = v.lower()
    if "reject" in vl: return "reject"
    if "major" in vl: return "major"
    if "minor" in vl: return "minor"
    return "accept"

def compute_agreement_stats() -> dict:
    log = st.session_state.agreement_log
    if not log:
        return {"total": 0, "agree": 0, "partial": 0, "disagree": 0, "rate": None}
    total = len(log)
    agree = sum(1 for r in log if r["feedback"] == "agree")
    partial = sum(1 for r in log if r["feedback"] == "partial")
    disagree = sum(1 for r in log if r["feedback"] == "disagree")
    rate = round((agree + 0.5 * partial) / total * 100, 1) if total else None
    return {"total": total, "agree": agree, "partial": partial, "disagree": disagree, "rate": rate}

# ─── REVIEW PROMPTS ───────────────────────────────────────────────────────────
def build_system_prompt(journal: str) -> str:
    return (
        f"You are a Senior Editor and double-blind Peer Reviewer for '{journal}', "
        "one of the most rigorous journals in Health Professions Education (HPE). "
        "Your reviews are precise, evidence-based, and constructive. "
        "You quote exact passages from the manuscript to substantiate every criticism. "
        "You never fabricate content. "
        "You apply CONSORT for RCTs, SRQR for qualitative research, COREQ for interviews/focus groups, "
        "and always evaluate educational outcomes through Kirkpatrick's four-level framework. "
        "You scrutinise the golden thread: the logical chain from research question through "
        "methodology, results, and conclusion. "
        "When you are uncertain about a verdict due to ambiguous methodology or borderline quality, "
        "you MUST flag this explicitly in the confidence field."
    )

def build_review_prompt(selected_criteria: list[str], journal: str) -> str:
    criteria_block = "\n".join(
        f"  {i+1}. {REVIEW_CRITERIA[c]}"
        for i, c in enumerate(selected_criteria)
    )
    return f"""Perform a comprehensive peer review of this manuscript submitted to '{journal}'.

SELECTED REVIEW CRITERIA:
{criteria_block}

Return ONLY a valid JSON object:

{{
  "verdict": "Accept | Minor Revisions | Major Revisions | Reject",
  "overall_score": <integer 1-100>,
  "confidence": "High | Moderate | Low",
  "confidence_note": "<what is uncertain>",
  "executive_summary": "<2-3 sentence assessment>",
  "scores": {{
    "novelty": <1-10>, "methodology": <1-10>, "clarity": <1-10>,
    "citations": <1-10>, "ethics": <1-10>
  }},
  "strengths": ["<strength>"],
  "weaknesses": [
    {{
      "section": "Abstract|Introduction|Methods|Results|Discussion|Citations",
      "issue": "<specific issue quoting manuscript>",
      "severity": "major|minor",
      "suggestion": "<concrete fix>"
    }}
  ],
  "section_comments": {{
    "abstract": "<comment>", "introduction": "<comment>",
    "methods": "<comment>", "results": "<comment>", "discussion": "<comment>"
  }},
  "golden_thread": "<coherence analysis>",
  "kirkpatrick_level": {{ "level": <1|2|3|4>, "justification": "<why>" }},
  "citation_audit": {{
    "missing_key_references": ["<Author Year — why relevant>"],
    "potentially_outdated": ["<citation — reason>"],
    "mismatches": "<issues or None identified>"
  }},
  "actionable_recommendations": ["<specific action>"],
  "author_feedback_summary": "<2-3 constructive sentences to authors>",
  "editor_note": "<confidential note to editor only>"
}}"""

# ─── DOCX GENERATORS ──────────────────────────────────────────────────────────
def create_docx(report: dict | None, raw: str) -> bytes:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    h = doc.add_heading("DentEdTech™ — Pre-Submission Review Report", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp = doc.sections[0].footer.paragraphs[0]
    fp.text = "CONFIDENTIAL — Generated by DentEdTech™. Powered by Anthropic API."
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    if report is None:
        doc.add_paragraph(raw)
    else:
        def h2(t): doc.add_heading(t, level=2)
        verdict = report.get("verdict", "—")
        score = report.get("overall_score", "—")
        confidence = report.get("confidence", "—")
        p = doc.add_paragraph()
        run = p.add_run(f"Verdict: {verdict}  |  Score: {score}/100  |  Confidence: {confidence}")
        run.bold = True; run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0x1A, 0x3A, 0x4A)
        doc.add_paragraph(report.get("executive_summary", ""))
        h2("Dimension Scores")
        for k, v in report.get("scores", {}).items():
            doc.add_paragraph(f"{k.capitalize()}: {v}/10", style="List Bullet")
        kp = report.get("kirkpatrick_level", {})
        if kp:
            h2("Kirkpatrick Level")
            doc.add_paragraph(f"Level {kp.get('level','?')}: {kp.get('justification','')}")
        h2("Golden Thread Analysis")
        doc.add_paragraph(report.get("golden_thread", ""))
        h2("Strengths")
        for s in report.get("strengths", []):
            doc.add_paragraph(s, style="List Bullet")
        h2("Weaknesses")
        for w in report.get("weaknesses", []):
            doc.add_paragraph(
                f"[{w.get('severity','').upper()} — {w.get('section','')}] "
                f"{w.get('issue','')}\n→ {w.get('suggestion','')}",
                style="List Bullet",
            )
        h2("Section-by-Section Comments")
        for sec, comment in report.get("section_comments", {}).items():
            p = doc.add_paragraph()
            p.add_run(sec.capitalize() + ": ").bold = True
            p.add_run(comment)
        h2("Citation Audit")
        ca = report.get("citation_audit", {})
        for ref in ca.get("missing_key_references", []):
            doc.add_paragraph(f"Missing: {ref}", style="List Bullet")
        for ref in ca.get("potentially_outdated", []):
            doc.add_paragraph(f"Outdated: {ref}", style="List Bullet")
        h2("Actionable Recommendations")
        for i, rec in enumerate(report.get("actionable_recommendations", []), 1):
            doc.add_paragraph(f"{i}. {rec}")

    buf = BytesIO(); doc.save(buf); return buf.getvalue()

def create_author_feedback_docx(report: dict) -> bytes:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    h = doc.add_heading("DentEdTech™ — Author Development Report", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp = doc.sections[0].footer.paragraphs[0]
    fp.text = "DentEdTech™ Author Feedback — For manuscript improvement only."
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def h2(t): doc.add_heading(t, level=2)
    note = doc.add_paragraph()
    note.add_run(
        "This report is designed to help you strengthen your manuscript before submission. "
        "All feedback is constructive and focused on enhancing quality."
    ).italic = True

    if report.get("author_feedback_summary"):
        h2("Overall Feedback")
        doc.add_paragraph(report["author_feedback_summary"])

    scores = report.get("scores", {})
    if scores:
        h2("Dimension Scores")
        for k, v in scores.items():
            doc.add_paragraph(f"{k.capitalize()}: {v}/10", style="List Bullet")

    h2("Logical Coherence (Golden Thread)")
    doc.add_paragraph(report.get("golden_thread", ""))

    if report.get("strengths"):
        h2("Strengths to Build On")
        for s in report["strengths"]:
            doc.add_paragraph(s, style="List Bullet")

    if report.get("weaknesses"):
        h2("Areas Requiring Attention")
        for w in report["weaknesses"]:
            p = doc.add_paragraph(style="List Bullet")
            p.add_run(f"[{w.get('severity','').upper()} — {w.get('section','')}] ").bold = True
            p.add_run(w.get("issue", ""))
            if w.get("suggestion"):
                doc.add_paragraph(f"Suggested: {w['suggestion']}", style="List Bullet")

    sc = report.get("section_comments", {})
    if sc:
        h2("Section-by-Section Feedback")
        for sec, comment in sc.items():
            p = doc.add_paragraph()
            p.add_run(sec.capitalize() + ": ").bold = True
            p.add_run(comment)

    ca = report.get("citation_audit", {})
    if ca:
        h2("Citation & Reference Feedback")
        for ref in ca.get("missing_key_references", []):
            doc.add_paragraph(f"Consider citing: {ref}", style="List Bullet")
        for ref in ca.get("potentially_outdated", []):
            doc.add_paragraph(f"Review currency: {ref}", style="List Bullet")

    if report.get("actionable_recommendations"):
        h2("Priority Actions")
        for i, rec in enumerate(report["actionable_recommendations"], 1):
            doc.add_paragraph(f"{i}. {rec}")

    buf = BytesIO(); doc.save(buf); return buf.getvalue()

# ─── NEW FEATURES ─────────────────────────────────────────────────────────────

# ─── 1. SEMANTIC SCHOLAR SIMILARITY SEARCH ────────────────────────────────────
def search_semantic_scholar(query: str, limit: int = 50) -> list[dict]:
    """Search Semantic Scholar for papers matching query."""
    url = "https://api.semanticscholar.org/graph/v1/paper/search"
    params = {
        "query": query,
        "limit": limit,
        "fields": "title,abstract,year,externalIds,url"
    }
    try:
        resp = requests.get(url, params=params, timeout=15)
        if resp.status_code == 200:
            return resp.json().get("data", [])
    except Exception:
        pass
    return []

def find_similar_papers(title: str, abstract: str, top_k: int = 10) -> list[tuple]:
    """Find semantically similar published papers."""
    candidates = search_semantic_scholar(title, limit=50)
    if not candidates:
        return []
    man_emb = get_embedding(title + " " + (abstract or ""))
    scored = []
    for paper in candidates:
        if not paper.get("abstract"):
            continue
        paper_txt = paper["title"] + " " + paper["abstract"]
        sim = cosine_sim(man_emb, get_embedding(paper_txt))
        scored.append((sim, paper))
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[:top_k]

# ─── 2. REFERENCE VALIDATION ──────────────────────────────────────────────────
CROSSREF_API = "https://api.crossref.org/works/"

def extract_dois(text: str) -> list[str]:
    doi_pattern = r'\b10\.\d{4,}/[^\s]+'
    return list(set(re.findall(doi_pattern, text)))

def validate_doi(doi: str) -> dict:
    try:
        r = requests.get(CROSSREF_API + doi, timeout=10)
        valid = r.status_code == 200
        return {"doi": doi, "valid": valid, "retracted": False, "error": None}
    except Exception as e:
        return {"doi": doi, "valid": False, "retracted": False, "error": str(e)}

def reference_validation_report(text: str) -> list[dict]:
    dois = extract_dois(text)
    return [validate_doi(d) for d in dois]

# ─── 3. JOURNAL FIT SCORING ───────────────────────────────────────────────────
def score_journal_fit(abstract: str) -> list[tuple]:
    if not abstract:
        return []
    man_emb = get_embedding(abstract)
    scores = []
    for name in JOURNALS:
        aim_text = JOURNAL_AIMS.get(name, name)
        j_emb = get_embedding(aim_text)
        scores.append((name, cosine_sim(man_emb, j_emb)))
    scores.sort(key=lambda x: x[1], reverse=True)
    return scores

# ─── 4. CONTENT GENERATOR ─────────────────────────────────────────────────────
def generate_content(user_prompt: str) -> str:
    system = (
        "You are a senior scientist and scientific writer in Health Professions Education. "
        "Expand the user's notes into publishable scientific text. "
        "Use real citations when possible; if not, give a placeholder and explain how to find it. "
        "After the text, list all suggested citations with justification. "
        "Be transparent about your choices."
    )
    resp = client.messages.create(
        model=FALLBACK_MODEL,
        max_tokens=3000,
        system=system,
        messages=[{"role": "user", "content": user_prompt}],
    )
    return resp.content[0].text

# ─── 5. LANGUAGE POLISHING ────────────────────────────────────────────────────
def polish_section(text: str, section_name: str = "pasted section") -> str:
    prompt = f"""Polish the following section from a scientific manuscript to be more 
clear, concise, and engaging for a top HPE journal. Preserve all scientific meaning.
After the rewritten text, list the changes you made and why.

Section ({section_name}):
{text}"""
    resp = client.messages.create(
        model=FALLBACK_MODEL,
        max_tokens=3000,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text

# ─── 6. COVER LETTER GENERATOR ────────────────────────────────────────────────
def generate_cover_letter(title: str, abstract: str, highlights: str, journal: str) -> str:
    prompt = f"""Draft a cover letter for the editor of {journal}.
Manuscript title: {title}
Abstract: {abstract}
Key highlights provided by author: {highlights}
The letter should be formal, persuasive, and include a statement of novelty.
After the letter, explain the reasoning behind the phrasing."""
    resp = client.messages.create(
        model=FALLBACK_MODEL,
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text

# ─── 7. CITATION SUGGESTER ───────────────────────────────────────────────────
def suggest_citations(claim: str, field: str = "Health Professions Education") -> str:
    prompt = f"""Given the claim: "{claim}" in the field of {field},
suggest 2-3 highly relevant, real academic citations (full reference), 
and explain why each supports the claim. If you cannot recall a real paper, 
state clearly you are using a placeholder and suggest search terms."""
    resp = client.messages.create(
        model=FALLBACK_MODEL,
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text

# ─── 8. REPORTING CHECKLIST ───────────────────────────────────────────────────
def check_guideline(manuscript: str, guideline_name: str) -> str:
    prompt = f"""Check if the following manuscript adheres to the {guideline_name} reporting guideline.
Output a checklist with each item marked as Yes/No/Partial, and explain the reasoning for each.
Manuscript (first portion):
{manuscript[:8000]}"""
    resp = client.messages.create(
        model=FALLBACK_MODEL,
        max_tokens=2500,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.content[0].text

# ─── 9. SIMILARITY AUDIT (existing from DentEdTech) ───────────────────────────
def extract_search_queries(pdf_text: str) -> list[str]:
    prompt = (
        "Read this manuscript excerpt and extract exactly 5 short search queries "
        "(4-8 words each) representing the most distinctive claims, methods, or findings. "
        "Return ONLY a JSON array of 5 strings.\n\n"
        f"MANUSCRIPT EXCERPT:\n{pdf_text[:8000]}"
    )
    try:
        r = client.messages.create(
            model=CHAT_MODEL, max_tokens=300,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = r.content[0].text.strip()
        return json.loads(raw[raw.index("["):raw.rindex("]")+1])
    except Exception:
        lines = [l.strip() for l in pdf_text.split("\n") if len(l.strip()) > 40]
        return [lines[0][:80]] if lines else ["health professions education quality assurance"]

def search_web(queries: list[str]) -> list[dict]:
    if not DDG_AVAILABLE:
        return []
    results, seen = [], set()
    try:
        with DDGS() as ddgs:
            for q in queries:
                try:
                    for r in ddgs.text(
                        f"{q} site:pubmed.ncbi.nlm.nih.gov OR site:researchgate.net "
                        f"OR site:tandfonline.com OR site:wiley.com OR site:springer.com "
                        f"OR site:sciencedirect.com",
                        max_results=3,
                    ):
                        url = r.get("href", "")
                        if url not in seen:
                            seen.add(url)
                            results.append({
                                "title": r.get("title", ""),
                                "url": url,
                                "body": r.get("body", ""),
                                "query": q,
                            })
                except Exception:
                    continue
    except Exception:
        pass
    return results[:12]

def build_similarity_prompt(search_results: list[dict]) -> str:
    search_block = ""
    if search_results:
        search_block = "\n\nSIMILAR PUBLISHED PAPERS FOUND ONLINE:\n"
        for i, r in enumerate(search_results, 1):
            search_block += (
                f"\n[{i}] Title: {r['title']}\n"
                f"    URL: {r['url']}\n"
                f"    Summary: {r['body'][:300]}\n"
            )
    return f"""Analyse this manuscript for originality and similarity risks.
1. Assess overlap with published literature.
2. Identify boilerplate passages.
3. Flag citation-free claims.
4. Check internal repetition.
5. Assess methods originality.
6. Provide overall risk and rewrite advice.
{search_block}

Return ONLY valid JSON:
{{
  "overall_risk_level": "Low | Moderate | High | Very High",
  "estimated_similarity_risk_percent": <integer 0-100>,
  "disclaimer": "AI-based. Not equivalent to Turnitin. Use institutional checker.",
  "similar_publications": [
    {{"title": "<title>", "url": "<url>", "overlap_description": "<what>",
      "overlap_type": "topic|methodology|findings|framing|significant overlap",
      "risk_level": "Low|Moderate|High", "recommendation": "<what to do>"}}
  ],
  "boilerplate_sections": [
    {{"section": "<where>", "passage": "<text>", "risk": "<why>", "suggestion": "<fix>"}}
  ],
  "citation_free_claims": [
    {{"passage": "<text>", "risk": "<what>", "suggestion": "<cite>"}}
  ],
  "internal_repetition": [
    {{"passage": "<text>", "appears_in": ["<sec1>", "<sec2>"]}}
  ],
  "methods_risk": "<assessment>",
  "priority_rewrites": ["<rewrite instruction>"],
  "submission_readiness": "<verdict>"
}}"""

# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        """
        <div style="padding:0.5rem 0 1rem;">
          <div style="font-family:'Playfair Display',Georgia,serif;
                      font-size:1.4rem;font-weight:600;color:#f0ebe0;
                      letter-spacing:-0.01em;line-height:1.2;">
            DentEdTech™
          </div>
          <div style="font-family:'Source Serif 4',Georgia,serif;
                      font-size:0.72rem;font-weight:300;font-style:italic;
                      color:#7a8a9a;margin-top:0.2rem;letter-spacing:0.04em;">
            Pre-Submission Intelligence
          </div>
          <div style="width:24px;height:1px;background:#b5903a;margin-top:0.8rem;"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption(f"Partner: {st.session_state.access_partner}")

    if st.session_state.is_admin:
        st.markdown(
            "<div style='font-family:Source Serif 4,serif;font-size:0.72rem;"
            "color:#8fd4b8;'>✧ Admin — unlimited analyses</div>",
            unsafe_allow_html=True,
        )
    else:
        used = st.session_state.analyses_this_session
        remaining = SESSION_ANALYSIS_CAP - used
        cap_pct = int(used / SESSION_ANALYSIS_CAP * 100)
        st.progress(cap_pct / 100, text=f"Session: {used}/{SESSION_ANALYSIS_CAP} analyses")
        if remaining == 0:
            st.error("Session limit reached.")

    with st.expander("🔒 Data & Privacy", expanded=False):
        st.markdown(
            f"| Item | Status |\n|---|---|\n"
            f"| Partner | {st.session_state.access_partner} |\n"
            f"| Files on disk | ✅ Never |\n"
            f"| API provider | Anthropic |\n"
            f"| Training on data | ✅ No |\n"
            f"| Documents | {st.session_state.upload_count} |"
        )

    stats = compute_agreement_stats()
    if stats["total"] > 0:
        with st.expander(f"📊 Agreement ({stats['total']} reviews)", expanded=False):
            if stats["rate"] is not None:
                st.metric("Rate", f"{stats['rate']}%")
            c1, c2, c3 = st.columns(3)
            c1.metric("✅", stats["agree"])
            c2.metric("⚠️", stats["partial"])
            c3.metric("❌", stats["disagree"])

    st.divider()

    uploaded = st.file_uploader("Upload manuscript (PDF)", type=["pdf"])
    if uploaded:
        if uploaded.name != st.session_state.pdf_name or not st.session_state.pdf_base64:
            with st.spinner("Encoding PDF…"):
                b64, txt, sha = encode_pdf(uploaded)
            if sha != st.session_state.pdf_hash:
                st.session_state.pdf_base64 = b64
                st.session_state.pdf_text = txt
                st.session_state.pdf_hash = sha
                st.session_state.pdf_name = uploaded.name
                st.session_state.report = None
                st.session_state.raw_report = ""
                st.session_state.chat_history = []
                st.session_state.similarity_report = None
                st.session_state.raw_similarity = ""
                st.session_state.search_results = []
                st.session_state.feedback_given = False
                st.session_state.ref_validation_report = None
                st.session_state.journal_fit_scores = None
                st.session_state.upload_count += 1
                st.success(f"✅ {uploaded.name}")
        st.caption(f"{len(st.session_state.pdf_text):,} chars")

    st.divider()
    journal = st.selectbox("Target journal", JOURNALS)
    st.markdown("**Review criteria**")
    selected_criteria = [
        k for k, label in REVIEW_CRITERIA.items()
        if st.checkbox(label, value=True, key=f"cb_{k}")
    ]

    st.divider()
    can_analyze = (
        bool(st.session_state.pdf_base64)
        and len(selected_criteria) > 0
        and check_session_cap()
    )
    if st.button("🚀 Run Full Analysis", disabled=not can_analyze, use_container_width=True):
        st.session_state.report = None
        st.session_state.raw_report = ""
        st.session_state.chat_history = []
        st.session_state.feedback_given = False
        st.session_state["_trigger_analysis"] = True

    st.divider()
    if st.button("🗑️ Clear Session", use_container_width=True):
        clear_session_data()
        st.success("Session cleared.")
        st.rerun()
    st.caption("⚠️ All data clears when you close this tab.")

# ─── ANALYSIS TRIGGER ─────────────────────────────────────────────────────────
if st.session_state.get("_trigger_analysis"):
    st.session_state["_trigger_analysis"] = False
    if not check_session_cap():
        st.error("Session limit reached.")
        st.stop()
    system = build_system_prompt(journal)
    prompt = build_review_prompt(selected_criteria, journal)
    phases = [
        "Phase 1 — Document structure audit",
        "Phase 2 — Methodology assessment",
        "Phase 3 — Citation audit",
        "Phase 4 — Generating review",
    ]
    progress = st.progress(0)
    status = st.status("Running analysis…", expanded=True)
    raw = None; model_used = PRIMARY_MODEL
    for i, phase in enumerate(phases):
        status.write(f"⚙️ {phase}")
        progress.progress((i + 1) / len(phases))
    try:
        status.write(f"🧠 Sending to {PRIMARY_MODEL}…")
        raw = call_api_with_pdf(system, prompt, PRIMARY_MODEL)
        model_used = PRIMARY_MODEL
    except NotFoundError:
        status.write(f"⚠️ Falling back to {FALLBACK_MODEL}…")
        try:
            raw = call_api_with_pdf(system, prompt, FALLBACK_MODEL)
            model_used = FALLBACK_MODEL
        except Exception:
            raw = call_api_with_text(system, prompt, FALLBACK_MODEL)
            model_used = FALLBACK_MODEL + " (text)"
    except Exception:
        try:
            raw = call_api_with_text(system, prompt, PRIMARY_MODEL)
            model_used = PRIMARY_MODEL + " (text)"
        except Exception as e2:
            status.update(label=f"Error: {e2}", state="error")
            st.stop()
    progress.progress(1.0)
    st.session_state.report = parse_json(raw)
    st.session_state.raw_report = raw
    st.session_state.model_used = model_used
    st.session_state.analyses_this_session += 1

    parsed_seed = parse_json(raw)
    if parsed_seed:
        verdict = parsed_seed.get("verdict", "—")
        score = parsed_seed.get("overall_score", "—")
        confidence = parsed_seed.get("confidence", "—")
        summary = parsed_seed.get("executive_summary", "")
        strengths = parsed_seed.get("strengths", [])
        weaknesses = parsed_seed.get("weaknesses", [])
        recs = parsed_seed.get("actionable_recommendations", [])
        seed_msg = (
            f"I have completed a full structured peer review of this manuscript.\n\n"
            f"VERDICT: {verdict} | Score: {score}/100 | Confidence: {confidence}\n\n"
            f"SUMMARY: {summary}\n\n"
            f"STRENGTHS ({len(strengths)}):\n" +
            "\n".join(f"- {s}" for s in strengths[:5]) + "\n\n"
            f"KEY WEAKNESSES ({len(weaknesses)}):\n" +
            "\n".join(f"- [{w.get('severity','').upper()}] {w.get('section','')}: {w.get('issue','')}" for w in weaknesses[:5]) + "\n\n"
            f"TOP RECOMMENDATIONS:\n" +
            "\n".join(f"{i+1}. {r}" for i, r in enumerate(recs[:5])) + "\n\n"
            f"I am ready to answer any questions about this review."
        )
    else:
        seed_msg = "I have completed the review. Ask me anything about the manuscript."

    st.session_state.chat_history = [
        {"role": "user", "content": [
            {"type": "document", "source": {"type": "base64", "media_type": "application/pdf",
                                            "data": st.session_state.pdf_base64}},
            {"type": "text", "text": "This is the manuscript. Answer in plain English — never return JSON."},
        ]},
        {"role": "assistant", "content": seed_msg},
    ]
    status.update(label="Analysis complete ✓", state="complete", expanded=False)
    st.rerun()

# ─── IDLE STATE ───────────────────────────────────────────────────────────────
if not st.session_state.report and not st.session_state.raw_report:
    st.markdown(
        """
        <div style="max-width:680px;margin:5rem auto;text-align:center;">
          <h1 style="font-family:'Playfair Display',Georgia,serif;
                     font-size:3rem;font-weight:600;color:#0f2535;
                     margin:0 0 0.4rem;letter-spacing:-0.03em;">
            DentEdTech™
          </h1>
          <div style="font-family:'Source Serif 4',Georgia,serif;
                      font-size:1.05rem;font-weight:300;font-style:italic;
                      color:#6a7a8a;margin-bottom:2rem;">
            Pre-Submission Intelligence for HPE Journals
          </div>
          <div style="width:48px;height:1px;background:#b5903a;margin:0 auto 2.5rem;"></div>
          <p style="font-family:'Source Serif 4',Georgia,serif;
                    font-size:0.92rem;color:#6a6460;line-height:1.85;">
            Upload a manuscript, select your target journal and criteria, then run the analysis.
            Access all pre-submission tools from the tabs above.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

# ─── MAIN DISPLAY ─────────────────────────────────────────────────────────────
report = st.session_state.report
raw = st.session_state.raw_report
show_cap_warning()
st.markdown(
    f"<div style='font-family:Source Serif 4,serif;font-size:0.75rem;"
    f"font-style:italic;color:#8a847a;margin-bottom:0.5rem;'>"
    f"Analysis: {st.session_state.model_used}</div>",
    unsafe_allow_html=True,
)

# Main tabs including all pre-submission features
tab_names = [
    "📋 Editor Report",
    "👤 Author Feedback",
    "🔍 Similarity Audit",
    "💬 Editor Chat",
    "📚 Reference Validation",
    "🎯 Journal Fit",
    "✨ Content Generator",
    "✍️ Polish Language",
    "📝 Cover Letter",
    "🔖 Citation Suggestions",
    "✔️ Reporting Checklist",
    "📊 Feedback",
]
tabs = st.tabs(tab_names)

# ─── TAB 0: EDITOR REPORT ─────────────────────────────────────────────────────
with tabs[0]:
    if report is None:
        st.warning("Could not parse structured JSON — showing raw report.")
        st.text_area("Raw report", raw, height=600)
    else:
        verdict = report.get("verdict", "Unknown")
        score = report.get("overall_score", "—")
        confidence = report.get("confidence", "Unknown")
        conf_note = report.get("confidence_note", "")

        col_v, col_c, col_s, col_dl = st.columns([3, 2, 1, 1])
        with col_v:
            st.markdown(render_badge(verdict, verdict_key(verdict), VERDICT_COLOURS), unsafe_allow_html=True)
        with col_c:
            st.markdown(render_badge(f"Confidence: {confidence}", confidence.lower(), CONFIDENCE_COLOURS), unsafe_allow_html=True)
        with col_s:
            st.metric("Score", f"{score}/100")
        with col_dl:
            st.download_button(
                "⬇️ Full Report",
                data=create_docx(report, raw),
                file_name="DentEdTech_Editor_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        if confidence.lower() in ("moderate", "low"):
            st.warning(f"⚠️ **Confidence flag:** {conf_note}")

        st.markdown(f"> {report.get('executive_summary','')}")
        st.divider()

        scores = report.get("scores", {})
        if scores:
            cols = st.columns(len(scores))
            for col, (k, v) in zip(cols, scores.items()):
                col.metric(k.capitalize(), f"{v}/10")

        kp = report.get("kirkpatrick_level", {})
        if kp:
            st.info(f"🎯 **Kirkpatrick Level {kp.get('level','?')}** — {kp.get('justification','')}")

        with st.expander("🧵 Golden Thread Analysis", expanded=True):
            st.write(report.get("golden_thread", "—"))

        if report.get("strengths"):
            with st.expander(f"✅ Strengths ({len(report['strengths'])})", expanded=True):
                for s in report["strengths"]:
                    st.markdown(f"- {s}")

        weaknesses = report.get("weaknesses", [])
        if weaknesses:
            with st.expander(f"⚠️ Weaknesses ({len(weaknesses)})", expanded=True):
                for w in weaknesses:
                    sev = w.get("severity", "minor")
                    st.markdown(
                        f"{'🔴' if sev=='major' else '🟡'} "
                        f"**[{sev.upper()} — {w.get('section','')}]** {w.get('issue','')}"
                    )
                    if w.get("suggestion"):
                        st.caption(f"→ {w['suggestion']}")

        sc = report.get("section_comments", {})
        if sc:
            with st.expander("📝 Section-by-Section Comments"):
                for section, comment in sc.items():
                    st.markdown(f"**{section.capitalize()}**")
                    st.write(comment)
                    st.divider()

        ca = report.get("citation_audit", {})
        if ca:
            with st.expander("📚 Citation Audit"):
                for ref in ca.get("missing_key_references", []):
                    st.markdown(f"- Missing: {ref}")
                for ref in ca.get("potentially_outdated", []):
                    st.markdown(f"- Outdated: {ref}")
                st.markdown(f"**Mismatches:** {ca.get('mismatches','None identified')}")

        recs = report.get("actionable_recommendations", [])
        if recs:
            with st.expander(f"✅ Actionable Recommendations ({len(recs)})", expanded=True):
                for i, rec in enumerate(recs, 1):
                    st.markdown(f"**{i}.** {rec}")

        if report.get("editor_note"):
            with st.expander("🔒 Confidential Note to Editor"):
                st.info(report["editor_note"])

        # Agreement feedback
        st.divider()
        st.markdown("#### 📊 Editor Agreement Feedback")
        if not st.session_state.feedback_given:
            col_a, col_b, col_c = st.columns(3)
            if col_a.button("✅ I agree", use_container_width=True):
                st.session_state.agreement_log.append({
                    "verdict": verdict, "score": score,
                    "feedback": "agree",
                    "timestamp": datetime.datetime.utcnow().isoformat(),
                })
                st.session_state.feedback_given = True
                st.rerun()
            if col_b.button("⚠️ Partially", use_container_width=True):
                st.session_state.agreement_log.append({
                    "verdict": verdict, "score": score,
                    "feedback": "partial",
                    "timestamp": datetime.datetime.utcnow().isoformat(),
                })
                st.session_state.feedback_given = True
                st.rerun()
            if col_c.button("❌ Disagree", use_container_width=True):
                st.session_state.agreement_log.append({
                    "verdict": verdict, "score": score,
                    "feedback": "disagree",
                    "timestamp": datetime.datetime.utcnow().isoformat(),
                })
                st.session_state.feedback_given = True
                st.rerun()
        else:
            st.success("✅ Feedback recorded — thank you.")

# ─── TAB 1: AUTHOR FEEDBACK ───────────────────────────────────────────────────
with tabs[1]:
    st.info("Constructive author-facing report — no verdict. Safe to share with authors.")
    if report is None:
        st.warning("Run a full analysis first.")
    else:
        with st.expander("👁️ Preview", expanded=True):
            st.markdown("**Overall feedback:**")
            st.write(report.get("author_feedback_summary", "No summary."))
            st.markdown("**Strengths:**")
            for s in report.get("strengths", []):
                st.markdown(f"- {s}")
            st.markdown("**Priority actions:**")
            for i, rec in enumerate(report.get("actionable_recommendations", []), 1):
                st.markdown(f"**{i}.** {rec}")
        st.download_button(
            "⬇️ Download Author Report (.docx)",
            data=create_author_feedback_docx(report),
            file_name="DentEdTech_Author_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

# ─── TAB 2: SIMILARITY AUDIT ──────────────────────────────────────────────────
with tabs[2]:
    st.info("⚠️ **Disclaimer:** Not equivalent to Turnitin. Use institutional checker before submission.")

    # Semantic Scholar Similarity
    st.subheader("Semantic Scholar Similarity Search")
    if st.button("🔬 Run Semantic Scholar Similarity", disabled=not bool(st.session_state.pdf_base64)):
        with st.spinner("Searching Semantic Scholar..."):
            title, abstract = extract_title_abstract(st.session_state.pdf_text)
            similar = find_similar_papers(title, abstract, top_k=10)
            st.session_state.semantic_similar = similar
            st.rerun()

    if "semantic_similar" in st.session_state and st.session_state.semantic_similar:
        for i, (score, paper) in enumerate(st.session_state.semantic_similar):
            st.markdown(f"**{i+1}. {paper.get('title','')}** — similarity: `{score:.3f}`")
            if paper.get("abstract"):
                st.caption(paper["abstract"][:300] + "...")
            if paper.get("url"):
                st.markdown(f"🔗 [{paper['url']}]({paper['url']})")
            st.divider()

    st.divider()

    # Existing DuckDuckGo similarity audit
    if not DDG_AVAILABLE:
        st.warning("Web search disabled. Install `duckduckgo-search` for full similarity audit.")

    if st.button("🕸️ Run Web Similarity Audit", disabled=not bool(st.session_state.pdf_base64)):
        st.session_state.similarity_report = None
        st.session_state.raw_similarity = ""
        st.session_state.search_results = []

        with st.status("Running similarity audit…", expanded=True) as ss:
            ss.write("🔎 Extracting key phrases…")
            queries = extract_search_queries(st.session_state.pdf_text)
            ss.write(f"   {len(queries)} queries")

            if DDG_AVAILABLE:
                ss.write("🌐 Searching web…")
                sr = search_web(queries)
                st.session_state.search_results = sr
                ss.write(f"   {len(sr)} papers found")
            else:
                sr = []

            ss.write("🧠 AI originality analysis…")
            sim_prompt = build_similarity_prompt(sr)
            sim_system = "You are an academic integrity specialist. Analyse for similarity risks."
            sim_raw = None
            try:
                sim_raw = call_api_with_pdf(sim_system, sim_prompt, FALLBACK_MODEL, max_tok=8000)
            except Exception:
                try:
                    sim_raw = call_api_with_text(sim_system, sim_prompt, FALLBACK_MODEL, max_tok=8000)
                except Exception as e2:
                    ss.update(label=f"Error: {e2}", state="error"); st.stop()

            parsed_sim = parse_json(sim_raw)
            st.session_state.similarity_report = parsed_sim
            st.session_state.raw_similarity = sim_raw
            ss.update(label="Complete ✓", state="complete")
        st.rerun()

    sim = st.session_state.similarity_report
    if sim:
        risk = sim.get("overall_risk_level", "Unknown")
        est = sim.get("estimated_similarity_risk_percent", "—")
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(render_badge(f"Risk: {risk}", risk.lower(), RISK_COLOURS), unsafe_allow_html=True)
        with col2:
            st.metric("Estimated %", f"~{est}%")
        st.caption(sim.get("disclaimer", ""))
        st.markdown(f"**Readiness:** {sim.get('submission_readiness','—')}")
        st.divider()

        pubs = sim.get("similar_publications", [])
        with st.expander(f"🌐 Similar Publications ({len(pubs)})", expanded=bool(pubs)):
            if pubs:
                for p in pubs:
                    rl = p.get("risk_level", "")
                    st.markdown(f"**{p.get('title','')}** {render_badge(rl, rl.lower(), RISK_COLOURS)}", unsafe_allow_html=True)
                    if p.get("url"): st.markdown(f"🔗 [{p['url']}]({p['url']})")
                    st.markdown(f"**Overlap:** {p.get('overlap_description','')}")
                    if p.get("recommendation"): st.success(f"✏️ {p['recommendation']}")
                    st.divider()
            else:
                st.success("✅ No significantly similar papers identified.")

        boiler = sim.get("boilerplate_sections", [])
        with st.expander(f"📋 Boilerplate ({len(boiler)})", expanded=bool(boiler)):
            if boiler:
                for item in boiler:
                    st.warning(f"🔴 {item.get('passage','')}")
                    st.success(f"✏️ {item.get('suggestion','')}")
                    st.divider()
            else:
                st.success("✅ No significant boilerplate.")

        claims = sim.get("citation_free_claims", [])
        with st.expander(f"📎 Citation-Free Claims ({len(claims)})", expanded=bool(claims)):
            if claims:
                for item in claims:
                    st.warning(f"🟡 {item.get('passage','')}")
                    st.success(f"✏️ {item.get('suggestion','')}")
                    st.divider()
            else:
                st.success("✅ All claims appear cited.")

        rewrites = sim.get("priority_rewrites", [])
        if rewrites:
            with st.expander(f"✏️ Priority Rewrites ({len(rewrites)})", expanded=True):
                for i, rw in enumerate(rewrites, 1):
                    st.markdown(f"**{i}.** {rw}")

# ─── TAB 3: EDITOR CHAT ───────────────────────────────────────────────────────
with tabs[3]:
    st.caption("Ask questions about the review or manuscript.")

    quick_prompts = [
        "Expand on the methodology critique",
        "Which citations are missing and why?",
        "How can the Discussion be strengthened?",
        "Explain the golden thread score",
        "What would reach Kirkpatrick Level 3 or 4?",
        "Suggest a revised abstract",
    ]
    cols = st.columns(3)
    for i, qp in enumerate(quick_prompts):
        if cols[i % 3].button(qp, key=f"qp_{i}", use_container_width=True):
            st.session_state._pending_chat = qp

    st.divider()

    for msg in st.session_state.chat_history[2:]:
        role = msg["role"]
        content = msg["content"] if isinstance(msg["content"], str) else str(msg["content"])
        with st.chat_message(role):
            st.markdown(content)

    pending = st.session_state.pop("_pending_chat", None)
    user_input = st.chat_input("Ask about the review…") or pending

    if user_input:
        with st.chat_message("user"):
            st.markdown(user_input)
        st.session_state.chat_history.append({"role": "user", "content": user_input})
        with st.chat_message("assistant"):
            with st.spinner("Thinking…"):
                try:
                    response = client.messages.create(
                        model=FALLBACK_MODEL,
                        max_tokens=2048,
                        system=(
                            "You are a Senior HPE Journal Editor discussing a peer review. "
                            "Answer in clear, natural English prose. Never return JSON. "
                            "Quote specific manuscript passages. Be constructive and precise."
                        ),
                        messages=st.session_state.chat_history,
                    )
                    reply = response.content[0].text
                except Exception as e:
                    reply = f"Error: {e}"
            st.markdown(reply)
        st.session_state.chat_history.append({"role": "assistant", "content": reply})

# ─── TAB 4: REFERENCE VALIDATION ──────────────────────────────────────────────
with tabs[4]:
    st.subheader("DOI Validation & Reference Check")
    st.caption("Validates all DOIs found in the manuscript against CrossRef.")

    if "manuscript" not in st.session_state and not st.session_state.pdf_text:
        st.warning("Upload a manuscript first.")
    else:
        if st.button("🔍 Validate References", disabled=not bool(st.session_state.pdf_text)):
            with st.spinner("Validating DOIs..."):
                ref_report = reference_validation_report(st.session_state.pdf_text)
                st.session_state.ref_validation_report = ref_report
                st.rerun()

        if st.session_state.ref_validation_report:
            valid_count = sum(1 for r in st.session_state.ref_validation_report if r["valid"])
            total_count = len(st.session_state.ref_validation_report)
            st.metric("Validation Results", f"{valid_count}/{total_count} valid")

            for r in st.session_state.ref_validation_report:
                status_icon = "✅" if r["valid"] else "❌"
                retracted = "⚠️ RETRACTED" if r.get("retracted") else ""
                st.markdown(f"{status_icon} `{r['doi']}` {retracted}")
                if r.get("error"):
                    st.caption(f"Error: {r['error']}")
            st.divider()

        # AI reference analysis
        if st.button("🤖 AI Reference Analysis", disabled=not bool(st.session_state.pdf_text)):
            with st.spinner("Analysing references with AI..."):
                prompt = (
                    "Analyse the reference list of this manuscript. Check for:\n"
                    "1. Currency — are recent (last 5 years) references included?\n"
                    "2. Completeness — any missing key papers in the field?\n"
                    "3. Format consistency — any mismatches between in-text and reference list?\n"
                    "4. Suggestions for additional references."
                )
                try:
                    ai_ref_analysis = call_api_with_text(
                        "You are a citation and reference specialist for HPE journals.",
                        prompt, FALLBACK_MODEL, max_tok=2000
                    )
                    st.markdown(ai_ref_analysis)
                except Exception as e:
                    st.error(f"Error: {e}")

# ─── TAB 5: JOURNAL FIT ───────────────────────────────────────────────────────
with tabs[5]:
    st.subheader("Journal Fit Scoring")
    st.caption("Semantic similarity between your manuscript abstract and journal aims & scope.")

    if not st.session_state.pdf_text:
        st.warning("Upload a manuscript first.")
    else:
        if st.button("🎯 Score Journal Fit"):
            with st.spinner("Computing journal fit scores..."):
                _, abstract = extract_title_abstract(st.session_state.pdf_text)
                scores = score_journal_fit(abstract)
                st.session_state.journal_fit_scores = scores[:10]  # Top 10
                st.rerun()

        if st.session_state.journal_fit_scores:
            st.markdown("### Top Journal Matches")
            for i, (name, score) in enumerate(st.session_state.journal_fit_scores):
                # Color-code based on score
                if score > 0.7:
                    color = "#1d6b52"
                elif score > 0.5:
                    color = "#b5903a"
                else:
                    color = "#7a2020"
                st.markdown(
                    f"**{i+1}.** {name} — "
                    f"<span style='color:{color};font-weight:600;'>{score:.3f}</span>",
                    unsafe_allow_html=True,
                )
            st.caption("Higher score = better semantic match with journal aims & scope.")

# ─── TAB 6: CONTENT GENERATOR ─────────────────────────────────────────────────
with tabs[6]:
    st.subheader("Idea → Polished Research Content")
    st.caption("Provide rough notes or bullet points. The AI expands them into publishable prose with citation suggestions.")

    user_idea = st.text_area(
        "Your rough notes / ideas",
        height=150,
        placeholder="e.g., 'Need a strong introduction about competency-based assessment in dental education. Mention recent systematic reviews and the shift from procedural to holistic assessment.'"
    )

    if st.button("✨ Generate Content"):
        if not user_idea.strip():
            st.warning("Please enter some hints.")
        else:
            with st.spinner("Writing polished content..."):
                result = generate_content(user_idea)
                st.session_state.generated_content = result
                st.rerun()

    if st.session_state.generated_content:
        st.markdown(st.session_state.generated_content)
        st.download_button(
            "⬇️ Download as Text",
            st.session_state.generated_content,
            "generated_content.txt",
        )

# ─── TAB 7: LANGUAGE POLISH ───────────────────────────────────────────────────
with tabs[7]:
    st.subheader("Polish Manuscript Section")
    st.caption("Paste any section from your manuscript to improve clarity, style, and impact.")

    section_text = st.text_area(
        "Paste a section to polish",
        height=200,
        placeholder="Paste your Abstract, Introduction paragraph, or any section here..."
    )

    if st.button("✍️ Polish Section"):
        if not section_text.strip():
            st.warning("Please paste some text.")
        else:
            with st.spinner("Improving style and clarity..."):
                polished = polish_section(section_text)
                st.session_state.polished_section = polished
                st.rerun()

    if st.session_state.polished_section:
        st.markdown(st.session_state.polished_section)
        st.download_button(
            "⬇️ Download Polished Text",
            st.session_state.polished_section,
            "polished_section.txt",
        )

# ─── TAB 8: COVER LETTER ──────────────────────────────────────────────────────
with tabs[8]:
    st.subheader("Auto Cover Letter Generator")
    st.caption("Generate a persuasive cover letter for your target journal.")

    highlights = st.text_area(
        "Key highlights (comma separated)",
        placeholder="e.g., First systematic review of X, Novel competency framework, Multi-site international study"
    )
    target_journal_cover = st.selectbox("Target journal (cover letter)", JOURNALS, key="cover_journal")

    if st.button("📝 Generate Cover Letter"):
        if not st.session_state.pdf_text:
            st.warning("Upload a manuscript for context.")
        elif not highlights.strip():
            st.warning("Please provide key highlights.")
        else:
            with st.spinner("Drafting cover letter..."):
                title, abstract = extract_title_abstract(st.session_state.pdf_text)
                letter = generate_cover_letter(title, abstract, highlights, target_journal_cover)
                st.session_state.cover_letter = letter
                st.rerun()

    if st.session_state.cover_letter:
        st.markdown(st.session_state.cover_letter)
        st.download_button(
            "⬇️ Download Cover Letter",
            st.session_state.cover_letter,
            "cover_letter.txt",
        )

# ─── TAB 9: CITATION SUGGESTIONS ──────────────────────────────────────────────
with tabs[9]:
    st.subheader("Citation Suggester")
    st.caption("Paste a claim or statement; get relevant citation suggestions.")

    claim = st.text_area(
        "Statement or claim",
        height=100,
        placeholder="e.g., 'Competency-based education has been shown to improve clinical reasoning skills in dental students.'"
    )
    field = st.text_input("Research field", value="Health Professions Education")

    if st.button("🔖 Suggest Citations"):
        if not claim.strip():
            st.warning("Please enter a claim.")
        else:
            with st.spinner("Finding relevant citations..."):
                suggestions = suggest_citations(claim, field)
                st.session_state.citation_suggestions = suggestions
                st.rerun()

    if st.session_state.citation_suggestions:
        st.markdown(st.session_state.citation_suggestions)

# ─── TAB 10: REPORTING CHECKLIST ──────────────────────────────────────────────
with tabs[10]:
    st.subheader("Reporting Guideline Checklist")
    st.caption("Check adherence to standard reporting guidelines.")

    guideline = st.selectbox(
        "Select guideline",
        ["CONSORT (RCTs)", "SRQR (Qualitative)", "COREQ (Interviews/Focus Groups)", "PRISMA (Systematic Reviews)", "STARD (Diagnostic Accuracy)"]
    )

    if st.button("✔️ Check Adherence"):
        if not st.session_state.pdf_text:
            st.warning("Upload a manuscript first.")
        else:
            with st.spinner(f"Checking {guideline} adherence..."):
                checklist = check_guideline(st.session_state.pdf_text, guideline)
                st.session_state.checklist_report = checklist
                st.rerun()

    if st.session_state.checklist_report:
        st.markdown(st.session_state.checklist_report)

# ─── TAB 11: FEEDBACK ─────────────────────────────────────────────────────────
with tabs[11]:
    st.subheader("Platform Feedback")
    st.caption("Help us improve DentEdTech™.")

    if st.session_state.feedback_submitted:
        st.success("✅ Thank you — your feedback has been recorded.")
        if st.button("Submit another response"):
            st.session_state.feedback_submitted = False
            st.rerun()
    else:
        with st.form("feedback_form", clear_on_submit=True):
            col_name, col_role = st.columns(2)
            with col_name:
                fb_name = st.text_input("Your name (optional)", placeholder="Dr. Jane Smith")
            with col_role:
                fb_role = st.selectbox("Your role", [
                    "Journal Editor", "Managing Editor", "Academic Researcher",
                    "Dental Educator", "HPE Educator", "Publishing Professional", "Other"
                ])
            fb_institution = st.text_input("Institution / Journal (optional)")
            fb_overall = st.select_slider(
                "Overall satisfaction",
                options=["Very dissatisfied", "Dissatisfied", "Neutral", "Satisfied", "Very satisfied"],
                value="Satisfied"
            )
            fb_accuracy = st.select_slider(
                "How accurate was the AI review?",
                options=["Very inaccurate", "Somewhat inaccurate", "Neutral", "Somewhat accurate", "Very accurate"],
                value="Somewhat accurate"
            )
            fb_worked_well = st.text_area("What worked well?", height=80)
            fb_improvements = st.text_area("What needs improvement?", height=80)
            fb_features = st.text_area("Feature requests?", height=80)

            if st.form_submit_button("Submit Feedback →", use_container_width=True):
                st.session_state.feedback_submitted = True
                st.rerun()