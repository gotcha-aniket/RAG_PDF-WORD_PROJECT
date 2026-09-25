"""
SynapseIQ — Enterprise Multi-Source Knowledge Intelligence Platform
A high-performance Chroma + Gemini RAG pipeline featuring Rate-Shield™ protection,
multimodal source ingestion (PDF, Word, PowerPoint, Web, YouTube), and MMR retrieval.

Run with:
    streamlit run app.py
"""

import os
import tempfile
import time
import re
import base64

import streamlit as st
from dotenv import load_dotenv

from langchain_core.documents import Document
from langchain_core.embeddings import Embeddings
from langchain_google_genai import ChatGoogleGenerativeAI, GoogleGenerativeAIEmbeddings
from langchain_community.document_loaders import PyPDFLoader, WebBaseLoader, YoutubeLoader
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_chroma import Chroma

from docx import Document as DocxReader
from pptx import Presentation

load_dotenv()

# ==================================================
# LOGO HELPER
# ==================================================

LOGO_PATH = os.path.join(os.path.dirname(__file__), "assets", "logo.png")

def get_logo_base64() -> str:
    """Returns base64 encoded logo string if available."""
    if os.path.exists(LOGO_PATH):
        try:
            with open(LOGO_PATH, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            return ""
    return ""

LOGO_B64 = get_logo_base64()

# ==================================================
# PAGE CONFIG + STYLING
# ==================================================

st.set_page_config(
    page_title="SynapseIQ • Enterprise Knowledge Engine",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

        :root {
            --bg-dark: #080C16;
            --bg-card: #0F172A;
            --bg-card-hover: #1E293B;
            --border-glass: rgba(255, 255, 255, 0.08);
            --border-glass-glow: rgba(56, 189, 248, 0.35);
            --accent-cyan: #06B6D4;
            --accent-cyan-glow: rgba(6, 182, 212, 0.35);
            --accent-sky: #38BDF8;
            --accent-indigo: #6366F1;
            --accent-violet: #8B5CF6;
            --text-primary: #F8FAFC;
            --text-secondary: #94A3B8;
            --text-dim: #64748B;
            --success: #10B981;
            --warning: #F59E0B;
        }

        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', sans-serif;
            color: var(--text-primary);
        }

        .stApp {
            background: radial-gradient(circle at 12% 10%, rgba(6, 182, 212, 0.07) 0%, transparent 40%),
                        radial-gradient(circle at 88% 85%, rgba(139, 92, 246, 0.08) 0%, transparent 45%),
                        radial-gradient(circle at 50% 50%, rgba(15, 23, 42, 0.5) 0%, transparent 80%),
                        #080C16;
            background-attachment: fixed;
        }

        .main .block-container {
            max-width: 900px;
            padding-top: 1.8rem;
            padding-bottom: 4rem;
        }

        /* ---------- Navbar & Header ---------- */
        .synapse-navbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 1.1rem 1.6rem;
            margin-bottom: 1.8rem;
            background: rgba(15, 23, 42, 0.7);
            backdrop-filter: blur(18px);
            border: 1px solid var(--border-glass);
            border-radius: 18px;
            box-shadow: 0 16px 36px -12px rgba(0, 0, 0, 0.55);
            position: relative;
            overflow: hidden;
        }

        .synapse-navbar::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0; height: 2px;
            background: linear-gradient(90deg, var(--accent-cyan), var(--accent-indigo), var(--accent-violet));
        }

        .synapse-brand {
            display: flex;
            align-items: center;
            gap: 1.15rem;
        }

        .synapse-logo-box {
            width: 54px;
            height: 54px;
            border-radius: 15px;
            padding: 2px;
            background: linear-gradient(135deg, var(--accent-cyan), var(--accent-violet));
            box-shadow: 0 0 24px var(--accent-cyan-glow);
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
        }

        .synapse-logo-img {
            width: 100%;
            height: 100%;
            border-radius: 13px;
            object-fit: cover;
        }

        .synapse-title-row {
            display: flex;
            align-items: center;
            flex-wrap: wrap;
            gap: 0.65rem;
            margin-bottom: 0.15rem;
        }

        .synapse-title {
            font-family: 'Outfit', sans-serif;
            font-size: 1.85rem;
            font-weight: 800;
            color: #FFFFFF;
            letter-spacing: -0.02em;
        }

        .synapse-highlight {
            background: linear-gradient(135deg, #38BDF8 20%, #A855F7 80%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }

        .synapse-badge {
            background: rgba(56, 189, 248, 0.12);
            color: #38BDF8;
            border: 1px solid rgba(56, 189, 248, 0.35);
            padding: 0.2rem 0.65rem;
            border-radius: 20px;
            font-size: 0.72rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }

        .synapse-status-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            background: rgba(16, 185, 129, 0.12);
            color: #34D399;
            border: 1px solid rgba(16, 185, 129, 0.28);
            padding: 0.22rem 0.7rem;
            border-radius: 20px;
            font-size: 0.74rem;
            font-weight: 600;
        }

        .synapse-status-pill.unready {
            background: rgba(245, 158, 11, 0.12);
            color: #FBBF24;
            border-color: rgba(245, 158, 11, 0.3);
        }

        .status-dot {
            width: 7px;
            height: 7px;
            background: #34D399;
            border-radius: 50%;
            box-shadow: 0 0 10px #34D399;
            animation: pulse-green 2s infinite;
        }

        .status-dot.unready {
            background: #FBBF24;
            box-shadow: 0 0 10px #FBBF24;
        }

        @keyframes pulse-green {
            0% { transform: scale(0.9); opacity: 0.7; }
            50% { transform: scale(1.25); opacity: 1; }
            100% { transform: scale(0.9); opacity: 0.7; }
        }

        .synapse-subtitle {
            font-size: 0.88rem;
            color: var(--text-secondary);
            font-weight: 400;
        }

        /* ---------- Sidebar Styling ---------- */
        section[data-testid="stSidebar"] {
            background: #0B101E !important;
            border-right: 1px solid var(--border-glass) !important;
        }

        section[data-testid="stSidebar"] .block-container {
            padding-top: 1.2rem;
            padding-left: 1.2rem;
            padding-right: 1.2rem;
        }

        .sidebar-brand-card {
            text-align: center;
            padding: 1.3rem 1rem;
            margin-bottom: 1.4rem;
            background: linear-gradient(180deg, rgba(30, 41, 59, 0.5) 0%, rgba(15, 23, 42, 0.75) 100%);
            border: 1px solid var(--border-glass);
            border-radius: 16px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.4);
        }

        .sidebar-brand-logo {
            width: 76px;
            height: 76px;
            border-radius: 18px;
            margin: 0 auto 0.8rem auto;
            display: block;
            box-shadow: 0 0 28px rgba(6, 182, 212, 0.4);
            border: 1px solid rgba(255, 255, 255, 0.15);
        }

        .sidebar-app-name {
            font-family: 'Outfit', sans-serif;
            font-size: 1.45rem;
            font-weight: 800;
            color: #FFFFFF;
            letter-spacing: -0.01em;
        }

        .sidebar-app-desc {
            font-size: 0.78rem;
            color: #94A3B8;
            margin-top: 0.25rem;
        }

        /* Metric Dashboard Grid */
        .metric-deck {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 0.5rem;
            margin: 0.8rem 0 1.2rem 0;
        }

        .metric-card {
            background: rgba(15, 23, 42, 0.65);
            border: 1px solid var(--border-glass);
            border-radius: 10px;
            padding: 0.6rem 0.3rem;
            text-align: center;
            transition: border-color 0.2s ease;
        }

        .metric-card:hover {
            border-color: var(--border-glass-glow);
        }

        .metric-val {
            font-family: 'Outfit', sans-serif;
            font-size: 1.25rem;
            font-weight: 700;
            color: var(--accent-sky);
        }

        .metric-lbl {
            font-size: 0.65rem;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--text-dim);
            margin-top: 0.15rem;
        }

        /* Source Badge Chips */
        .source-tag {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            background: rgba(30, 41, 59, 0.55);
            border: 1px solid rgba(255, 255, 255, 0.09);
            border-radius: 8px;
            padding: 0.35rem 0.65rem;
            margin: 0.25rem 0.35rem 0.25rem 0;
            font-size: 0.78rem;
            color: #E2E8F0;
            transition: all 0.2s ease;
        }

        .source-tag:hover {
            background: rgba(56, 189, 248, 0.15);
            border-color: rgba(56, 189, 248, 0.45);
            transform: translateY(-1px);
        }

        /* Hero Welcome Deck */
        .hero-deck {
            background: linear-gradient(135deg, rgba(30, 41, 59, 0.45) 0%, rgba(15, 23, 42, 0.7) 100%);
            border: 1px solid var(--border-glass);
            border-radius: 18px;
            padding: 1.8rem 2rem;
            margin-bottom: 2rem;
            box-shadow: 0 12px 36px -10px rgba(0, 0, 0, 0.5);
            position: relative;
            overflow: hidden;
        }

        .hero-deck::before {
            content: '';
            position: absolute;
            top: 0; left: 0; right: 0; height: 3px;
            background: linear-gradient(90deg, #06B6D4, #6366F1, #EC4899);
        }

        .hero-headline {
            font-family: 'Outfit', sans-serif;
            font-size: 1.45rem;
            font-weight: 700;
            color: #FFFFFF;
            margin-bottom: 0.4rem;
        }

        .hero-lead {
            font-size: 0.92rem;
            color: var(--text-secondary);
            max-width: 65ch;
            line-height: 1.55;
            margin-bottom: 1.4rem;
        }

        .feature-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 1rem;
            margin-top: 1rem;
        }

        .feature-card {
            background: rgba(15, 23, 42, 0.65);
            border: 1px solid var(--border-glass);
            border-radius: 12px;
            padding: 1.1rem;
            transition: transform 0.2s ease, border-color 0.2s ease;
        }

        .feature-card:hover {
            transform: translateY(-3px);
            border-color: var(--accent-sky);
        }

        .feature-icon-box {
            font-size: 1.6rem;
            margin-bottom: 0.6rem;
        }

        .feature-heading {
            font-family: 'Outfit', sans-serif;
            font-size: 0.98rem;
            font-weight: 600;
            color: #F8FAFC;
            margin-bottom: 0.3rem;
        }

        .feature-body {
            font-size: 0.8rem;
            color: var(--text-secondary);
            line-height: 1.45;
        }

        /* Chat Messages */
        [data-testid="stChatMessage"] {
            background: rgba(15, 23, 42, 0.45) !important;
            border: 1px solid var(--border-glass) !important;
            border-radius: 16px !important;
            padding: 1.1rem 1.4rem !important;
            margin-bottom: 1.1rem !important;
            backdrop-filter: blur(12px) !important;
        }

        [data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-user"]) {
            background: rgba(30, 41, 59, 0.35) !important;
            border-left: 3px solid #38BDF8 !important;
        }

        [data-testid="stChatMessage"]:has([data-testid="chatAvatarIcon-assistant"]) {
            background: rgba(15, 23, 42, 0.65) !important;
            border-left: 3px solid #A855F7 !important;
        }

        .source-chunk-card {
            background: rgba(15, 23, 42, 0.8);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-left: 3px solid #06B6D4;
            border-radius: 10px;
            padding: 0.85rem 1.1rem;
            margin-bottom: 0.65rem;
            font-size: 0.84rem;
            color: #CBD5E1;
            line-height: 1.55;
        }

        .source-chunk-meta {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 0.45rem;
            font-size: 0.74rem;
            color: #38BDF8;
            font-weight: 600;
            text-transform: uppercase;
            letter-spacing: 0.04em;
        }

        /* Buttons & Controls */
        .stButton button {
            border-radius: 10px !important;
            font-weight: 600 !important;
            font-family: 'Plus Jakarta Sans', sans-serif !important;
            transition: all 0.2s ease !important;
        }

        .stButton button[kind="primary"] {
            background: linear-gradient(135deg, #0284C7, #6366F1) !important;
            color: #FFFFFF !important;
            border: 1px solid rgba(255, 255, 255, 0.15) !important;
            box-shadow: 0 4px 18px rgba(2, 132, 199, 0.4) !important;
        }

        .stButton button[kind="primary"]:hover {
            background: linear-gradient(135deg, #0369A1, #4F46E5) !important;
            box-shadow: 0 6px 24px rgba(2, 132, 199, 0.6) !important;
            transform: translateY(-1px);
        }

        /* Inputs & Textareas */
        .stTextInput input, .stSelectbox select, [data-baseweb="select"] {
            border-radius: 10px !important;
        }

        /* Chat Input */
        [data-testid="stChatInput"] {
            border-radius: 18px !important;
            border: 1px solid rgba(56, 189, 248, 0.35) !important;
            background: rgba(15, 23, 42, 0.85) !important;
            box-shadow: 0 8px 28px rgba(0, 0, 0, 0.45) !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==================================================
# TOP NAVBAR / HERO BANNER
# ==================================================

is_ready = st.session_state.get("vector_db") is not None
status_badge_class = "" if is_ready else "unready"
status_dot_class = "" if is_ready else "unready"
status_text = "Vector Mesh: Online" if is_ready else "Vector Mesh: Awaiting Index"

logo_html = (
    f'<img src="data:image/png;base64,{LOGO_B64}" class="synapse-logo-img" alt="SynapseIQ Logo" />'
    if LOGO_B64
    else '<span style="font-size:1.8rem;">⚡</span>'
)

st.markdown(
    f"""
    <div class="synapse-navbar">
        <div class="synapse-brand">
            <div class="synapse-logo-box">
                {logo_html}
            </div>
            <div>
                <div class="synapse-title-row">
                    <span class="synapse-title">Synapse<span class="synapse-highlight">IQ</span></span>
                    <span class="synapse-badge">Enterprise RAG v2.4</span>
                    <span class="synapse-status-pill {status_badge_class}">
                        <span class="status-dot {status_dot_class}"></span> {status_text}
                    </span>
                </div>
                <div class="synapse-subtitle">Autonomous Multi-Source Knowledge Mesh & Contextual Reasoning Engine</div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ==================================================
# SESSION STATE INITIALIZATION
# ==================================================

if "raw_docs" not in st.session_state:
    st.session_state.raw_docs = []          # list of langchain Document objects
if "sources_list" not in st.session_state:
    st.session_state.sources_list = []      # list of {"name":..., "type":...}
if "vector_db" not in st.session_state:
    st.session_state.vector_db = None
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "indexed_chunks_count" not in st.session_state:
    st.session_state.indexed_chunks_count = 0
if "preset_prompt" not in st.session_state:
    st.session_state.preset_prompt = None

# ==================================================
# LOADER FUNCTIONS
# ==================================================

def load_pdf(file_bytes: bytes, filename: str):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        docs = PyPDFLoader(tmp_path).load()
        for d in docs:
            d.metadata["source"] = filename
            d.metadata["type"] = "pdf"
        return docs
    finally:
        os.unlink(tmp_path)


def load_docx(file_bytes: bytes, filename: str):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        doc = DocxReader(tmp_path)
        text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        return [Document(page_content=text, metadata={"source": filename, "type": "docx"})]
    finally:
        os.unlink(tmp_path)


def load_pptx(file_bytes: bytes, filename: str):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        prs = Presentation(tmp_path)
        docs = []
        for i, slide in enumerate(prs.slides, start=1):
            texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    texts.append(shape.text_frame.text)
            slide_text = "\n".join(t for t in texts if t.strip())
            if slide_text.strip():
                docs.append(Document(
                    page_content=slide_text,
                    metadata={"source": filename, "type": "pptx", "slide": i}
                ))
        return docs
    finally:
        os.unlink(tmp_path)


def load_website(url: str):
    docs = WebBaseLoader(url).load()
    for d in docs:
        d.metadata["source"] = url
        d.metadata["type"] = "website"
    return docs


def load_youtube(url: str):
    loader = YoutubeLoader.from_youtube_url(url, add_video_info=False)
    docs = loader.load()
    for d in docs:
        d.metadata["source"] = url
        d.metadata["type"] = "youtube"
    return docs


FILE_LOADERS = {
    "pdf": load_pdf,
    "docx": load_docx,
    "pptx": load_pptx,
}

# ==================================================
# RATE-SHIELD™ EMBEDDINGS & RETRY ENGINE
# ==================================================

class RateLimitedGeminiEmbeddings(Embeddings):
    """
    Wraps GoogleGenerativeAIEmbeddings with batching, delay, and exponential backoff
    to strictly adhere to Google Free Tier rate limits (100 requests / minute)
    and prevent RESOURCE_EXHAUSTED (429) errors.
    """
    def __init__(
        self,
        model: str = "gemini-embedding-001",
        google_api_key: str | None = None,
        batch_size: int = 20,
        delay_seconds: float = 0.6,
        status_container=None,
    ):
        resolved_key = google_api_key or os.getenv("GOOGLE_API_KEY")
        self.base_embeddings = GoogleGenerativeAIEmbeddings(
            model=model, google_api_key=resolved_key
        )
        self.batch_size = max(1, min(batch_size, 30))
        self.delay_seconds = delay_seconds
        self.status_container = status_container

    def embed_documents(self, texts: list[str]) -> list[list[float]]:
        all_embeddings: list[list[float]] = []
        total = len(texts)

        for i in range(0, total, self.batch_size):
            batch = texts[i : i + self.batch_size]
            batch_end = min(i + len(batch), total)
            if self.status_container:
                self.status_container.info(f"⚡ Rate-Shield™: Embedding chunks {i + 1}–{batch_end} of {total}...")

            max_retries = 5
            for attempt in range(max_retries):
                try:
                    embeddings_batch = self.base_embeddings.embed_documents(batch)
                    all_embeddings.extend(embeddings_batch)
                    break
                except Exception as e:
                    err_msg = str(e)
                    if ("429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg) and attempt < max_retries - 1:
                        wait_time = 5.0 * (attempt + 1)
                        match = re.search(r"retry in (\d+(?:\.\d+)?)s", err_msg, re.IGNORECASE)
                        if match:
                            wait_time = max(wait_time, float(match.group(1)) + 1.0)
                        if self.status_container:
                            self.status_container.warning(
                                f"⏳ Quota limit approached. Cooling down {wait_time:.1f}s before retry ({attempt + 1}/{max_retries})..."
                            )
                        time.sleep(wait_time)
                    else:
                        raise e

            # Brief pause between batches to prevent spiking RPM quota
            if i + self.batch_size < total and self.delay_seconds > 0:
                time.sleep(self.delay_seconds)

        return all_embeddings

    def embed_query(self, text: str) -> list[float]:
        max_retries = 4
        for attempt in range(max_retries):
            try:
                return self.base_embeddings.embed_query(text)
            except Exception as e:
                err_msg = str(e)
                if ("429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg) and attempt < max_retries - 1:
                    time.sleep(3.0 * (attempt + 1))
                else:
                    raise e


def generate_answer_with_retry(llm: ChatGoogleGenerativeAI, prompt: str, max_retries: int = 4) -> str:
    """Invokes LLM with retry on 429 and safely extracts response text."""
    for attempt in range(max_retries):
        try:
            response = llm.invoke(prompt)
            if hasattr(response, "content"):
                content = response.content
                if isinstance(content, str):
                    return content
                elif isinstance(content, list):
                    text_parts = []
                    for part in content:
                        if isinstance(part, dict) and "text" in part:
                            text_parts.append(part["text"])
                        elif hasattr(part, "text"):
                            text_parts.append(part.text)
                        else:
                            text_parts.append(str(part))
                    return "".join(text_parts)
                else:
                    return str(content)
            return str(response)
        except Exception as e:
            err_msg = str(e)
            if ("429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg) and attempt < max_retries - 1:
                wait_time = 4.0 * (attempt + 1)
                match = re.search(r"retry in (\d+(?:\.\d+)?)s", err_msg, re.IGNORECASE)
                if match:
                    wait_time = max(wait_time, float(match.group(1)) + 1.0)
                time.sleep(wait_time)
            else:
                raise e

# ==================================================
# SIDEBAR — CONTROL DECK
# ==================================================

with st.sidebar:
    # Brand Identity Card
    sidebar_logo_html = (
        f'<img src="data:image/png;base64,{LOGO_B64}" class="sidebar-brand-logo" alt="SynapseIQ Logo" />'
        if LOGO_B64
        else '<div style="font-size:3rem; margin-bottom:0.5rem;">⚡</div>'
    )
    st.markdown(
        f"""
        <div class="sidebar-brand-card">
            {sidebar_logo_html}
            <div class="sidebar-app-name">Synapse<span style="color:#38BDF8;">IQ</span></div>
            <div class="sidebar-app-desc">Multi-Source Neural Knowledge Mesh</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Metrics Overview
    sources_cnt = len(st.session_state.sources_list)
    chunks_cnt = st.session_state.indexed_chunks_count or len(st.session_state.raw_docs)
    db_status = "READY" if st.session_state.vector_db is not None else "EMPTY"
    
    st.markdown(
        f"""
        <div class="metric-deck">
            <div class="metric-card">
                <div class="metric-val">{sources_cnt}</div>
                <div class="metric-lbl">Sources</div>
            </div>
            <div class="metric-card">
                <div class="metric-val">{chunks_cnt}</div>
                <div class="metric-lbl">Chunks</div>
            </div>
            <div class="metric-card">
                <div class="metric-val" style="color:{'#34D399' if db_status=='READY' else '#F59E0B'}">{db_status}</div>
                <div class="metric-lbl">Chroma DB</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("⚙️ Engine Configuration")
    api_key = st.text_input(
        "Google API Key",
        value=os.getenv("GOOGLE_API_KEY", ""),
        type="password",
        help="Loaded automatically from .env if present. You can override it here.",
    )

    selected_model = st.selectbox(
        "Gemini Generation Model",
        options=[
            "gemini-3.8-flash",
            "gemini-3.1-flash-lite",
            "gemini-3.5-flash",
            "gemini-flash-latest",
        ],
        index=0,
        help=(
            "• gemini-3.8-flash: State-of-the-art fast reasoning\n"
            "• gemini-3.1-flash-lite: Best for strict quota limits (high RPM, low token burn)\n"
            "• gemini-3.5-flash: Standard Flash"
        ),
    )

    st.divider()
    st.subheader("📥 Ingestion Hub")
    tab_files, tab_web, tab_yt = st.tabs(["📁 Documents", "🌐 Web URL", "📺 YouTube"])

    with tab_files:
        uploaded_files = st.file_uploader(
            "Upload PDF, Word (.docx), or PowerPoint (.pptx)",
            type=["pdf", "docx", "pptx"],
            accept_multiple_files=True,
        )
        if st.button("Ingest Files into Mesh", use_container_width=True):
            if not uploaded_files:
                st.warning("Select at least one file first.")
            else:
                for f in uploaded_files:
                    ext = f.name.split(".")[-1].lower()
                    loader_fn = FILE_LOADERS.get(ext)
                    if not loader_fn:
                        st.error(f"Unsupported file format: {f.name}")
                        continue
                    try:
                        docs = loader_fn(f.read(), f.name)
                        st.session_state.raw_docs.extend(docs)
                        st.session_state.sources_list.append({"name": f.name, "type": ext})
                        st.success(f"Ingested: {f.name} ({len(docs)} section(s))")
                    except Exception as e:
                        st.error(f"Failed to load {f.name}: {e}")

    with tab_web:
        website_url = st.text_input("Live Website URL", placeholder="https://example.com/research-paper")
        if st.button("Ingest Website", use_container_width=True):
            if not website_url.strip():
                st.warning("Enter a valid URL.")
            else:
                try:
                    docs = load_website(website_url.strip())
                    st.session_state.raw_docs.extend(docs)
                    st.session_state.sources_list.append({"name": website_url.strip(), "type": "website"})
                    st.success(f"Ingested webpage ({len(docs)} section(s))")
                except Exception as e:
                    st.error(f"Failed to load URL: {e}")

    with tab_yt:
        yt_url = st.text_input("YouTube Video URL", placeholder="https://www.youtube.com/watch?v=...")
        if st.button("Ingest Video Transcript", use_container_width=True):
            if not yt_url.strip():
                st.warning("Enter a YouTube URL.")
            else:
                try:
                    docs = load_youtube(yt_url.strip())
                    st.session_state.raw_docs.extend(docs)
                    st.session_state.sources_list.append({"name": yt_url.strip(), "type": "youtube"})
                    st.success("Ingested video transcript successfully")
                except Exception as e:
                    st.error(f"Failed to retrieve video transcript: {e}")

    st.divider()
    st.subheader("🗂️ Active Knowledge Sources")
    if st.session_state.sources_list:
        for s in st.session_state.sources_list:
            icon = {"pdf": "📄", "docx": "📝", "pptx": "📊", "website": "🌐", "youtube": "▶️"}.get(s["type"], "📎")
            st.markdown(
                f'<span class="source-tag">{icon} <b>{s["type"].upper()}</b> {s["name"][:32]}</span>',
                unsafe_allow_html=True,
            )
        if st.button("Reset Knowledge Base", use_container_width=True):
            st.session_state.raw_docs = []
            st.session_state.sources_list = []
            st.session_state.vector_db = None
            st.session_state.indexed_chunks_count = 0
            st.rerun()
    else:
        st.caption("No sources ingested yet.")

    st.divider()
    st.subheader("🎯 Precision Retrieval Tuning")
    chunk_size = st.slider("Chunk Size (characters)", 500, 2000, 1000, step=100)
    chunk_overlap = st.slider("Chunk Overlap (characters)", 0, 500, 200, step=50)
    top_k = st.slider(
        "Chunks to Retrieve (k)",
        2, 20, 4,
        help="Recommended 3–5 chunks. Keeping k lower saves token quota and avoids rate limits.",
    )
    fetch_k = st.slider("Candidates Before MMR (fetch_k)", top_k, 40, max(15, top_k * 2))

    build_clicked = st.button("⚡ Build / Rebuild Knowledge Index", type="primary", use_container_width=True)

    if st.session_state.vector_db is not None:
        st.success("Knowledge Mesh is indexed and active.")

    if st.session_state.chat_history:
        st.divider()
        st.subheader("💬 Chat Management")
        chat_export = "# SynapseIQ Conversation Export\n\n"
        for q, a, _ in st.session_state.chat_history:
            chat_export += f"**User:**\n{q}\n\n**SynapseIQ:**\n{a}\n\n---\n\n"
        st.download_button(
            label="📥 Export Chat History (.md)",
            data=chat_export,
            file_name="synapseiq_chat_export.md",
            mime="text/markdown",
            use_container_width=True,
        )
        if st.button("Clear Chat Conversation", use_container_width=True):
            st.session_state.chat_history = []
            st.rerun()

# ==================================================
# BUILD VECTOR INDEX
# ==================================================

if build_clicked:
    resolved_api_key = api_key or os.getenv("GOOGLE_API_KEY")
    if not resolved_api_key:
        st.sidebar.error("Provide a Google API key in the Setup section.")
    elif not st.session_state.raw_docs:
        st.sidebar.error("Add at least one source document first.")
    else:
        status_box = st.sidebar.empty()
        with st.spinner("Synthesizing and indexing knowledge mesh..."):
            splitter = RecursiveCharacterTextSplitter(
                chunk_size=chunk_size, chunk_overlap=chunk_overlap
            )
            chunks = splitter.split_documents(st.session_state.raw_docs)

            embeddings = RateLimitedGeminiEmbeddings(
                model="gemini-embedding-001",
                google_api_key=resolved_api_key,
                batch_size=20,
                delay_seconds=0.6,
                status_container=status_box,
            )
            try:
                st.session_state.vector_db = Chroma.from_documents(
                    documents=chunks, embedding=embeddings
                )
                st.session_state.indexed_chunks_count = len(chunks)
                st.session_state.chat_history = []
                status_box.empty()
                st.sidebar.success(f"Indexed {len(chunks)} chunks across {len(st.session_state.sources_list)} source(s)!")
                st.rerun()
            except Exception as e:
                status_box.empty()
                st.sidebar.error(f"Failed to build index: {e}")

# ==================================================
# EMPTY STATE / WELCOME HERO DECK
# ==================================================

if not st.session_state.chat_history:
    st.markdown(
        """
        <div class="hero-deck">
            <div class="hero-headline">Welcome to SynapseIQ</div>
            <div class="hero-lead">
                Ingest multi-format files, web articles, and video transcripts into a single unified vector mesh.
                Ask complex questions and receive grounded, cited answers backed by Rate-Shield™ protection.
            </div>
            <div class="feature-grid">
                <div class="feature-card">
                    <div class="feature-icon-box">📚</div>
                    <div class="feature-heading">Omnichannel Ingestion</div>
                    <div class="feature-body">Ingest PDFs, Word docs, PowerPoint decks, web pages, and YouTube video captions into one unified index.</div>
                </div>
                <div class="feature-card">
                    <div class="feature-icon-box">🛡️</div>
                    <div class="feature-heading">Rate-Shield™ Embeddings</div>
                    <div class="feature-body">Autonomous 20-chunk micro-batching and adaptive exponential backoff to eliminate 429 quota exhaustion.</div>
                </div>
                <div class="feature-card">
                    <div class="feature-icon-box">🎯</div>
                    <div class="feature-heading">MMR Semantic Retrieval</div>
                    <div class="feature-body">Maximal Marginal Relevance ranking ensures diverse, non-redundant contextual evidence with exact source chunk citations.</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.session_state.vector_db is not None:
        st.markdown('<p style="font-size:0.9rem; color:#94A3B8; font-weight:600; margin-bottom:0.6rem;">💡 Suggested Explorations:</p>', unsafe_allow_html=True)
        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            if st.button("📊 Synthesize Core Themes", use_container_width=True):
                st.session_state.preset_prompt = "Provide a comprehensive synthesis of the key themes, conclusions, and insights across all sources."
                st.rerun()
        with col_p2:
            if st.button("🔍 Contrast Methodologies", use_container_width=True):
                st.session_state.preset_prompt = "Identify and compare the main arguments, methodologies, or data points discussed in the sources."
                st.rerun()
        with col_p3:
            if st.button("📋 Extract Actionable Takeaways", use_container_width=True):
                st.session_state.preset_prompt = "Extract the top 5 most actionable insights, facts, or recommendations from the provided sources."
                st.rerun()

# ==================================================
# CHAT CONVERSATION STREAM
# ==================================================

for q, a, sources in st.session_state.chat_history:
    with st.chat_message("user"):
        st.write(q)
    with st.chat_message("assistant"):
        st.write(a)
        if sources:
            with st.expander(f"🔍 Inspect {len(sources)} Verified Evidence Chunk(s)"):
                for i, (content, meta) in enumerate(sources, 1):
                    tag = meta.get("source", "unknown")
                    meta_type = meta.get("type", "source").upper()
                    st.markdown(
                        f"""
                        <div class="source-chunk-card">
                            <div class="source-chunk-meta">
                                <span>Evidence Chunk #{i}</span>
                                <span>[{meta_type}] {tag[:50]}</span>
                            </div>
                            {content}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

# ==================================================
# CHAT INPUT & EXECUTION
# ==================================================

question_input = st.chat_input("Ask any question across your ingested knowledge mesh...")

# Handle either chat input or preset inquiry click
query_to_run = question_input
if not query_to_run and st.session_state.preset_prompt:
    query_to_run = st.session_state.preset_prompt
    st.session_state.preset_prompt = None

if query_to_run:
    resolved_api_key = api_key or os.getenv("GOOGLE_API_KEY")
    if not st.session_state.vector_db:
        st.error("⚠️ Please add sources and click '⚡ Build / Rebuild Knowledge Index' in the sidebar first.")
    elif not resolved_api_key:
        st.error("⚠️ Please provide a Google API Key in the sidebar Setup.")
    else:
        with st.chat_message("user"):
            st.write(query_to_run)

        with st.chat_message("assistant"):
            with st.spinner("Analyzing semantic evidence..."):
                retriever = st.session_state.vector_db.as_retriever(
                    search_type="mmr",
                    search_kwargs={"k": top_k, "fetch_k": fetch_k},
                )
                retrieved_docs = retriever.invoke(query_to_run)
                context = "\n\n".join(
                    f"[Source: {d.metadata.get('source', 'unknown')} | Type: {d.metadata.get('type', 'doc')}]\n{d.page_content}"
                    for d in retrieved_docs
                )

                prompt = f"""
You are SynapseIQ, an elite enterprise research assistant.
Answer the user's question using ONLY the verified evidence provided in the context below.

Guidelines:
- If the exact answer cannot be determined from the context, say:
  "I could not find sufficient information in the provided sources to answer this question."
- Always maintain precision and cite relevant document names or types when making assertions.
- Do not fabricate or extrapolate information beyond the sources.

Context:
{context}

User Question:
{query_to_run}
"""
                try:
                    llm = ChatGoogleGenerativeAI(
                        model=selected_model,
                        google_api_key=resolved_api_key,
                        max_retries=3,
                    )
                    answer = generate_answer_with_retry(llm, prompt)
                    st.write(answer)

                    sources = [(d.page_content, d.metadata) for d in retrieved_docs]
                    if sources:
                        with st.expander(f"🔍 Inspect {len(sources)} Verified Evidence Chunk(s)"):
                            for i, (content, meta) in enumerate(sources, 1):
                                tag = meta.get("source", "unknown")
                                meta_type = meta.get("type", "source").upper()
                                st.markdown(
                                    f"""
                                    <div class="source-chunk-card">
                                        <div class="source-chunk-meta">
                                            <span>Evidence Chunk #{i}</span>
                                            <span>[{meta_type}] {tag[:50]}</span>
                                        </div>
                                        {content}
                                    </div>
                                    """,
                                    unsafe_allow_html=True,
                                )

                    st.session_state.chat_history.append((query_to_run, answer, sources))

                except Exception as e:
                    err_msg = str(e)
                    if "429" in err_msg or "RESOURCE_EXHAUSTED" in err_msg:
                        st.error(
                            "⚠️ **Gemini Rate Limit (Quota) Throttled**\n\n"
                            "Google's free tier has temporarily throttled requests. Recommended actions:\n"
                            "- Wait 10–15 seconds and re-submit your prompt.\n"
                            "- Switch to **`gemini-3.1-flash-lite`** in the sidebar (higher limits & lower token footprint).\n"
                            "- Decrease **Chunks to Retrieve (k)** slider in sidebar retrieval settings."
                        )
                    else:
                        st.error(f"Error generating answer: {e}")
