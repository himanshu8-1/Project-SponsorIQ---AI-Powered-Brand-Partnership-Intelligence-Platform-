"""
================================================================================
 SponsorIQ — Brand Partnership Intelligence Platform
================================================================================
 A closed-data sponsorship recommendation system with optional AI mode.

 GOLDEN RULES (enforced throughout the code):
   · NO EXTERNAL DATA IN DATASET MODE     · NO HALLUCINATION IN DATASET MODE
   · NO INVENTED SCORES IN DATASET MODE   · NO FORCED RECOMMENDATIONS
   · AI MODE IS EXPLICIT & LABELLED       · NO SILENT FALLBACKS

 Every value displayed carries a source label:
   [DB]         Known from internal dataset
   [CALC]       Calculated from dataset
   [DECK]       Extracted from uploaded deck
   [USER]       Entered by the user
   [AI]         AI-Extracted / AI-Suggested
   [NA]         Data Not Available
================================================================================
"""

# ==============================================================================
# SECTION 1 — IMPORTS & GLOBAL CONFIG
# ==============================================================================
from __future__ import annotations
import io
import os
import re
import math
import json
import hashlib
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

# ---------- Optional dependencies (graceful degradation) ----------
try:
    from pypdf import PdfReader
    PDF_OK = True
except Exception:
    try:
        from PyPDF2 import PdfReader
        PDF_OK = True
    except Exception:
        PDF_OK = False

try:
    from pptx import Presentation
    PPTX_OK = True
except Exception:
    PPTX_OK = False

try:
    import docx
    DOCX_OK = True
except Exception:
    DOCX_OK = False

try:
    import mysql.connector as mysql_conn
    MYSQL_OK = True
except Exception:
    MYSQL_OK = False

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

# ---------- Optional LLM SDKs ----------
try:
    from openai import OpenAI
    import openai as _openai_module
    OPENAI_OK = True
    try:
        OPENAI_VERSION = _openai_module.__version__
    except Exception:
        OPENAI_VERSION = "unknown"
except Exception:
    OPENAI_OK = False
    OPENAI_VERSION = "not installed"

try:
    import anthropic
    ANTHROPIC_OK = True
except Exception:
    ANTHROPIC_OK = False


# ==============================================================================
# SECTION 2 — CONSTANTS · SOURCE LABELS · CONFIGURABLE WEIGHTS
# ==============================================================================

APP_TITLE = "SponsorIQ"
APP_TAGLINE = "Brand Partnership Intelligence Platform"
SCORING_VERSION = "v1.0-deterministic"
DATASET_VERSION_DEFAULT = "unknown"

# Source labels — every value in the UI must use one of these
SRC_DB    = "Internal Database"
SRC_DECK  = "Uploaded Deck"
SRC_CALC  = "Calculated"
SRC_USER  = "User Input"
SRC_NA    = "Data Not Available"
SRC_AI       = "AI-Suggested"
SRC_AI_EXTR  = "AI-Extracted"

# Default scoring weights (baked in — no UI to change them)
DEFAULT_WEIGHTS: Dict[str, float] = {
    "audience":    0.25,
    "budget":      0.20,
    "industry":    0.20,
    "geography":   0.15,
    "brand_value": 0.10,
    "competition": 0.10,
}

# Hard-filter thresholds
DEFAULT_THRESHOLDS: Dict[str, float] = {
    "budget_min_ratio":     0.40,
    "min_data_completeness":0.55,
    "min_dimensions_scored":4,
    "min_match_for_recommend": 65.0,
    "min_match_for_potential": 50.0,
}

# Recency multipliers for historical sponsorship
DEFAULT_RECENCY = {
    "0-12_months":  1.00,
    "12-24_months": 0.80,
    "24-36_months": 0.60,
    "36+_months":   0.40,
}

# Audience canonical tags
AUDIENCE_TAGS = {
    "youth":            ["youth", "gen z", "gen-z", "18-30", "young", "student"],
    "urban_millennials":["millennial", "urban", "metro"],
    "families":         ["family", "families", "household", "parents", "kids"],
    "rural":            ["rural", "semi-urban", "tier 2", "tier 3", "bharat", "village"],
    "professionals":    ["professional", "corporate", "cxo", "employee", "working"],
    "mass":             ["mass", "general", "pan-india", "broad", "everyone"],
    "premium":          ["premium", "affluent", "luxury", "hni", "hnw", "elite"],
    "b2b":              ["b2b", "enterprise", "business", "trade", "dealer", "institutional"],
    "sports_fans":      ["sports fan", "cricket fan", "football fan", "fan"],
    "music_fans":       ["music", "edm", "concert", "festival goer"],
    "fitness":          ["fitness", "runner", "running", "gym", "marathon", "athlete"],
    "foodies":          ["foodie", "food lover", "culinary"],
}

# Industry ↔ event-category affinity
INDUSTRY_EVENT_AFFINITY: Dict[str, Dict[str, float]] = {
    "Beauty & Personal Care": {
        "Sports - Cricket":        0.75,
        "Sports - Kabaddi":        0.70,
        "Sports - Football":       0.75,
        "Music - EDM Festival":    0.90,
        "Music - Multi-genre Festival": 0.90,
        "Fashion - Trade & Runway Show": 0.95,
        "Fashion - Runway Show":   0.95,
        "Entertainment - Film Awards":   0.90,
        "Entertainment - Film Festival": 0.85,
        "Cultural Festival":       0.80,
        "Religious - Pilgrimage":  0.60,
        "Business - Conference":   0.40,
        "Trade Show - Automobile": 0.35,
    },
    "FMCG": {
        "Sports - Cricket":        0.95,
        "Sports - Kabaddi":        0.90,
        "Sports - Football":       0.85,
        "Music - EDM Festival":    0.80,
        "Music - Multi-genre Festival": 0.80,
        "Entertainment - Film Awards":   0.85,
        "Cultural Festival":       0.90,
        "Religious - Pilgrimage":  0.90,
        "Food & Wine - Music & Wine Festival": 0.85,
        "Food - Culinary Festival": 0.95,
        "Retail - E-commerce Sale Event": 0.85,
        "Business - Conference":   0.45,
    },
    "Banking & FinTech": {
        "Sports - Cricket":        0.95,
        "Sports - Kabaddi":        0.80,
        "Sports - Football":       0.85,
        "Business - Conference":   0.90,
        "Business - Investment Summit": 0.95,
        "Business - Tech Conference": 0.90,
        "Entertainment - Film Awards": 0.70,
        "Cultural Festival":       0.65,
        "Trade Show - General Commerce": 0.70,
        "Sports - Marathon":       0.85,
    },
    "Fashion & Lifestyle": {
        "Fashion - Trade & Runway Show": 1.00,
        "Fashion - Runway Show":   1.00,
        "Music - EDM Festival":    0.85,
        "Music - Multi-genre Festival": 0.85,
        "Entertainment - Film Awards": 0.95,
        "Entertainment - Film Festival": 0.90,
        "Culture - Art Fair":      0.85,
        "Sports - Cricket":        0.70,
        "Retail - E-commerce Sale Event": 0.90,
        "Retail - Fashion E-commerce Sale": 0.95,
    },
    "Automobile & EV": {
        "Sports - Cricket":        0.85,
        "Sports - Motorsport":     1.00,
        "Trade Show - Automobile": 1.00,
        "Sports - Football":       0.75,
        "Business - Conference":   0.55,
        "Music - EDM Festival":    0.65,
    },
    "Consumer Electronics": {
        "Sports - Cricket":        0.85,
        "Sports - Kabaddi":        0.70,
        "Sports - Football":       0.75,
        "Music - EDM Festival":    0.90,
        "Music - Multi-genre Festival": 0.90,
        "Entertainment - Film Awards": 0.85,
        "Trade Show - Telecom & Tech": 0.95,
        "Business - Tech Conference": 0.95,
        "Retail - E-commerce Sale Event": 0.90,
        "Entertainment - Pop Culture Convention": 0.95,
    },
    "Health & Wellness": {
        "Sports - Marathon":       1.00,
        "Sports - Cricket":        0.70,
        "Sports - Multi-discipline": 0.85,
        "Sports - Athletics":      0.95,
        "Business - Conference":   0.55,
        "Food - Culinary Festival": 0.70,
    },
    "Home & Kitchen": {
        "Cultural Festival":       0.75,
        "Religious - Pilgrimage":  0.70,
        "Trade Show - General Commerce": 0.80,
        "Design - Trade Exhibition": 0.95,
        "Retail - E-commerce Sale Event": 0.80,
        "Food - Culinary Festival": 0.85,
    },
    "Real Estate & Infrastructure": {
        "Business - Conference":   0.80,
        "Business - Investment Summit": 0.85,
        "Trade Show - General Commerce": 0.70,
        "Sports - Cricket":        0.55,
        "Cultural Festival":       0.50,
    },
    "Others": {},
}


# ==============================================================================
# SECTION 2.5 — AI MODE CONFIG
# ==============================================================================

AI_PROVIDERS: Dict[str, Dict[str, Any]] = {
    "Google Gemini": {
        "kind": "openai",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "default_model": "gemini-2.5-flash",
        # NOTE: model names below are a conservative starting point.
        # If any 404s, open this URL in your browser (paste your key):
        #   https://generativelanguage.googleapis.com/v1beta/models?key=YOUR_KEY
        # and use the names it returns. You can also tick
        # "Type custom model name" in the sidebar to enter any name.
        "models": [
            # Confirmed available by Google's own 404 error message
            "gemini-3.1-pro-preview",
            # Gemini 2.5 family — still served for many projects
            "gemini-2.5-flash",
            "gemini-2.5-flash-lite",
            "gemini-2.5-pro",
            # Older — may be restricted or shut down
            "gemini-2.0-flash",
            "gemini-1.5-flash",
        ],
    },
    "OpenAI": {
        "kind": "openai",
        "base_url": "https://api.openai.com/v1",
        "default_model": "gpt-4o-mini",
        "models": ["gpt-4o-mini", "gpt-4o", "gpt-4-turbo", "gpt-3.5-turbo"],
    },
    "Groq": {
        "kind": "openai",
        "base_url": "https://api.groq.com/openai/v1",
        "default_model": "llama-3.3-70b-versatile",
        "models": ["llama-3.3-70b-versatile",
                   "llama-3.1-8b-instant",
                   "mixtral-8x7b-32768"],
    },
    "OpenRouter": {
        "kind": "openai",
        "base_url": "https://openrouter.ai/api/v1",
        "default_model": "meta-llama/llama-3.3-70b-instruct:free",
        "models": [
            "meta-llama/llama-3.3-70b-instruct:free",
            "google/gemma-2-9b-it:free",
            "mistralai/mistral-7b-instruct:free",
        ],
    },
    "Anthropic (Claude)": {
        "kind": "anthropic",
        "base_url": "https://api.anthropic.com",
        "default_model": "claude-3-5-sonnet-20241022",
        "models": ["claude-3-5-sonnet-20241022",
                   "claude-3-5-haiku-20241022",
                   "claude-3-opus-20240229"],
    },
    "OpenAI-compatible (custom)": {
        "kind": "openai",
        "base_url": "",
        "default_model": "",
        "models": [],
    },
}

# Bump this whenever AI_PROVIDERS changes — forces widget state reset
AI_PROVIDER_SCHEMA_VERSION = "v4"

AI_EXTRACTION_SYSTEM = (
    "You are a precise sponsorship-deck information extractor. "
    "You MUST return ONLY valid JSON. Never guess, never invent. "
    "If a field is not literally present in the deck, set it to null "
    "(or [] for arrays)."
)

AI_RECOMMENDATION_SYSTEM = (
    "You are a senior brand-sponsorship strategist working for an Indian "
    "sponsorship agency. You evaluate potential brand sponsors for a given "
    "event using commercial logic. You MUST return ONLY valid JSON matching "
    "the required schema. Never invent fake numbers, never invent brand "
    "names — every brand you list must actually exist and do sponsorships. "
    "Be candid: it is acceptable to return few or no recommendations."
)

AI_CATEGORY_LIST = [
    "Sports - Cricket", "Sports - Football", "Sports - Kabaddi",
    "Sports - Marathon", "Sports - Motorsport", "Sports - Athletics",
    "Sports - Multi-discipline", "Sports - Hockey", "Sports - Badminton",
    "Sports - Table Tennis", "Sports - Volleyball",
    "Music - EDM Festival", "Music - Multi-genre Festival",
    "Fashion - Runway Show", "Fashion - Trade & Runway Show",
    "Entertainment - Film Awards", "Entertainment - Film Festival",
    "Entertainment - Pop Culture Convention", "Entertainment - Comedy Festival",
    "Cultural Festival", "Religious - Pilgrimage",
    "Business - Conference", "Business - Investment Summit",
    "Business - Tech Conference",
    "Trade Show - Automobile", "Trade Show - Telecom & Tech",
    "Trade Show - General Commerce",
    "Culture - Art Fair", "Design - Trade Exhibition",
    "Food - Culinary Festival", "Food & Wine - Music & Wine Festival",
    "Retail - E-commerce Sale Event", "Retail - Fashion E-commerce Sale",
]

AI_ROLE_LIST = [
    "Title Sponsor", "Presenting Sponsor", "Co-Sponsor", "Powered By",
    "Associate Sponsor", "Official Partner", "Category Partner",
    "Media Partner", "Merchandise Partner",
]


# ==============================================================================
# SECTION 3 — DATA LAYER (Excel OR MySQL)
# ==============================================================================

SHEET_ALIASES = {
    "brands":             ["Indian_Brands_Dataset", "Brands", "brand_master", "Brand Master"],
    "events":             ["Indian_Events_Dataset", "Events", "event_master", "Event Master"],
    "sponsorships":       ["Sponsorship_History", "Sponsorships", "sponsorship_history"],
    "decision_makers":    ["Decision_Makers", "DecisionMakers"],
    "crm":                ["Sales_CRM", "CRM", "Sales CRM"],
    "campaigns":          ["Marketing_Campaigns", "Campaigns"],
    "benchmarks":         ["Industry_Benchmark", "IndustryBenchmark", "Benchmarks"],
    "calendar":           ["Calendar", "calendar_dim"],
}

@dataclass
class SponsorIQData:
    """Container for all internal data."""
    brands: pd.DataFrame = field(default_factory=pd.DataFrame)
    events: pd.DataFrame = field(default_factory=pd.DataFrame)
    sponsorships: pd.DataFrame = field(default_factory=pd.DataFrame)
    decision_makers: pd.DataFrame = field(default_factory=pd.DataFrame)
    crm: pd.DataFrame = field(default_factory=pd.DataFrame)
    campaigns: pd.DataFrame = field(default_factory=pd.DataFrame)
    benchmarks: pd.DataFrame = field(default_factory=pd.DataFrame)
    calendar: pd.DataFrame = field(default_factory=pd.DataFrame)
    dataset_version: str = DATASET_VERSION_DEFAULT
    loaded_from: str = ""
    loaded_at: str = ""
    load_errors: List[str] = field(default_factory=list)


def _first_existing_sheet(xls: pd.ExcelFile, aliases: List[str]) -> Optional[str]:
    names = {n.strip().lower(): n for n in xls.sheet_names}
    for a in aliases:
        if a.strip().lower() in names:
            return names[a.strip().lower()]
    return None


def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_from_excel(path: str) -> SponsorIQData:
    """Load every sheet we recognise from the workbook."""
    xls = pd.ExcelFile(path)
    data = SponsorIQData()
    data.loaded_from = f"Excel · {os.path.basename(path)}"
    data.loaded_at = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    with open(path, "rb") as f:
        data.dataset_version = hashlib.md5(f.read()).hexdigest()[:8]

    for key, aliases in SHEET_ALIASES.items():
        sheet = _first_existing_sheet(xls, aliases)
        if sheet is None:
            data.load_errors.append(f"Sheet not found for '{key}' (looked for {aliases})")
            continue
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
            df = _clean_columns(df)
            setattr(data, key, df)
        except Exception as e:
            data.load_errors.append(f"Failed to read '{sheet}': {e}")
    return data


def load_from_mysql() -> Optional[SponsorIQData]:
    """Optional: load from MySQL if credentials are provided in .env"""
    if not MYSQL_OK:
        return None
    host = os.getenv("MYSQL_HOST")
    user = os.getenv("MYSQL_USER")
    pwd  = os.getenv("MYSQL_PASSWORD")
    db   = os.getenv("MYSQL_DB")
    if not all([host, user, pwd, db]):
        return None
    try:
        cnx = mysql_conn.connect(host=host, user=user, password=pwd, database=db)
        data = SponsorIQData()
        data.loaded_from = f"MySQL · {host}/{db}"
        data.loaded_at = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
        data.dataset_version = "mysql-live"
        for key, aliases in SHEET_ALIASES.items():
            table = aliases[0].lower()
            try:
                df = pd.read_sql(f"SELECT * FROM `{table}`", cnx)
                setattr(data, key, _clean_columns(df))
            except Exception:
                pass
        cnx.close()
        return data
    except Exception:
        return None


def _find_local_excel() -> Optional[str]:
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        "Indian_Brands_Dataset.xlsx",
        "indian_brands_dataset.xlsx",
        "SponsorIQ.xlsx",
        "SponsorIQ_Dataset.xlsx",
    ]
    for c in candidates:
        p = os.path.join(here, c)
        if os.path.exists(p):
            return p
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


@st.cache_data(show_spinner=False, ttl=3600)
def get_data(excel_path: Optional[str] = None,
             uploaded_bytes: Optional[bytes] = None,
             uploaded_name: Optional[str] = None) -> SponsorIQData:
    """Load internal data — priority: path → uploaded → MySQL → local."""
    if excel_path and os.path.exists(excel_path):
        return load_from_excel(excel_path)

    if uploaded_bytes and uploaded_name:
        tmp = f"/tmp/sponsoriq_{uploaded_name}"
        with open(tmp, "wb") as f:
            f.write(uploaded_bytes)
        return load_from_excel(tmp)

    mysql_data = load_from_mysql()
    if mysql_data:
        return mysql_data

    local = _find_local_excel()
    if local:
        return load_from_excel(local)

    return SponsorIQData()


# ==============================================================================
# SECTION 4 — NORMALISATION & VALIDATION
# ==============================================================================

def normalise_industry(s: Any) -> Optional[str]:
    """Map free-text industry to canonical buckets — no invented mappings."""
    if pd.isna(s) or s is None:
        return None
    v = str(s).strip().lower()
    if not v:
        return None
    table = {
        "beauty & personal care": "Beauty & Personal Care",
        "beauty":                 "Beauty & Personal Care",
        "personal care":          "Beauty & Personal Care",
        "cosmetics":              "Beauty & Personal Care",
        "fmcg":                   "FMCG",
        "food":                   "FMCG",
        "beverage":               "FMCG",
        "banking & fintech":      "Banking & FinTech",
        "banking":                "Banking & FinTech",
        "fintech":                "Banking & FinTech",
        "finance":                "Banking & FinTech",
        "fashion & lifestyle":    "Fashion & Lifestyle",
        "fashion":                "Fashion & Lifestyle",
        "apparel":                "Fashion & Lifestyle",
        "automobile & ev":        "Automobile & EV",
        "automobile":             "Automobile & EV",
        "auto":                   "Automobile & EV",
        "consumer electronics":   "Consumer Electronics",
        "electronics":            "Consumer Electronics",
        "health & wellness":      "Health & Wellness",
        "healthcare":             "Health & Wellness",
        "pharma":                 "Health & Wellness",
        "home & kitchen":         "Home & Kitchen",
        "home":                   "Home & Kitchen",
        "real estate & infrastructure": "Real Estate & Infrastructure",
        "real estate":            "Real Estate & Infrastructure",
        "infrastructure":         "Real Estate & Infrastructure",
        "others":                 "Others",
    }
    for k, v in table.items():
        if k in v:
            return v
    return str(s).strip()


def tokens(s: Any) -> set:
    if pd.isna(s) or not s:
        return set()
    return {t for t in re.split(r"[^a-z0-9]+", str(s).lower()) if len(t) > 2}


def extract_audience_tags(s: Any) -> set:
    """Map free-text audience string to canonical audience tags."""
    if pd.isna(s) or not s:
        return set()
    low = str(s).lower()
    found = set()
    for tag, kws in AUDIENCE_TAGS.items():
        for kw in kws:
            if kw in low:
                found.add(tag)
                break
    return found


def validate_data(data: SponsorIQData) -> Dict[str, Any]:
    """Produce a data-quality report — no fixes, only observations."""
    report: Dict[str, Any] = {"tables": {}, "overall_quality_pct": 0.0}

    def table_quality(df: pd.DataFrame, critical_cols: List[str]) -> Dict[str, Any]:
        if df is None or df.empty:
            return {"rows": 0, "missing_critical": 0, "duplicates": 0, "quality_pct": 0.0}
        present = [c for c in critical_cols if c in df.columns]
        missing = sum(int(df[c].isna().sum()) for c in present)
        total = len(df) * max(len(present), 1)
        dup = int(df.duplicated().sum())
        quality = 0.0
        if total > 0:
            quality = max(0.0, 1.0 - (missing / total))
        return {"rows": len(df), "missing_critical": missing,
                "duplicates": dup, "quality_pct": round(quality * 100, 1)}

    report["tables"]["brands"] = table_quality(
        data.brands,
        ["Brand_ID", "Brand_Name", "Industry", "Target_Audience",
         "Sponsorship_Budget_Cr", "HQ_City"],
    )
    report["tables"]["events"] = table_quality(
        data.events,
        ["Event_ID", "Event_Name", "Event_Category", "Cities",
         "Sponsorship_Cost_Cr"],
    )
    report["tables"]["sponsorships"] = table_quality(
        data.sponsorships,
        ["Sponsorship_ID", "Brand_ID", "Event_ID", "Year",
         "Investment_Cr"],
    )
    report["tables"]["decision_makers"] = table_quality(
        data.decision_makers, ["Contact_ID", "Brand_ID", "Name"]
    )
    report["tables"]["crm"] = table_quality(
        data.crm, ["CRM_ID", "Brand_ID", "Deal_Stage"]
    )
    report["tables"]["campaigns"] = table_quality(
        data.campaigns, ["Campaign_ID", "Brand_ID", "Budget_Cr"]
    )

    qs = [v["quality_pct"] for v in report["tables"].values() if v["rows"] > 0]
    report["overall_quality_pct"] = round(sum(qs) / len(qs), 1) if qs else 0.0
    return report


# ==============================================================================
# SECTION 5 — DECK PARSER (PDF · PPTX · DOCX · XLSX · TXT)
# ==============================================================================

def _pdf_text(b: bytes) -> str:
    if not PDF_OK:
        raise RuntimeError("Install pypdf: pip install pypdf")
    r = PdfReader(io.BytesIO(b))
    return "\n".join((p.extract_text() or "") for p in r.pages)


def _pptx_text(b: bytes) -> str:
    if not PPTX_OK:
        raise RuntimeError("Install python-pptx: pip install python-pptx")
    prs = Presentation(io.BytesIO(b))
    out: List[str] = []
    for i, s in enumerate(prs.slides, 1):
        out.append(f"\n--- Slide {i} ---")
        for shp in s.shapes:
            try:
                if shp.has_text_frame and shp.text_frame.text.strip():
                    out.append(shp.text_frame.text)
                if getattr(shp, "has_table", False):
                    for row in shp.table.rows:
                        out.append(" | ".join(c.text for c in row.cells))
            except Exception:
                pass
        try:
            if s.has_notes_slide:
                nt = s.notes_slide.notes_text_frame.text
                if nt and nt.strip():
                    out.append("[Notes] " + nt)
        except Exception:
            pass
    return "\n".join(out)


def _docx_text(b: bytes) -> str:
    if not DOCX_OK:
        raise RuntimeError("Install python-docx: pip install python-docx")
    d = docx.Document(io.BytesIO(b))
    parts = [p.text for p in d.paragraphs if p.text.strip()]
    for t in d.tables:
        for row in t.rows:
            parts.append(" | ".join(c.text for c in row.cells))
    return "\n".join(parts)


def _xlsx_text(b: bytes) -> str:
    xls = pd.ExcelFile(io.BytesIO(b))
    out: List[str] = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        out.append(f"\n--- Sheet: {sheet} ---")
        out.append(df.to_string(index=False))
    return "\n".join(out)


def parse_deck(filename: str, content: bytes) -> Tuple[str, Optional[str]]:
    """Return (text, error_message)."""
    name = filename.lower()
    try:
        if name.endswith(".pdf"):
            txt = _pdf_text(content)
        elif name.endswith((".pptx", ".ppt")):
            txt = _pptx_text(content)
        elif name.endswith(".docx"):
            txt = _docx_text(content)
        elif name.endswith((".xlsx", ".xls")):
            txt = _xlsx_text(content)
        else:
            txt = content.decode("utf-8", errors="ignore")
    except Exception as e:
        return "", f"Could not parse deck: {e}"

    txt = re.sub(r"[ \t]+", " ", txt)
    txt = re.sub(r"\n{3,}", "\n\n", txt)
    return txt.strip(), None


# ==============================================================================
# SECTION 6 — DECK FIELD EXTRACTION (deterministic · explainable · no guessing)
# ==============================================================================

@dataclass
class DeckBrief:
    """Structured brief extracted from a deck (or from a selected event)."""
    event_name: Optional[str] = None
    event_category: Optional[str] = None
    artist_or_property: Optional[str] = None
    cities: List[str] = field(default_factory=list)
    number_of_cities: Optional[int] = None
    audience_text: Optional[str] = None
    audience_tags: List[str] = field(default_factory=list)
    reach_text: Optional[str] = None
    footfall_text: Optional[str] = None
    duration_text: Optional[str] = None
    sponsorship_roles: List[str] = field(default_factory=list)
    sponsorship_value_min_cr: Optional[float] = None
    sponsorship_value_max_cr: Optional[float] = None
    category_exclusivity: Optional[str] = None
    deliverables: List[str] = field(default_factory=list)
    raw_char_count: int = 0
    extraction_confidence: str = "Low"
    missing_fields: List[str] = field(default_factory=list)
    source: str = SRC_DECK

    def to_dict(self) -> Dict[str, Any]:
        return {k: (v if not isinstance(v, list) else list(v))
                for k, v in self.__dict__.items()}


BUDGET_RX_1 = re.compile(
    r"(?:₹|rs\.?|inr|rupees)\s*([\d,]+(?:\.\d+)?)\s*(crore|cr\b|lakh|lakhs|million|mn\b)",
    re.IGNORECASE)
BUDGET_RX_2 = re.compile(
    r"([\d,]+(?:\.\d+)?)\s*(crore|cr\b|lakh|lakhs|million|mn\b)", re.IGNORECASE)


def _to_cr(v: float, unit: str) -> float:
    u = unit.lower()
    if u.startswith("lakh"):  return v / 100.0
    if u.startswith("million") or u == "mn": return v / 10.0
    return v


def extract_budget_range(text: str) -> Tuple[Optional[float], Optional[float]]:
    vals: List[float] = []
    for m in BUDGET_RX_1.finditer(text):
        vals.append(_to_cr(float(m.group(1).replace(",", "")), m.group(2)))
    if not vals:
        for m in BUDGET_RX_2.finditer(text):
            v = _to_cr(float(m.group(1).replace(",", "")), m.group(2))
            if 0.1 <= v <= 5000:
                vals.append(v)
    vals = [v for v in vals if 0.05 <= v <= 20000]
    return (min(vals), max(vals)) if vals else (None, None)


def extract_cities(text: str, known_cities: List[str]) -> List[str]:
    """Return cities that appear literally in the text AND exist in our known set."""
    low = text.lower()
    hits = []
    for c in known_cities:
        if not c or len(c) < 4:
            continue
        if c.lower() in low:
            hits.append(c)
    return sorted(set(hits))


def extract_number_before(text: str, keyword: str) -> Optional[int]:
    """Look for '<number> <keyword>' style mentions."""
    pattern = re.compile(rf"(\d[\d,]*)\s*(?:\+\s*)?{re.escape(keyword)}", re.IGNORECASE)
    m = pattern.search(text)
    if not m:
        return None
    try:
        return int(m.group(1).replace(",", ""))
    except Exception:
        return None


def extract_deck_brief(deck_text: str, brands_df: pd.DataFrame,
                       events_df: pd.DataFrame) -> DeckBrief:
    """
    Deterministic extraction. Only pulls what is LITERALLY in the deck.
    Anything not found stays None / [] — never guessed.
    """
    brief = DeckBrief(raw_char_count=len(deck_text), source=SRC_DECK)
    low = deck_text.lower()

    # ---- Event name ----
    if events_df is not None and not events_df.empty and "Event_Name" in events_df.columns:
        for ev in events_df["Event_Name"].dropna().astype(str):
            if ev and ev.lower() in low:
                brief.event_name = ev
                break

    if brief.event_name is None:
        for line in deck_text.splitlines():
            s = line.strip(" -·•\t")
            if 5 <= len(s) <= 90 and not s.lower().startswith(("http", "www")):
                if sum(c.isalpha() for c in s) >= 5:
                    brief.event_name = s
                    break

    # ---- Event category ----
    category_keywords = {
        "Sports - Cricket":   ["cricket", "ipl", "bcci"],
        "Sports - Football":  ["football", "isl", "soccer"],
        "Sports - Kabaddi":   ["kabaddi", "pro kabaddi"],
        "Sports - Marathon":  ["marathon", "half marathon", "10k", "25k"],
        "Sports - Motorsport":["motogp", "formula", "racing", "motorsport"],
        "Music - EDM Festival": ["edm", "sunburn", "electronic music"],
        "Music - Multi-genre Festival": ["music festival", "live music", "concert",
                                          "lollapalooza", "nh7"],
        "Fashion - Runway Show": ["fashion week", "runway", "couture", "fashion tour"],
        "Entertainment - Film Awards": ["film award", "iifa", "filmfare"],
        "Entertainment - Film Festival": ["film festival", "iffi", "mami"],
        "Entertainment - Pop Culture Convention": ["comic con", "pop culture"],
        "Entertainment - Comedy Festival": ["comedy", "stand-up", "standup"],
        "Cultural Festival":  ["cultural", "heritage", "arts festival", "crafts mela"],
        "Religious - Pilgrimage": ["kumbh", "pilgrimage", "yatra", "rath yatra"],
        "Business - Conference": ["conference", "summit", "conclave"],
        "Trade Show - Automobile": ["auto expo", "automobile trade"],
        "Trade Show - Telecom & Tech": ["mobile congress", "telecom"],
        "Culture - Art Fair": ["art fair", "art festival"],
        "Design - Trade Exhibition": ["design week", "design exhibition"],
        "Food - Culinary Festival": ["food festival", "culinary"],
        "Food & Wine - Music & Wine Festival": ["wine festival", "sula fest"],
        "Retail - E-commerce Sale Event": ["sale event", "e-commerce sale", "big billion"],
        "Retail - Fashion E-commerce Sale": ["fashion sale", "end of reason"],
        "Sports - Multi-discipline": ["khelo india", "national games", "multi-sport"],
        "Sports - Athletics": ["athletics", "federation cup"],
        "Sports - Hockey": ["hockey", "hockey india"],
        "Sports - Badminton": ["badminton", "pbl"],
        "Sports - Table Tennis": ["table tennis", "utt"],
        "Sports - Volleyball": ["volleyball", "pro volleyball"],
    }
    for cat, kws in category_keywords.items():
        if any(k in low for k in kws):
            brief.event_category = cat
            break

    # ---- Artist / property name ----
    m = re.search(r"(?:feat\.?|ft\.?|featuring|live by|performed by)\s+([A-Z][A-Za-z .'\-]{2,40})", deck_text)
    if m:
        brief.artist_or_property = m.group(1).strip()

    # ---- Cities ----
    if brands_df is not None and not brands_df.empty and "HQ_City" in brands_df.columns:
        known_cities = sorted({str(c).strip() for c in brands_df["HQ_City"].dropna() if str(c).strip()})
    else:
        known_cities = []
    if events_df is not None and not events_df.empty and "Cities" in events_df.columns:
        for cell in events_df["Cities"].dropna().astype(str):
            known_cities.extend([c.strip() for c in re.split(r"[,;|/]", cell) if c.strip()])
    known_cities = sorted(set(known_cities))
    brief.cities = extract_cities(deck_text, known_cities)

    n_cities = extract_number_before(deck_text, "cities") or extract_number_before(deck_text, "city")
    brief.number_of_cities = n_cities or (len(brief.cities) or None)

    # ---- Audience ----
    aud_phrases = [
        "family", "families", "youth", "students", "millennials", "urban",
        "rural", "professionals", "hni", "premium", "mass", "b2b",
        "fitness", "runners", "foodies", "music fans", "sports fans",
        "college", "campus", "gen z", "young adults",
    ]
    found_aud = [p for p in aud_phrases if p in low]
    if found_aud:
        brief.audience_text = ", ".join(sorted(set(found_aud)))
        brief.audience_tags = sorted(extract_audience_tags(brief.audience_text))

    # ---- Reach ----
    reach_m = re.search(r"(\d[\d,\.]*\s*(?:lakh|crore|million|mn|k|cr)\b[^.,;\n]{0,40})",
                        low, re.IGNORECASE)
    if reach_m:
        brief.reach_text = reach_m.group(1).strip()
    rk_m = re.search(r"reach[:\s]*([^\n]{3,80})", deck_text, re.IGNORECASE)
    if rk_m and not brief.reach_text:
        brief.reach_text = rk_m.group(1).strip()

    # ---- Footfall ----
    ff_m = re.search(r"footfall[:\s]*([^\n]{3,80})", deck_text, re.IGNORECASE)
    if ff_m:
        brief.footfall_text = ff_m.group(1).strip()

    # ---- Duration ----
    dur_m = re.search(r"(?:duration|tour duration|period)[:\s]*([^\n]{2,60})",
                      deck_text, re.IGNORECASE)
    if dur_m:
        brief.duration_text = dur_m.group(1).strip()

    # ---- Sponsorship roles ----
    role_terms = [
        "Title Sponsor", "Presenting Sponsor", "Co-Sponsor", "Powered By",
        "Associate Sponsor", "Official Partner", "Category Partner",
        "Media Partner", "Merchandise Partner",
    ]
    brief.sponsorship_roles = [r for r in role_terms if r.lower() in low]

    # ---- Budget ----
    bmin, bmax = extract_budget_range(deck_text)
    brief.sponsorship_value_min_cr = bmin
    brief.sponsorship_value_max_cr = bmax

    # ---- Category exclusivity ----
    if "exclusivity" in low or "exclusive" in low:
        brief.category_exclusivity = "Category exclusivity mentioned"
    if "exclusive category" in low:
        brief.category_exclusivity = "Exclusive category mentioned"

    # ---- Deliverables ----
    deliverable_kw = ["logo", "branding", "activation", "tv", "digital",
                       "social media", "on-ground", "hospitality", "tickets",
                       "naming rights", "led", "hoarding"]
    for kw in deliverable_kw:
        if kw in low:
            brief.deliverables.append(kw.title())

    # ---- Missing fields ----
    if brief.event_name is None:           brief.missing_fields.append("event_name")
    if brief.event_category is None:       brief.missing_fields.append("event_category")
    if not brief.cities:                   brief.missing_fields.append("cities")
    if not brief.audience_tags:            brief.missing_fields.append("audience")
    if brief.sponsorship_value_max_cr is None: brief.missing_fields.append("budget")
    if not brief.sponsorship_roles:        brief.missing_fields.append("sponsorship_roles")

    # ---- Extraction confidence ----
    if len(brief.missing_fields) <= 1:
        brief.extraction_confidence = "High"
    elif len(brief.missing_fields) <= 3:
        brief.extraction_confidence = "Medium"
    else:
        brief.extraction_confidence = "Low"

    return brief


def brief_from_event_row(row: pd.Series) -> DeckBrief:
    """Build a DeckBrief from a row in the Events sheet."""
    b = DeckBrief(source=SRC_DB)
    b.event_name = str(row.get("Event_Name")) if pd.notna(row.get("Event_Name")) else None
    b.event_category = str(row.get("Event_Category")) if pd.notna(row.get("Event_Category")) else None
    cities_cell = row.get("Cities")
    if pd.notna(cities_cell):
        b.cities = [c.strip() for c in re.split(r"[,;|/]", str(cities_cell)) if c.strip()]
        b.number_of_cities = len(b.cities)
    aud = row.get("Audience_Type")
    if pd.notna(aud):
        b.audience_text = str(aud)
        b.audience_tags = sorted(extract_audience_tags(str(aud)))
    reach = row.get("Reach_Million")
    if pd.notna(reach):
        b.reach_text = f"{reach} Million"
    ff = row.get("Annual_Footfall")
    if pd.notna(ff):
        b.footfall_text = str(ff)
    sc = row.get("Sponsorship_Cost_Cr")
    if pd.notna(sc):
        b.sponsorship_value_min_cr = float(sc)
        b.sponsorship_value_max_cr = float(sc)
    return b


# ==============================================================================
# SECTION 7 — RECOMMENDATION ENGINE
# ==============================================================================

@dataclass
class Requirement:
    sponsorship_role: Optional[str] = None
    target_budget_cr: Optional[float] = None
    target_cities: List[str] = field(default_factory=list)
    target_audience: Optional[str] = None
    target_industry: Optional[str] = None
    exclusivity_required: bool = False
    user_override: bool = False


@dataclass
class MatchResult:
    brand_id: str
    brand_name: str
    industry: Optional[str]
    match_score: Optional[float]
    confidence: str
    data_completeness: float
    audience_fit: Optional[float]
    budget_fit: Optional[float]
    industry_fit: Optional[float]
    geography_fit: Optional[float]
    brand_value_fit: Optional[float]
    competition_fit: Optional[float]
    historical_fit: Optional[float]
    recommended_role: Optional[str]
    category: str
    reasons: List[str]
    risks: List[str]
    missing_data: List[str]
    evidence: Dict[str, Any]
    hard_filter_failures: List[str]


def _safe(v: Any) -> Optional[float]:
    try:
        if v is None or pd.isna(v):
            return None
        return float(v)
    except Exception:
        return None


def apply_hard_filters(brand: pd.Series, req: Requirement,
                       thresholds: Dict[str, float]) -> List[str]:
    failures: List[str] = []
    if req.target_budget_cr is not None:
        b = _safe(brand.get("Sponsorship_Budget_Cr"))
        if b is not None:
            if b < req.target_budget_cr * thresholds["budget_min_ratio"]:
                failures.append("Budget Mismatch")
    return failures


def _audience_score(brand_aud: Any, brief: DeckBrief,
                    req: Requirement) -> Optional[float]:
    brand_tags = extract_audience_tags(brand_aud)
    if not brand_tags:
        return None
    event_tags = set(brief.audience_tags)
    if req.target_audience:
        event_tags |= extract_audience_tags(req.target_audience)
    if not event_tags:
        return None
    overlap = len(brand_tags & event_tags)
    denom = max(len(event_tags), 1)
    return min(100.0, 100.0 * overlap / denom)


def _industry_score(brand_industry: Any, brief: DeckBrief,
                    req: Requirement) -> Optional[float]:
    if pd.isna(brand_industry) or not brand_industry:
        return None
    canon = normalise_industry(brand_industry) or str(brand_industry)
    event_cat = brief.event_category or req.target_industry
    if not event_cat:
        return None
    affinity = INDUSTRY_EVENT_AFFINITY.get(canon, {}).get(event_cat)
    if affinity is not None:
        return round(affinity * 100.0, 1)
    b_tokens = tokens(canon) | tokens(brand_industry)
    e_tokens = tokens(event_cat) | tokens(brief.event_name or "")
    if not b_tokens or not e_tokens:
        return None
    overlap = len(b_tokens & e_tokens)
    if overlap == 0:
        return 25.0
    return min(100.0, 25.0 + overlap * 25.0)


def _budget_score(brand_budget: Any, req: Requirement) -> Optional[float]:
    b = _safe(brand_budget)
    if b is None or req.target_budget_cr is None or req.target_budget_cr <= 0:
        return None
    ratio = b / req.target_budget_cr
    if ratio >= 1.0:
        if ratio <= 3.0:
            return 100.0
        return max(60.0, 100.0 - (ratio - 3.0) * 10)
    if ratio >= 0.9:  return 92.0
    if ratio >= 0.75: return 82.0
    if ratio >= 0.6:  return 70.0
    if ratio >= 0.5:  return 60.0
    if ratio >= 0.4:  return 50.0
    return max(10.0, ratio * 100.0)


def _geography_score(brand: pd.Series, brief: DeckBrief,
                     req: Requirement) -> Optional[float]:
    event_cities = set(c.strip().lower() for c in (brief.cities or []))
    event_cities |= set(c.strip().lower() for c in (req.target_cities or []))
    if not event_cities:
        return None
    hq_city = str(brand.get("HQ_City") or "").strip().lower()
    hq_state = str(brand.get("HQ_State") or "").strip().lower()
    if not hq_city and not hq_state:
        return None
    if hq_city and hq_city in event_cities:
        return 100.0
    if hq_state and hq_state in event_cities:
        return 85.0
    return 35.0


def _brand_value_score(brand: pd.Series, brief: DeckBrief) -> Optional[float]:
    b_type = str(brand.get("Brand_Type") or "").strip().upper()
    target = str(brand.get("Target_Audience") or "").lower()
    if not b_type and not target:
        return None
    event_aud = " ".join(brief.audience_tags).lower() or (brief.audience_text or "").lower()
    if not event_aud:
        return None
    score = 50.0
    if "b2b" in b_type.lower() and "b2b" in event_aud:
        score += 35
    if "b2c" in b_type.lower() and ("mass" in event_aud or "premium" in event_aud
                                     or "youth" in event_aud or "families" in event_aud):
        score += 20
    if "premium" in target and "premium" in event_aud:
        score += 25
    if "mass" in target and "mass" in event_aud:
        score += 20
    return min(100.0, score)


def _competition_score(brand_id: str, brief: DeckBrief,
                       sponsorships: pd.DataFrame,
                       events: pd.DataFrame) -> Tuple[Optional[float], List[str]]:
    if sponsorships is None or sponsorships.empty:
        return None, []
    risks: List[str] = []
    try:
        brand_rows = sponsorships[sponsorships["Brand_ID"].astype(str) == str(brand_id)]
    except Exception:
        return None, []
    if brand_rows.empty:
        return None, []
    if events is not None and not events.empty and "Event_ID" in events.columns:
        ev_lookup = events.set_index("Event_ID")["Event_Category"].to_dict()
        same_cat = 0
        for _, r in brand_rows.iterrows():
            ev_id = r.get("Event_ID")
            cat = ev_lookup.get(ev_id)
            if cat and brief.event_category and cat == brief.event_category:
                same_cat += 1
        if same_cat >= 2:
            risks.append(f"Brand has {same_cat} prior sponsorships in the same event category")
            return 55.0, risks
        if same_cat == 1:
            return 75.0, risks
    return 80.0, risks


def _historical_score(brand_id: str, brief: DeckBrief,
                      sponsorships: pd.DataFrame,
                      events: pd.DataFrame,
                      recency: Dict[str, float]) -> Tuple[Optional[float], Dict[str, Any]]:
    evid: Dict[str, Any] = {
        "previous_sponsorships": 0,
        "relevant_category_sponsorships": 0,
        "avg_investment_cr": None,
        "latest_year": None,
        "cities_sponsored": [],
        "roles": [],
    }
    if sponsorships is None or sponsorships.empty:
        return None, evid
    try:
        rows = sponsorships[sponsorships["Brand_ID"].astype(str) == str(brand_id)]
    except Exception:
        return None, evid
    if rows.empty:
        return None, evid

    evid["previous_sponsorships"] = len(rows)
    evid["avg_investment_cr"] = round(float(rows["Investment_Cr"].mean()), 2) \
        if "Investment_Cr" in rows.columns else None
    if "Year" in rows.columns:
        try:
            evid["latest_year"] = int(rows["Year"].max())
        except Exception:
            pass
    if "Sponsorship_Type" in rows.columns:
        evid["roles"] = sorted(set(rows["Sponsorship_Type"].dropna().astype(str)))[:6]

    same_cat_count = 0
    if events is not None and not events.empty and "Event_ID" in events.columns:
        cat_map = events.set_index("Event_ID")["Event_Category"].to_dict()
        for _, r in rows.iterrows():
            cat = cat_map.get(r.get("Event_ID"))
            if cat and brief.event_category and cat == brief.event_category:
                same_cat_count += 1
    evid["relevant_category_sponsorships"] = same_cat_count

    if "Year" not in rows.columns:
        return 55.0, evid

    now_year = datetime.utcnow().year
    weighted = 0.0
    total_w = 0.0
    for _, r in rows.iterrows():
        try:
            y = int(r["Year"])
        except Exception:
            continue
        months = max(0, (now_year - y) * 12)
        if months <= 12:  w = recency["0-12_months"]
        elif months <= 24: w = recency["12-24_months"]
        elif months <= 36: w = recency["24-36_months"]
        else:              w = recency["36+_months"]
        weighted += w
        total_w += 1
    if total_w == 0:
        return None, evid

    base = min(1.0, weighted / max(1.0, total_w * 1.0))
    cat_boost = min(1.0, same_cat_count / 3.0) * 0.3
    score = min(100.0, (base * 0.7 + cat_boost) * 100.0)
    return round(score, 1), evid


def recommend_brands(
    brief: DeckBrief,
    req: Requirement,
    brands_df: pd.DataFrame,
    sponsorships_df: pd.DataFrame,
    events_df: pd.DataFrame,
    weights: Dict[str, float],
    thresholds: Dict[str, float],
    recency: Dict[str, float],
) -> List[MatchResult]:
    if brands_df is None or brands_df.empty:
        return []

    results: List[MatchResult] = []

    for _, brand in brands_df.iterrows():
        brand_id = str(brand.get("Brand_ID") or "")
        brand_name = str(brand.get("Brand_Name") or "")
        if not brand_id or not brand_name:
            continue

        hf = apply_hard_filters(brand, req, thresholds)

        aud   = _audience_score(brand.get("Target_Audience"), brief, req)
        bud   = _budget_score(brand.get("Sponsorship_Budget_Cr"), req)
        ind   = _industry_score(brand.get("Industry"), brief, req)
        geo   = _geography_score(brand, brief, req)
        bval  = _brand_value_score(brand, brief)
        comp, comp_risks = _competition_score(brand_id, brief,
                                              sponsorships_df, events_df)
        hist, hist_evid = _historical_score(brand_id, brief,
                                             sponsorships_df, events_df, recency)

        components = {
            "audience":    aud,
            "budget":      bud,
            "industry":    ind,
            "geography":   geo,
            "brand_value": bval,
            "competition": comp,
        }
        available = {k: v for k, v in components.items() if v is not None}
        completeness = len(available) / len(components)

        if len(available) < thresholds["min_dimensions_scored"]:
            match_score = None
        else:
            total_w = sum(weights[k] for k in available)
            if total_w <= 0:
                match_score = None
            else:
                weighted = sum(available[k] * weights[k] for k in available)
                match_score = round(weighted / total_w, 1)

        if completeness >= 0.85 and len(available) >= 5:
            confidence = "High"
        elif completeness >= 0.65 and len(available) >= 4:
            confidence = "Medium"
        else:
            confidence = "Low"

        missing = [k for k, v in components.items() if v is None]

        if hf:
            category = "Not Recommended"
        elif match_score is None:
            category = "Not Recommended"
        elif match_score >= thresholds["min_match_for_recommend"] and \
             completeness >= thresholds["min_data_completeness"]:
            category = "Recommended"
        elif match_score >= thresholds["min_match_for_potential"]:
            category = "Potential Fit"
        else:
            category = "Not Recommended"

        role = None
        if match_score is not None and bud is not None:
            if match_score >= 80 and bud >= 80:
                role = "Title Sponsor"
            elif match_score >= 72:
                role = "Co-Sponsor"
            elif match_score >= 65:
                role = "Associate Sponsor"
            else:
                role = "Official Partner"

        reasons: List[str] = []
        if ind is not None and ind >= 60:
            reasons.append(f"Industry fit is strong ({ind:.0f}/100) for a "
                           f"{brief.event_category or 'this event category'} property.")
        if aud is not None and aud >= 60:
            reasons.append(f"Audience overlap is {aud:.0f}/100 — brand speaks to "
                           f"{brand.get('Target_Audience', 'the same audience')}.")
        if bud is not None and bud >= 70 and req.target_budget_cr:
            reasons.append(
                f"Sponsorship budget of ₹{_safe(brand.get('Sponsorship_Budget_Cr')):.1f} Cr "
                f"comfortably covers the ₹{req.target_budget_cr:.1f} Cr ask.")
        if geo is not None and geo >= 80:
            reasons.append(f"Headquartered in {brand.get('HQ_City')}, "
                           f"inside the event's footprint.")
        if hist is not None and hist >= 60 and hist_evid["previous_sponsorships"]:
            reasons.append(
                f"{hist_evid['previous_sponsorships']} prior sponsorships on record "
                f"({hist_evid['relevant_category_sponsorships']} in the same category).")
        if not reasons:
            reasons.append("No single dimension stood out — see component scores.")

        risks: List[str] = list(comp_risks)
        if bud is not None and bud < 50:
            risks.append("Budget fit is weak versus the required sponsorship value.")
        if aud is None:
            risks.append("Audience compatibility could not be evaluated "
                         "(brand target audience missing in dataset).")
        if bud is None:
            risks.append("Budget compatibility could not be evaluated "
                         "(brand sponsorship budget missing).")
        if geo is None:
            risks.append("Geography compatibility could not be evaluated.")

        evidence = {
            "sponsorship_budget_cr": _safe(brand.get("Sponsorship_Budget_Cr")),
            "annual_revenue_cr":     _safe(brand.get("Annual_Revenue_Cr")),
            "marketing_spend_cr":    _safe(brand.get("Marketing_Spend_Cr")),
            "industry":              brand.get("Industry"),
            "sub_category":          brand.get("Sub_Category"),
            "target_audience":       brand.get("Target_Audience"),
            "hq_city":               brand.get("HQ_City"),
            "hq_state":              brand.get("HQ_State"),
            "brand_type":            brand.get("Brand_Type"),
            "required_budget_cr":    req.target_budget_cr,
            "event_cities":          brief.cities,
            "event_category":        brief.event_category,
            "historical":            hist_evid,
        }

        results.append(MatchResult(
            brand_id=brand_id,
            brand_name=brand_name,
            industry=str(brand.get("Industry") or "") or None,
            match_score=match_score,
            confidence=confidence,
            data_completeness=round(completeness * 100, 1),
            audience_fit=aud,
            budget_fit=bud,
            industry_fit=ind,
            geography_fit=geo,
            brand_value_fit=bval,
            competition_fit=comp,
            historical_fit=hist,
            recommended_role=role,
            category=category,
            reasons=reasons,
            risks=risks,
            missing_data=missing,
            evidence=evidence,
            hard_filter_failures=hf,
        ))

    order = {"Recommended": 0, "Potential Fit": 1, "Not Recommended": 2}
    results.sort(key=lambda r: (
        order.get(r.category, 3),
        -(r.match_score if r.match_score is not None else -1)
    ))
    return results


# ==============================================================================
# SECTION 7.5 — AI LAYER (optional · user-supplied API key · no silent fallback)
# ==============================================================================

def _parse_json_blob(raw: str) -> Optional[dict]:
    """Robustly extract a JSON object from an LLM response."""
    if not raw:
        return None
    raw = raw.strip()
    try:
        return json.loads(raw)
    except Exception:
        pass
    clean = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw,
                   flags=re.MULTILINE).strip()
    try:
        return json.loads(clean)
    except Exception:
        pass
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if m:
        try:
            return json.loads(m.group(0))
        except Exception:
            pass
    return None


def _ai_chat_json(
    system: str,
    user: str,
    api_key: str,
    provider_name: str,
    base_url: str,
    model: str,
    max_tokens: int = 4096,
) -> Tuple[Optional[dict], Optional[str]]:
    """
    Single entry point for every LLM call. Returns (parsed_json, error).
    Includes timeout, exception-class logging, and strict base_url handling.
    Automatically picks max_tokens vs max_completion_tokens based on the model.
    """
    if not api_key or not api_key.strip():
        return None, "API key is empty."
    if not model:
        return None, "Model name is empty."

    provider = AI_PROVIDERS.get(provider_name, {})
    kind = provider.get("kind", "openai")

    debug = (f"[provider={provider_name} | model={model} | "
             f"base_url={base_url or '(default)'} | sdk={OPENAI_VERSION}]")

    try:
        if kind == "anthropic":
            if not ANTHROPIC_OK:
                return None, (f"Anthropic SDK not installed. {debug} — "
                              f"Run: pip install anthropic")
            client = anthropic.Anthropic(api_key=api_key.strip(), timeout=60.0)
            resp = client.messages.create(
                model=model,
                max_tokens=max_tokens,
                system=system,
                messages=[{"role": "user", "content": user}],
            )
            raw = "".join(getattr(b, "text", "") for b in resp.content)
        else:
            if not OPENAI_OK:
                return None, (f"OpenAI SDK not installed. {debug} — "
                              f"Run: pip install openai")
            effective_base = base_url.strip() if base_url else None
            client = OpenAI(api_key=api_key.strip(),
                            base_url=effective_base, timeout=60.0)

            create_kwargs: Dict[str, Any] = dict(
                model=model,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                temperature=0.2,
            )

            # Newer reasoning models (Gemini 3.x, o1/o3/o4) reject `max_tokens`
            # and require `max_completion_tokens` instead.
            model_lower = model.lower()
            needs_completion_tokens = (
                "gemini-3" in model_lower
                or model_lower.startswith(("o1", "o3", "o4"))
            )
            if needs_completion_tokens:
                create_kwargs["max_completion_tokens"] = max_tokens
            else:
                create_kwargs["max_tokens"] = max_tokens

            try:
                resp = client.chat.completions.create(**create_kwargs)
            except Exception as inner_e:
                # If the model rejected max_tokens but we tried it, retry with
                # max_completion_tokens (and vice-versa) once.
                msg = str(inner_e).lower()
                if "max_tokens" in msg and "max_completion_tokens" in msg:
                    create_kwargs.pop("max_tokens", None)
                    create_kwargs["max_completion_tokens"] = max_tokens
                    resp = client.chat.completions.create(**create_kwargs)
                elif "max_completion_tokens" in msg and "max_tokens" in msg:
                    create_kwargs.pop("max_completion_tokens", None)
                    create_kwargs["max_tokens"] = max_tokens
                    resp = client.chat.completions.create(**create_kwargs)
                else:
                    raise

            raw = resp.choices[0].message.content or ""
    except Exception as e:
        return None, f"[{type(e).__name__}] {e}  {debug}"

    data = _parse_json_blob(raw)
    if data is None:
        preview = (raw or "")[:200].replace("\n", " ")
        return None, (f"Model did not return valid JSON. {debug} — "
                      f"Response preview: {preview}")
    return data, None


def ai_test_connection(api_key: str, provider: str,
                       base_url: str, model: str) -> Tuple[bool, str]:
    """Lightweight ping so the user knows their key works before running."""
    data, err = _ai_chat_json(
        system="You reply with JSON only.",
        user='Reply with exactly: {"ok": true}',
        api_key=api_key,
        provider_name=provider,
        base_url=base_url,
        model=model,
        max_tokens=50,
    )
    if err:
        return False, err
    if isinstance(data, dict) and data.get("ok") is True:
        return True, f"Connection OK · model={model}"
    return True, f"Response received (model={model})."


def ai_extract_deck_brief(
    deck_text: str,
    api_key: str,
    provider: str,
    base_url: str,
    model: str,
) -> Tuple[Optional[DeckBrief], Optional[str]]:
    """AI-powered brief extraction. Returns DeckBrief or (None, error)."""
    text = deck_text[:30000]

    user_prompt = f"""Extract a structured sponsorship brief from the deck below.

Return STRICT JSON with EXACTLY these keys:
{{
  "event_name": string or null,
  "event_category": one of {AI_CATEGORY_LIST} or null,
  "artist_or_property": string or null,
  "cities": [strings],
  "number_of_cities": integer or null,
  "audience_text": string or null,
  "audience_tags": subset of {list(AUDIENCE_TAGS.keys())},
  "reach_text": string or null,
  "footfall_text": string or null,
  "duration_text": string or null,
  "sponsorship_roles": subset of {AI_ROLE_LIST},
  "sponsorship_value_min_cr": number or null,
  "sponsorship_value_max_cr": number or null,
  "category_exclusivity": string or null,
  "deliverables": [strings],
  "extraction_confidence": "High" or "Medium" or "Low",
  "missing_fields": [strings]
}}

RULES
- Only use values LITERALLY present in the deck. No guessing.
- Budget values must be in ₹ Crore (convert lakh → /100, million → /10).
- If a field is not present, use null (or [] for arrays).

DECK TEXT:
\"\"\"
{text}
\"\"\"
"""

    data, err = _ai_chat_json(
        AI_EXTRACTION_SYSTEM, user_prompt,
        api_key, provider, base_url, model,
    )
    if err:
        return None, err
    if not isinstance(data, dict):
        return None, "Unexpected AI response format."

    b = DeckBrief(source=SRC_AI_EXTR)
    b.event_name = data.get("event_name") or None
    b.event_category = data.get("event_category") or None
    b.artist_or_property = data.get("artist_or_property") or None
    b.cities = [str(c).strip() for c in (data.get("cities") or []) if str(c).strip()]
    b.number_of_cities = data.get("number_of_cities") or (len(b.cities) or None)
    b.audience_text = data.get("audience_text") or None
    b.audience_tags = [t for t in (data.get("audience_tags") or [])
                       if t in AUDIENCE_TAGS]
    b.reach_text = data.get("reach_text") or None
    b.footfall_text = data.get("footfall_text") or None
    b.duration_text = data.get("duration_text") or None
    b.sponsorship_roles = [r for r in (data.get("sponsorship_roles") or [])
                           if r in AI_ROLE_LIST]
    b.sponsorship_value_min_cr = _safe(data.get("sponsorship_value_min_cr"))
    b.sponsorship_value_max_cr = _safe(data.get("sponsorship_value_max_cr"))
    b.category_exclusivity = data.get("category_exclusivity") or None
    b.deliverables = [str(x) for x in (data.get("deliverables") or [])]
    b.raw_char_count = len(deck_text)
    b.extraction_confidence = data.get("extraction_confidence") or "Medium"
    b.missing_fields = [str(x) for x in (data.get("missing_fields") or [])]
    return b, None


def ai_recommend_brands(
    brief: DeckBrief,
    req: Requirement,
    api_key: str,
    provider: str,
    base_url: str,
    model: str,
    n_recommendations: int = 15,
) -> Tuple[List[MatchResult], Optional[str]]:
    """AI-powered brand recommendations. Returns (list, error)."""
    brief_json = json.dumps(brief.to_dict(), indent=2, default=str)

    user_prompt = f"""Recommend real, existing brands that would plausibly sponsor
this event, ranked by fit.

EVENT BRIEF:
{brief_json}

COMMERCIAL REQUIREMENT:
- Sponsorship role sought: {req.sponsorship_role or "Any"}
- Required sponsorship value (₹ Cr): {req.target_budget_cr if req.target_budget_cr is not None else "Not specified"}
- Preferred industry: {req.target_industry or "Any"}
- Mandatory cities: {", ".join(req.target_cities) if req.target_cities else "Not specified"}
- Target audience: {req.target_audience or brief.audience_text or "Not specified"}
- Category exclusivity required: {req.exclusivity_required}

Return STRICT JSON in exactly this shape:
{{
  "recommendations": [
    {{
      "brand_name": "...",
      "industry": "...",
      "match_score": 0-100,
      "confidence": "High" | "Medium" | "Low",
      "data_completeness": 0-100,
      "audience_fit": 0-100 or null,
      "budget_fit": 0-100 or null,
      "industry_fit": 0-100 or null,
      "geography_fit": 0-100 or null,
      "brand_value_fit": 0-100 or null,
      "competition_fit": 0-100 or null,
      "historical_fit": 0-100 or null,
      "recommended_role": "Title Sponsor" | "Co-Sponsor" | "Associate Sponsor" | "Official Partner",
      "category": "Recommended" | "Potential Fit" | "Not Recommended",
      "reasons": ["short factual reason", "..."],
      "risks": ["short risk", "..."],
      "missing_data": ["what you could not evaluate"],
      "evidence": {{
         "notes": "brief reasoning",
         "estimated_budget_cr": number or null,
         "hq_city": "..." or null
      }}
    }}
  ]
}}

RULES
- Use ONLY real brands (e.g. Tata, Amul, Zomato, boAt, Titan, ICICI, Red Bull, etc.).
- Do NOT invent brand names. If unsure, omit the brand entirely.
- Every numeric score must have a matching short reason.
- Aim for around {n_recommendations} recommendations, sorted best first.
- If nothing fits well, return fewer items and be honest.
- If you are uncertain, set "confidence" to "Low".
"""

    data, err = _ai_chat_json(
        AI_RECOMMENDATION_SYSTEM, user_prompt,
        api_key, provider, base_url, model,
        max_tokens=6000,
    )
    if err:
        return [], err

    recs = (data or {}).get("recommendations") or []
    results: List[MatchResult] = []
    for i, r in enumerate(recs):
        try:
            brand_name = str(r.get("brand_name") or "").strip()
            if not brand_name:
                continue
            results.append(MatchResult(
                brand_id=f"AI-{i:03d}",
                brand_name=brand_name,
                industry=r.get("industry"),
                match_score=_safe(r.get("match_score")),
                confidence=r.get("confidence") or "Low",
                data_completeness=_safe(r.get("data_completeness")) or 0.0,
                audience_fit=_safe(r.get("audience_fit")),
                budget_fit=_safe(r.get("budget_fit")),
                industry_fit=_safe(r.get("industry_fit")),
                geography_fit=_safe(r.get("geography_fit")),
                brand_value_fit=_safe(r.get("brand_value_fit")),
                competition_fit=_safe(r.get("competition_fit")),
                historical_fit=_safe(r.get("historical_fit")),
                recommended_role=r.get("recommended_role"),
                category=r.get("category") or "Potential Fit",
                reasons=list(r.get("reasons") or []),
                risks=list(r.get("risks") or []),
                missing_data=list(r.get("missing_data") or []),
                evidence=r.get("evidence") or {},
                hard_filter_failures=[],
            ))
        except Exception:
            continue

    order = {"Recommended": 0, "Potential Fit": 1, "Not Recommended": 2}
    results.sort(key=lambda x: (
        order.get(x.category, 3),
        -(x.match_score if x.match_score is not None else -1),
    ))
    return results, None


# ==============================================================================
# SECTION 8 — UI HELPERS
# ==============================================================================

def src_badge(src: str) -> str:
    colors = {
        SRC_DB:      "#0ea5e9",
        SRC_DECK:    "#8b5cf6",
        SRC_CALC:    "#10b981",
        SRC_USER:    "#f59e0b",
        SRC_NA:      "#94a3b8",
        SRC_AI:      "#6366f1",
        SRC_AI_EXTR: "#6366f1",
    }
    c = colors.get(src, "#64748b")
    return (f'<span style="display:inline-block;padding:1px 8px;border-radius:999px;'
            f'background:{c}20;color:{c};font-size:.7rem;font-weight:700;'
            f'margin-left:6px;">{src}</span>')


def fmt(v: Any, default: str = "Data Not Available") -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
        return default
    return str(v)


def fmt_cr(v: Any) -> str:
    x = _safe(v)
    if x is None:
        return SRC_NA
    return f"₹{x:,.2f} Cr"


def fmt_pct(v: Any) -> str:
    x = _safe(v)
    if x is None:
        return SRC_NA
    return f"{x:.0f}/100"


def kv_card(label: str, value: str) -> str:
    """Render a Key/Value card that safely wraps long values."""
    return (
        '<div class="kpi-card" style="padding:10px 14px;margin-bottom:10px;">'
        f'<div class="k">{label}</div>'
        f'<div class="v" style="font-size:1.05rem;line-height:1.35;'
        f'word-break:break-word;white-space:normal;">{value}</div>'
        '</div>'
    )


# ==============================================================================
# SECTION 9 — PAGE CONFIG & GLOBAL STYLE
# ==============================================================================

st.set_page_config(
    page_title=f"{APP_TITLE} · {APP_TAGLINE}",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  .stApp {background: #f8fafc;}
  .block-container {padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1400px;}
  .siq-hero {
      background: linear-gradient(120deg,#0f172a 0%,#1e3a8a 55%,#0ea5e9 100%);
      padding: 26px 32px; border-radius: 14px; color: #fff;
      margin-bottom: 20px;
  }
  .siq-hero h1 {margin: 0; font-size: 1.8rem; font-weight: 800; letter-spacing:-.3px;}
  .siq-hero p  {margin: 6px 0 0; opacity: .9; font-size: .95rem;}
  .kpi-card {
      background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px;
      padding: 14px 18px; height: 100%;
  }
  .kpi-card .k {font-size:.78rem; color:#64748b; font-weight:600; letter-spacing:.3px; text-transform:uppercase;}
  .kpi-card .v {font-size:1.55rem; font-weight:800; color:#0f172a; margin-top:4px;}
  .brand-card {
      background:#ffffff; border:1px solid #e2e8f0; border-left:6px solid #0ea5e9;
      border-radius:12px; padding:16px 20px; margin-bottom:14px;
      box-shadow: 0 1px 3px rgba(0,0,0,.03);
  }
  .brand-card.recommended {border-left-color: #10b981;}
  .brand-card.potential   {border-left-color: #f59e0b;}
  .brand-card.notrec      {border-left-color: #ef4444;}
  .brand-card h4 {margin:0 0 6px; font-size:1.08rem; color:#0f172a;}
  .brand-card .meta {font-size:.82rem; color:#475569; margin-bottom:8px;}
  .brand-card .section {margin-top:10px; font-size:.86rem; color:#1e293b;}
  .brand-card .section b {color:#0f172a;}
  .pill {display:inline-block; padding:3px 11px; border-radius:999px;
         font-size:.72rem; font-weight:700; margin-right:6px;}
  .pill-green  {background:#dcfce7; color:#166534;}
  .pill-blue   {background:#dbeafe; color:#1e40af;}
  .pill-amber  {background:#fef3c7; color:#92400e;}
  .pill-red    {background:#fee2e2; color:#991b1b;}
  .pill-slate  {background:#e2e8f0; color:#475569;}
  .small {font-size:.78rem; color:#64748b;}
  .evid-row {display:flex; justify-content:space-between; font-size:.82rem;
             padding:3px 0; border-bottom:1px dashed #e2e8f0;}
  .evid-row .k {color:#64748b;}
  .evid-row .v {color:#0f172a; font-weight:600;}
</style>
""", unsafe_allow_html=True)


# ==============================================================================
# SECTION 10 — SESSION STATE DEFAULTS
# ==============================================================================
if "weights" not in st.session_state:
    st.session_state["weights"] = dict(DEFAULT_WEIGHTS)
if "thresholds" not in st.session_state:
    st.session_state["thresholds"] = dict(DEFAULT_THRESHOLDS)
if "recency" not in st.session_state:
    st.session_state["recency"] = dict(DEFAULT_RECENCY)
if "recommendation_log" not in st.session_state:
    st.session_state["recommendation_log"] = []
if "crm_additions" not in st.session_state:
    st.session_state["crm_additions"] = []
if "ai_enabled" not in st.session_state:
    st.session_state["ai_enabled"] = False
if "ai_api_key" not in st.session_state:
    st.session_state["ai_api_key"] = ""


# ==============================================================================
# SECTION 11 — SIDEBAR
# ==============================================================================
with st.sidebar:
    st.markdown(f"### 🎯 {APP_TITLE}")
    st.caption(APP_TAGLINE)
    st.divider()

    page = st.radio(
        "Navigation",
        [
            "🏠 Home",
            "🔍 Find Sponsors",
            "🏢 Brand Intelligence",
            "📤 Deck Analyzer",
            "🗂️ CRM",
        ],
        label_visibility="collapsed",
    )

    # ---------------- AI MODE ----------------
    st.divider()
    st.markdown("#### 🤖 AI Mode")

    ai_enabled = st.toggle(
        "Enable AI-powered analysis",
        value=st.session_state.get("ai_enabled", False),
        help=("When ON, decks are parsed and brands recommended by AI. "
              "The internal dataset becomes optional."),
    )
    st.session_state["ai_enabled"] = ai_enabled

    ai_cfg: Dict[str, Any] = {
        "provider": None, "api_key": "", "base_url": "", "model": ""
    }

    if ai_enabled:
        # Force-reset widget state if provider schema changed
        if st.session_state.get("ai_provider_schema") != AI_PROVIDER_SCHEMA_VERSION:
            for k in ("ai_provider", "ai_model", "ai_model_custom",
                      "ai_custom_model_toggle", "ai_custom_model_text"):
                st.session_state.pop(k, None)
            st.session_state["ai_provider_schema"] = AI_PROVIDER_SCHEMA_VERSION

        provider_names = list(AI_PROVIDERS.keys())
        default_idx = provider_names.index("Google Gemini") \
            if "Google Gemini" in provider_names else 0

        provider_name = st.selectbox(
            "Provider",
            provider_names,
            index=default_idx,
            key="ai_provider",
        )
        meta = AI_PROVIDERS[provider_name]

        # ---- SINGLE source of truth for the key: the widget key only ----
        api_key = st.text_input(
            "API Key",
            type="password",
            key="ai_api_key",
            placeholder="Paste your key here",
        )

        if provider_name == "OpenAI-compatible (custom)":
            base_url = st.text_input(
                "Base URL", key="ai_base_url",
                placeholder="https://api.example.com/v1",
            )
        else:
            base_url = meta["base_url"]
            st.caption(f"Endpoint: `{base_url}`")

        # ---- Model selection: dropdown OR free-text override ----
        allow_custom_model = st.checkbox(
            "Type custom model name",
            key="ai_custom_model_toggle",
            help=("Tick this if your provider's model isn't listed "
                  "(e.g. a new Gemini model). You can then type any name."),
        )

        if allow_custom_model:
            model = st.text_input(
                "Model name",
                value=st.session_state.get("ai_custom_model_text", "")
                      or meta.get("default_model", ""),
                key="ai_custom_model_text",
                placeholder="e.g. gemini-3.1-pro-preview",
            )
        elif meta["models"]:
            model = st.selectbox("Model", meta["models"], key="ai_model")
        else:
            model = st.text_input(
                "Model", value=meta["default_model"], key="ai_model_custom"
            )

        ai_cfg = {
            "provider": provider_name,
            "api_key": api_key,
            "base_url": base_url,
            "model": model,
        }

        st.caption(
            f"`openai` SDK: v{OPENAI_VERSION} · "
            f"`anthropic` SDK: {'installed' if ANTHROPIC_OK else 'not installed'}"
        )

        if api_key:
            if st.button("🔌 Test connection", use_container_width=True):
                with st.spinner("Testing…"):
                    ok, msg = ai_test_connection(
                        api_key, provider_name, base_url, model
                    )
                st.session_state["ai_verified"]   = ok
                st.session_state["ai_verify_msg"] = msg

            if "ai_verify_msg" in st.session_state:
                if st.session_state.get("ai_verified"):
                    st.success(st.session_state["ai_verify_msg"])
                else:
                    st.error(st.session_state["ai_verify_msg"])
        else:
            st.warning("⚠️ Enter your API key to activate AI mode.")

        with st.expander("🔧 Debug — current AI config", expanded=False):
            st.write({
                "provider": provider_name,
                "model": model,
                "base_url": base_url,
                "key_prefix": (api_key[:6] + "…") if api_key else "(empty)",
                "key_length": len(api_key) if api_key else 0,
                "openai_sdk": OPENAI_VERSION,
                "anthropic_sdk_installed": ANTHROPIC_OK,
            })

        with st.expander("📖 Where do I find the model list?", expanded=False):
            st.markdown(
                "**Gemini** — open this URL in a browser, replacing "
                "`YOUR_KEY`:\n\n"
                "`https://generativelanguage.googleapis.com/v1beta/models?key=YOUR_KEY`\n\n"
                "Every entry under `\"name\": \"models/...\"` is a model your "
                "key can call. Strip the `models/` prefix and use the rest "
                "(e.g. `gemini-3.1-pro-preview`). Then tick "
                "**Type custom model name** above and paste it in."
            )

    # ---------------- DATA SOURCE ----------------
    st.divider()
    st.markdown("#### 📦 Data Source")

    data = get_data()

    if data.brands.empty and not ai_enabled:
        st.warning("No internal dataset loaded. Upload the workbook, "
                   "or enable AI mode.")
        uploaded = st.file_uploader("Upload internal workbook (.xlsx)",
                                    type=["xlsx", "xls"])
        if uploaded is not None:
            data = get_data(uploaded_bytes=uploaded.read(),
                            uploaded_name=uploaded.name)
    elif data.brands.empty and ai_enabled:
        st.info("AI mode is on — internal dataset is optional.")
        uploaded = st.file_uploader("(Optional) Upload workbook (.xlsx)",
                                    type=["xlsx", "xls"])
        if uploaded is not None:
            data = get_data(uploaded_bytes=uploaded.read(),
                            uploaded_name=uploaded.name)
    else:
        st.success(f"Loaded: {data.loaded_from}")
        st.caption(f"Version: `{data.dataset_version}` · {data.loaded_at}")

    if not data.brands.empty:
        st.caption(
            f"Brands: **{len(data.brands):,}** · "
            f"Events: **{len(data.events):,}** · "
            f"History: **{len(data.sponsorships):,}**"
        )

    st.divider()
    st.caption(
        "🔒 In dataset mode, only internal data is used — nothing external.\n\n"
        "🤖 In AI mode, deck parsing and brand suggestions are produced by "
        "your configured LLM and clearly labelled as **AI-Suggested**."
    )


# ==============================================================================
# SECTION 12 — PAGE: HOME
# ==============================================================================
if page == "🏠 Home":
    st.markdown(f"""
    <div class="siq-hero">
      <h1>🏠 {APP_TITLE} — Executive Overview</h1>
      <p>{APP_TAGLINE} · Deterministic recommendations · 100% traceable</p>
    </div>
    """, unsafe_allow_html=True)

    if ai_enabled:
        st.info("🤖 **AI mode is ON.** Deck parsing and brand suggestions "
                "are handled by your configured LLM. Values labelled "
                "**AI-Suggested** are model output, not internal data.")

    if data.brands.empty and not ai_enabled:
        st.error("Internal dataset not loaded. Upload the workbook in the "
                 "sidebar, or enable AI mode.")
        st.stop()

    if data.brands.empty and ai_enabled:
        st.info("No internal dataset loaded — Home KPIs require the dataset. "
                "Use **Find Sponsors** or **Deck Analyzer** to work in AI mode.")
        st.stop()

    n_brands = len(data.brands)
    n_events = len(data.events)
    n_history = len(data.sponsorships)
    total_budget = 0.0
    if not data.brands.empty and "Sponsorship_Budget_Cr" in data.brands.columns:
        total_budget = data.brands["Sponsorship_Budget_Cr"].fillna(0).sum()

    active_pipeline = 0
    if not data.crm.empty and "Deal_Stage" in data.crm.columns:
        active_pipeline = len(data.crm[~data.crm["Deal_Stage"].isin(
            ["Closed Won", "Closed Lost"])])

    # ---- Avg Brand-IP Match ----
    avg_match = "Data Not Available"
    if not data.brands.empty and not data.events.empty:
        try:
            sample_brands = data.brands.head(300)
            max_events = min(len(data.events), 40)
            best_top10: List[float] = []
            best_avg = 0.0

            for _, ev_row in data.events.head(max_events).iterrows():
                ev_brief = brief_from_event_row(ev_row)
                ev_req = Requirement(
                    target_budget_cr=_safe(ev_row.get("Sponsorship_Cost_Cr"))
                )
                res = recommend_brands(
                    ev_brief, ev_req, sample_brands, data.sponsorships,
                    data.events,
                    st.session_state["weights"],
                    st.session_state["thresholds"],
                    st.session_state["recency"],
                )
                scored = sorted(
                    [r.match_score for r in res if r.match_score is not None],
                    reverse=True,
                )
                if len(scored) < 5:
                    continue
                top10 = scored[:10]
                avg = sum(top10) / len(top10)
                if avg > best_avg:
                    best_avg = avg
                    best_top10 = top10

            if best_top10:
                avg_match = f"{best_avg:.1f}%"
        except Exception:
            pass

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="kpi-card"><div class="k">Total Brands</div>'
                    f'<div class="v">{n_brands:,}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="kpi-card"><div class="k">Total Events</div>'
                    f'<div class="v">{n_events:,}</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="kpi-card"><div class="k">Sponsorship Records</div>'
                    f'<div class="v">{n_history:,}</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="kpi-card"><div class="k">Total Budget Pool</div>'
                    f'<div class="v">₹{total_budget:,.0f} Cr</div></div>',
                    unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'<div class="kpi-card"><div class="k">Active CRM Pipeline</div>'
                    f'<div class="v">{active_pipeline:,}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="kpi-card"><div class="k">Avg Brand-IP Match</div>'
                    f'<div class="v">{avg_match}</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="kpi-card"><div class="k">Scoring Version</div>'
                    f'<div class="v" style="font-size:1rem;">{SCORING_VERSION}</div></div>',
                    unsafe_allow_html=True)

    st.divider()

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Top industries by brand count")
        if "Industry" in data.brands.columns:
            top = data.brands["Industry"].value_counts().head(10)
            st.bar_chart(top)
        else:
            st.info("Industry column not available.")

    with col_b:
        st.subheader("Top event categories")
        if "Event_Category" in data.events.columns:
            top = data.events["Event_Category"].value_counts().head(10)
            st.bar_chart(top)
        else:
            st.info("Event_Category column not available.")

    st.divider()

    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader("Recent CRM activity")
        if not data.crm.empty:
            cols = [c for c in ["CRM_ID", "Brand_ID", "Deal_Stage", "Deal_Value_Cr",
                                "Probability"] if c in data.crm.columns]
            st.dataframe(data.crm[cols].head(10), use_container_width=True,
                         hide_index=True)
        else:
            st.info("CRM sheet not loaded.")

    with col_b:
        st.subheader("High-budget brands with low sponsorship history")
        if not data.brands.empty and not data.sponsorships.empty:
            hist_counts = data.sponsorships.groupby("Brand_ID").size().to_dict()
            b = data.brands.copy()
            b["_hist"] = b["Brand_ID"].astype(str).map(hist_counts).fillna(0)
            if "Sponsorship_Budget_Cr" in b.columns:
                b = b.sort_values("Sponsorship_Budget_Cr", ascending=False)
                high_budget_low_hist = b[b["_hist"] <= 1].head(10)
                cols = [c for c in ["Brand_Name", "Industry",
                                    "Sponsorship_Budget_Cr"] if c in b.columns]
                st.dataframe(high_budget_low_hist[cols],
                             use_container_width=True, hide_index=True)
        else:
            st.info("Insufficient internal data.")


# ==============================================================================
# SECTION 13 — PAGE: FIND SPONSORS  (main workflow)
# ==============================================================================
elif page == "🔍 Find Sponsors":
    st.markdown(f"""
    <div class="siq-hero">
      <h1>🔍 Find Sponsors</h1>
      <p>Upload a deck (AI or dataset mode) → set the commercial ask
         → get a ranked, explained list of best-fit brands.</p>
    </div>
    """, unsafe_allow_html=True)

    if not ai_enabled and data.brands.empty:
        st.error("Internal dataset not loaded. Either upload the workbook "
                 "in the sidebar or enable AI mode.")
        st.stop()

    # ---------- STEP 1: event source ----------
    st.subheader("1 · Event / Property")

    source_options = ["📤 Upload deck"]
    if not data.events.empty:
        source_options.append("🗃️ Select internal event")

    src_mode = st.radio(
        "How do you want to define the event?",
        source_options,
        horizontal=True,
    )

    brief: Optional[DeckBrief] = None

    if src_mode == "📤 Upload deck":
        up = st.file_uploader(
            "Drop your deck (PDF, PPTX, DOCX, XLSX)",
            type=["pdf", "pptx", "ppt", "docx", "xlsx", "xls", "txt"],
        )
        if up is not None:
            with st.spinner("Parsing deck…"):
                txt, err = parse_deck(up.name, up.read())
            if err:
                st.error(err)
            elif len(txt) < 60:
                st.warning("Very little text extracted. The deck may be image-based.")
            else:
                with st.expander(f"📄 Extracted text ({len(txt):,} chars)",
                                 expanded=False):
                    st.text_area("Raw text", txt, height=200,
                                 label_visibility="collapsed")

                if ai_enabled and ai_cfg["api_key"]:
                    with st.spinner("🤖 AI is extracting the brief…"):
                        ai_brief, ai_err = ai_extract_deck_brief(
                            txt,
                            ai_cfg["api_key"],
                            ai_cfg["provider"],
                            ai_cfg["base_url"],
                            ai_cfg["model"],
                        )
                    if ai_err:
                        st.error(f"AI extraction failed: {ai_err}")
                        st.info("Falling back to deterministic extraction.")
                        brief = extract_deck_brief(txt, data.brands, data.events)
                        brief.source = SRC_DECK
                    else:
                        brief = ai_brief
                        st.success("✅ AI extraction complete.")
                else:
                    brief = extract_deck_brief(txt, data.brands, data.events)

    else:  # select internal event
        ev_options = data.events.apply(
            lambda r: f"{r.get('Event_ID')} · {r.get('Event_Name')}", axis=1
        ).tolist()
        pick = st.selectbox("Choose an event", ["—"] + ev_options)
        if pick != "—":
            ev_id = pick.split(" · ")[0]
            row = data.events[
                data.events["Event_ID"].astype(str) == ev_id
            ].iloc[0]
            brief = brief_from_event_row(row)
            st.success(f"Selected event: **{row.get('Event_Name')}**")

    if brief is None:
        st.info("Upload a deck or pick an event to continue.")
        st.stop()

    # ---------- STEP 2: extracted brief ----------
    st.divider()
    st.subheader("2 · Extracted event brief")
    st.caption("Every value below is labelled with its source.")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Event name** {src_badge(brief.source)}",
                    unsafe_allow_html=True)
        st.write(fmt(brief.event_name))
        st.markdown(f"**Category** {src_badge(brief.source)}",
                    unsafe_allow_html=True)
        st.write(fmt(brief.event_category))
        st.markdown(f"**Cities** {src_badge(brief.source)}",
                    unsafe_allow_html=True)
        st.write(", ".join(brief.cities) if brief.cities else SRC_NA)
        st.markdown(f"**Number of cities** {src_badge(brief.source)}",
                    unsafe_allow_html=True)
        st.write(fmt(brief.number_of_cities))
    with col2:
        st.markdown(f"**Audience** {src_badge(brief.source)}",
                    unsafe_allow_html=True)
        st.write(fmt(brief.audience_text))
        st.markdown(f"**Reach** {src_badge(brief.source)}",
                    unsafe_allow_html=True)
        st.write(fmt(brief.reach_text))
        st.markdown(f"**Footfall** {src_badge(brief.source)}",
                    unsafe_allow_html=True)
        st.write(fmt(brief.footfall_text))
        st.markdown(f"**Sponsorship value range** {src_badge(brief.source)}",
                    unsafe_allow_html=True)
        if (brief.sponsorship_value_min_cr is not None
                or brief.sponsorship_value_max_cr is not None):
            lo = brief.sponsorship_value_min_cr or "?"
            hi = brief.sponsorship_value_max_cr or "?"
            st.write(f"₹{lo} – ₹{hi} Cr")
        else:
            st.write(SRC_NA)

    st.caption(
        f"Extraction confidence: **{brief.extraction_confidence}** · "
        f"Missing fields: "
        f"{', '.join(brief.missing_fields) if brief.missing_fields else 'none'}"
    )

    # ---------- STEP 3: commercial requirements ----------
    st.divider()
    st.subheader("3 · Commercial requirement")

    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        role = st.selectbox(
            "Sponsorship role required",
            ["Title Sponsor", "Presenting Sponsor", "Co-Sponsor",
             "Associate Sponsor", "Official Partner", "Category Partner",
             "Media Partner", "Any"],
            index=0,
        )
    with rc2:
        default_budget = (brief.sponsorship_value_max_cr
                          or brief.sponsorship_value_min_cr or 0.0)
        budget = st.number_input(
            "Target sponsorship value (₹ Cr)",
            min_value=0.0, value=float(default_budget), step=0.5,
            help="Required sponsorship value the brand must be able to fund.",
        )
    with rc3:
        target_aud = st.text_input(
            "Target audience override (optional)",
            value=brief.audience_text or "",
            help="Leave blank to use the deck's audience.",
        )

    rc4, rc5 = st.columns(2)
    with rc4:
        cities_input = st.text_input(
            "Mandatory cities (comma-separated, optional)",
            value=", ".join(brief.cities) if brief.cities else "",
        )
    with rc5:
        industry_options = ["Any"]
        if (not data.brands.empty
                and "Industry" in data.brands.columns):
            industry_options += sorted(
                {str(x).strip() for x in data.brands["Industry"].dropna().unique()}
            )
        target_industry = st.selectbox("Preferred industry (optional)",
                                       industry_options)

    exclusivity = st.checkbox("Category exclusivity required", value=False)

    user_override = False
    if default_budget and abs(budget - default_budget) > 0.01:
        user_override = True
        st.info("**User Override Applied** — the target budget you entered "
                "takes priority over the deck's stated range.")

    run = st.button("🚀 Find Best-Matching Brands", type="primary",
                    use_container_width=True)

    if run:
        if target_aud:
            brief.audience_text = target_aud
            brief.audience_tags = sorted(extract_audience_tags(target_aud))
        if cities_input.strip():
            brief.cities = [c.strip() for c in cities_input.split(",") if c.strip()]
            brief.number_of_cities = len(brief.cities)

        req = Requirement(
            sponsorship_role=None if role == "Any" else role,
            target_budget_cr=budget if budget > 0 else None,
            target_cities=brief.cities,
            target_audience=target_aud or None,
            target_industry=None if target_industry == "Any" else target_industry,
            exclusivity_required=exclusivity,
            user_override=user_override,
        )

        results: List[MatchResult] = []

        if ai_enabled and ai_cfg["api_key"]:
            with st.spinner("🤖 AI is evaluating sponsor matches…"):
                results, ai_err = ai_recommend_brands(
                    brief, req,
                    ai_cfg["api_key"],
                    ai_cfg["provider"],
                    ai_cfg["base_url"],
                    ai_cfg["model"],
                )
            if ai_err:
                st.error(f"AI recommendation failed: {ai_err}")
                if not data.brands.empty:
                    st.info("Falling back to internal dataset engine.")
                    results = recommend_brands(
                        brief, req, data.brands, data.sponsorships,
                        data.events,
                        st.session_state["weights"],
                        st.session_state["thresholds"],
                        st.session_state["recency"],
                    )
        else:
            with st.spinner("Scoring every brand in the internal dataset…"):
                results = recommend_brands(
                    brief, req, data.brands, data.sponsorships,
                    data.events,
                    st.session_state["weights"],
                    st.session_state["thresholds"],
                    st.session_state["recency"],
                )

        recommended = [r for r in results if r.category == "Recommended"]
        potential   = [r for r in results if r.category == "Potential Fit"]
        not_rec     = [r for r in results if r.category == "Not Recommended"]

        st.session_state["recommendation_log"].append({
            "timestamp": datetime.utcnow().isoformat(),
            "mode": "AI" if (ai_enabled and ai_cfg["api_key"]) else "Dataset",
            "event_name": brief.event_name,
            "event_category": brief.event_category,
            "role": role,
            "target_budget_cr": budget,
            "n_brands_total": len(results),
            "n_recommended": len(recommended),
            "n_potential": len(potential),
            "scoring_version": SCORING_VERSION,
            "dataset_version": data.dataset_version,
            "user_override": user_override,
        })

        st.divider()
        st.subheader("4 · Recommendation summary")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Brands evaluated", f"{len(results):,}")
        m2.metric("Recommended", f"{len(recommended)}")
        m3.metric("Potential Fit", f"{len(potential)}")
        m4.metric("Not Recommended", f"{len(not_rec)}")

        if not recommended and not potential:
            st.warning(
                "**No brand currently satisfies all mandatory criteria.** "
                "Closest available matches are shown below."
            )

        def render_card(r: MatchResult, _idx: int):
            cls = {"Recommended": "recommended",
                   "Potential Fit": "potential",
                   "Not Recommended": "notrec"}.get(r.category, "notrec")
            score_txt = f"{r.match_score:.0f}/100" \
                if r.match_score is not None else SRC_NA
            conf_pill = {"High": "pill-green", "Medium": "pill-amber",
                         "Low": "pill-red"}.get(r.confidence, "pill-slate")
            cat_pill = {"Recommended": "pill-green",
                        "Potential Fit": "pill-amber",
                        "Not Recommended": "pill-red"}.get(r.category, "pill-slate")
            is_ai = r.brand_id.startswith("AI-")
            src_pill = (
                f'<span class="pill pill-blue">AI-Suggested</span>'
                if is_ai else
                f'<span class="pill pill-slate">Internal Dataset</span>'
            )

            st.markdown(f"""
            <div class="brand-card {cls}">
              <h4>{r.brand_name}
                 <span class="small">· {fmt(r.industry)}</span>
                 <span class="pill {cat_pill}">{r.category}</span>
                 <span class="pill {conf_pill}">{r.confidence} Confidence</span>
                 {src_pill}
              </h4>
              <div class="meta">
                 Match Score: <b>{score_txt}</b> &nbsp;·&nbsp;
                 Data Completeness: <b>{r.data_completeness:.0f}%</b> &nbsp;·&nbsp;
                 Recommended Role: <b>{fmt(r.recommended_role)}</b>
              </div>
            </div>
            """, unsafe_allow_html=True)

            with st.expander(
                f"🔎 Details, evidence & component scores — {r.brand_name}"
            ):
                cols = st.columns(4)
                cols[0].metric("Audience Fit", fmt_pct(r.audience_fit))
                cols[1].metric("Budget Fit", fmt_pct(r.budget_fit))
                cols[2].metric("Industry Fit", fmt_pct(r.industry_fit))
                cols[3].metric("Geography Fit", fmt_pct(r.geography_fit))

                cols = st.columns(3)
                cols[0].metric("Brand Value Fit", fmt_pct(r.brand_value_fit))
                cols[1].metric("Competition Fit", fmt_pct(r.competition_fit))
                cols[2].metric("Historical Relevance", fmt_pct(r.historical_fit))

                st.markdown("**Why this brand matches**")
                for reason in r.reasons:
                    st.markdown(f"- {reason}")

                if r.risks:
                    st.markdown("**Risks & limitations**")
                    for risk in r.risks:
                        st.markdown(f"- {risk}")

                if r.missing_data:
                    st.markdown("**Unable to evaluate**")
                    for m in r.missing_data:
                        st.markdown(f"- {str(m).replace('_',' ').title()}")

                if r.hard_filter_failures:
                    st.error("**Hard filter failures:** "
                             + ", ".join(r.hard_filter_failures))

                st.markdown("**Evidence**")
                e = r.evidence or {}
                if is_ai:
                    for k, v in e.items():
                        if v in (None, "", [], {}):
                            continue
                        st.markdown(
                            f'<div class="evid-row">'
                            f'<span class="k">{str(k).replace("_"," ").title()}</span>'
                            f'<span class="v">{fmt(v)}</span></div>',
                            unsafe_allow_html=True,
                        )
                else:
                    rows = [
                        ("Sponsorship Budget (brand)",
                         fmt_cr(e.get("sponsorship_budget_cr"))),
                        ("Required Sponsorship",
                         fmt_cr(e.get("required_budget_cr"))),
                        ("Industry", fmt(e.get("industry"))),
                        ("Sub-category", fmt(e.get("sub_category"))),
                        ("Brand Target Audience", fmt(e.get("target_audience"))),
                        ("Brand HQ",
                         f"{fmt(e.get('hq_city'))}, {fmt(e.get('hq_state'))}"),
                        ("Brand Type", fmt(e.get("brand_type"))),
                        ("Event Category", fmt(e.get("event_category"))),
                        ("Event Cities",
                         ", ".join(e.get("event_cities") or []) or SRC_NA),
                    ]
                    for k, v in rows:
                        st.markdown(
                            f'<div class="evid-row"><span class="k">{k}</span>'
                            f'<span class="v">{v}</span></div>',
                            unsafe_allow_html=True,
                        )
                    hist = e.get("historical") or {}
                    if hist:
                        st.markdown(
                            f"<div class='small' style='margin-top:8px'>"
                            f"Historical: {hist.get('previous_sponsorships',0)} prior "
                            f"sponsorships · "
                            f"{hist.get('relevant_category_sponsorships',0)} "
                            f"in same category · latest year "
                            f"{hist.get('latest_year', SRC_NA)}</div>",
                            unsafe_allow_html=True,
                        )

                # ---- CRM action ----
                c1, c2 = st.columns([1, 3])
                with c1:
                    if st.button("➕ Add to CRM",
                                 key=f"crm_add_{_idx}_{r.brand_id}"):
                        entry = {
                            "Brand_ID": r.brand_id,
                            "Brand_Name": r.brand_name,
                            "Event": brief.event_name,
                            "Role": r.recommended_role,
                            "Match_Score": r.match_score,
                            "Deal_Value_Cr": req.target_budget_cr,
                            "Source": "AI-Suggested" if is_ai else "Internal Dataset",
                            "Timestamp": datetime.utcnow().isoformat(),
                        }
                        bucket = st.session_state["crm_additions"]
                        dup = any(
                            e.get("Brand_ID") == entry["Brand_ID"]
                            and e.get("Event") == entry["Event"]
                            for e in bucket
                        )
                        if dup:
                            st.info(f"{r.brand_name} is already in the staging list.")
                        else:
                            bucket.append(entry)
                            st.success(f"Added {r.brand_name} to CRM staging list.")
                            st.rerun()
                with c2:
                    st.caption(
                        "Adding to CRM creates a staging record only — "
                        "it does not overwrite any existing CRM rows."
                    )

        if recommended:
            st.markdown("### ✅ Recommended")
            st.caption(f"{len(recommended)} brands — strong fit.")
            for i, r in enumerate(recommended[:30]):
                render_card(r, i)

        if potential:
            st.markdown("### ⚠️ Potential Fit")
            st.caption(f"{len(potential)} brands — strong signals but incomplete.")
            for i, r in enumerate(potential[:15]):
                render_card(r, 1000 + i)

        if not_rec:
            with st.expander(f"❌ Not Recommended ({len(not_rec)})"):
                for r in not_rec[:40]:
                    st.markdown(
                        f"**{r.brand_name}** — "
                        f"{', '.join(r.hard_filter_failures) or r.category}  "
                        f"<span class='small'>(Match: "
                        f"{r.match_score if r.match_score is not None else SRC_NA}"
                        f" · Completeness: {r.data_completeness:.0f}%)</span>",
                        unsafe_allow_html=True,
                    )

        st.divider()
        st.subheader("5 · Export")
        export_rows = []
        for r in results:
            export_rows.append({
                "Brand_ID": r.brand_id,
                "Brand_Name": r.brand_name,
                "Industry": r.industry,
                "Source": "AI-Suggested" if r.brand_id.startswith("AI-")
                          else "Internal Dataset",
                "Category": r.category,
                "Match_Score": r.match_score,
                "Confidence": r.confidence,
                "Data_Completeness_%": r.data_completeness,
                "Audience_Fit": r.audience_fit,
                "Budget_Fit": r.budget_fit,
                "Industry_Fit": r.industry_fit,
                "Geography_Fit": r.geography_fit,
                "Brand_Value_Fit": r.brand_value_fit,
                "Competition_Fit": r.competition_fit,
                "Historical_Fit": r.historical_fit,
                "Recommended_Role": r.recommended_role,
                "Why_Match": " | ".join(r.reasons),
                "Risks": " | ".join(r.risks),
                "Missing_Data": ", ".join(r.missing_data),
                "Hard_Filter_Failures": ", ".join(r.hard_filter_failures),
            })
        export_df = pd.DataFrame(export_rows)
        st.download_button(
            "⬇️ Download recommendation report (CSV)",
            data=export_df.to_csv(index=False).encode("utf-8"),
            file_name="sponsoriq_recommendations.csv",
            mime="text/csv",
            use_container_width=True,
        )

        with st.expander("📋 Full recommendation table"):
            st.dataframe(export_df, use_container_width=True, hide_index=True)


# ==============================================================================
# SECTION 15 — PAGE: BRAND INTELLIGENCE
# ==============================================================================
elif page == "🏢 Brand Intelligence":
    st.markdown("""
    <div class="siq-hero">
      <h1>🏢 Brand Intelligence</h1>
      <p>Deep-dive into a single brand using only internal data.</p>
    </div>
    """, unsafe_allow_html=True)

    if data.brands.empty:
        st.error("Brands dataset not loaded. This page requires the internal "
                 "workbook — it does not run in AI-only mode.")
        st.stop()

    opts = data.brands.apply(
        lambda r: f"{r.get('Brand_ID')} · {r.get('Brand_Name')}", axis=1).tolist()
    pick = st.selectbox("Select brand", ["—"] + opts)
    if pick == "—":
        st.stop()
    brand_id = pick.split(" · ")[0]
    brand = data.brands[data.brands["Brand_ID"].astype(str) == brand_id].iloc[0]

    # ---------- Header ----------
    st.subheader(f"{fmt(brand.get('Brand_Name'))}")
    st.caption(f"Brand ID: {fmt(brand.get('Brand_ID'))} · "
               f"Parent: {fmt(brand.get('Parent_Company'))} · "
               f"HQ: {fmt(brand.get('HQ_City'))}, {fmt(brand.get('HQ_State'))}")

    # ---------- Identity cards ----------
    st.markdown("##### Identity")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(kv_card("Industry", fmt(brand.get("Industry"))),
                    unsafe_allow_html=True)
    with c2:
        st.markdown(kv_card("Sub-category", fmt(brand.get("Sub_Category"))),
                    unsafe_allow_html=True)
    with c3:
        st.markdown(kv_card("Brand Type", fmt(brand.get("Brand_Type"))),
                    unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(kv_card("Target Audience", fmt(brand.get("Target_Audience"))),
                    unsafe_allow_html=True)
    with c2:
        st.markdown(kv_card("HQ City", fmt(brand.get("HQ_City"))),
                    unsafe_allow_html=True)
    with c3:
        st.markdown(kv_card("HQ State", fmt(brand.get("HQ_State"))),
                    unsafe_allow_html=True)

    # ---------- Financials ----------
    st.markdown("##### Financials")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(kv_card("Annual Revenue", fmt_cr(brand.get("Annual_Revenue_Cr"))),
                    unsafe_allow_html=True)
    with c2:
        st.markdown(kv_card("Marketing Spend", fmt_cr(brand.get("Marketing_Spend_Cr"))),
                    unsafe_allow_html=True)
    with c3:
        st.markdown(kv_card("Sponsorship Budget",
                            fmt_cr(brand.get("Sponsorship_Budget_Cr"))),
                    unsafe_allow_html=True)
    with c4:
        st.markdown(kv_card("Employees", fmt(brand.get("Employee_Count"))),
                    unsafe_allow_html=True)

    # ---------- Full record ----------
    with st.expander("📋 Full brand record (all columns from internal dataset)"):
        full = pd.DataFrame({
            "Field": [str(c) for c in data.brands.columns],
            "Value": [fmt(brand.get(c)) for c in data.brands.columns],
        })
        st.dataframe(full, use_container_width=True, hide_index=True)

    st.divider()

    # ---------- Historical ----------
    st.subheader("Historical sponsorship records")
    if not data.sponsorships.empty:
        hist = data.sponsorships[data.sponsorships["Brand_ID"].astype(str) == brand_id]
        if hist.empty:
            st.info("No sponsorship records found for this brand in the internal dataset.")
        else:
            merged = hist.merge(
                data.events[["Event_ID", "Event_Name", "Event_Category"]],
                on="Event_ID", how="left")
            show = [c for c in ["Year", "Event_Name", "Event_Category",
                                "Sponsorship_Type", "Investment_Cr", "Status"]
                    if c in merged.columns]
            st.dataframe(merged[show], use_container_width=True, hide_index=True)
    else:
        st.info("Sponsorship history dataset not loaded.")

    # ---------- Decision makers ----------
    st.subheader("Decision makers")
    if not data.decision_makers.empty:
        dm = data.decision_makers[
            data.decision_makers["Brand_ID"].astype(str) == brand_id]
        if dm.empty:
            st.info("No decision-maker records for this brand.")
        else:
            show = [c for c in ["Name", "Designation", "Department",
                                "Email", "Phone"] if c in dm.columns]
            st.dataframe(dm[show], use_container_width=True, hide_index=True)
            st.caption("⚠️ Decision-maker data is placeholder/synthetic per the "
                       "dataset notes — do not use for real outreach.")
    else:
        st.info("Decision makers dataset not loaded.")

    # ---------- CRM ----------
    st.subheader("CRM activity")
    if not data.crm.empty:
        crm = data.crm[data.crm["Brand_ID"].astype(str) == brand_id]
        if crm.empty:
            st.info("No CRM records for this brand.")
        else:
            show = [c for c in ["Deal_Stage", "Deal_Value_Cr", "Probability",
                                "Proposal_Status", "Meeting_Date"]
                    if c in crm.columns]
            st.dataframe(crm[show], use_container_width=True, hide_index=True)

    # ---------- Marketing campaigns ----------
    st.subheader("Marketing campaigns")
    if not data.campaigns.empty:
        cmp_ = data.campaigns[data.campaigns["Brand_ID"].astype(str) == brand_id]
        if cmp_.empty:
            st.info("No campaign records for this brand.")
        else:
            show = [c for c in ["Campaign_Name", "Campaign_Type", "Platform",
                                "Budget_Cr", "ROI", "Start_Date"]
                    if c in cmp_.columns]
            st.dataframe(cmp_[show], use_container_width=True, hide_index=True)


# ==============================================================================
# SECTION 17 — PAGE: DECK ANALYZER (standalone)
# ==============================================================================
elif page == "📤 Deck Analyzer":
    st.markdown("""
    <div class="siq-hero">
      <h1>📤 Deck Analyzer</h1>
      <p>Upload a sponsorship deck and inspect exactly what the system
         extracted — every field is labelled with its source.</p>
    </div>
    """, unsafe_allow_html=True)

    st.caption("Supported formats: **PDF · PPTX · DOCX · XLSX · TXT**")

    up = st.file_uploader(
        "Drop your deck",
        type=["pdf", "pptx", "ppt", "docx", "xlsx", "xls", "txt"],
    )
    if up is None:
        st.stop()

    with st.spinner("Parsing…"):
        txt, err = parse_deck(up.name, up.read())
    if err:
        st.error(err)
        st.stop()

    st.success(f"File: **{up.name}** · {len(txt):,} characters extracted")

    with st.expander("📄 Raw extracted text", expanded=False):
        st.text_area("Raw text", txt, height=250, label_visibility="collapsed")

    if ai_enabled and ai_cfg["api_key"]:
        with st.spinner("🤖 AI is extracting the brief…"):
            ai_brief, ai_err = ai_extract_deck_brief(
                txt,
                ai_cfg["api_key"],
                ai_cfg["provider"],
                ai_cfg["base_url"],
                ai_cfg["model"],
            )
        if ai_err:
            st.error(f"AI extraction failed: {ai_err}")
            st.info("Falling back to deterministic extraction.")
            brief = extract_deck_brief(txt, data.brands, data.events)
        else:
            brief = ai_brief
            st.success("✅ AI extraction complete.")
    else:
        brief = extract_deck_brief(txt, data.brands, data.events)

    st.subheader("Extracted brief")
    st.caption(f"Extraction confidence: **{brief.extraction_confidence}** · "
               f"Source: **{brief.source}**")

    edited = st.data_editor(
        pd.DataFrame([
            {"Field": "Event Name", "Value": brief.event_name or ""},
            {"Field": "Event Category", "Value": brief.event_category or ""},
            {"Field": "Artist/Property", "Value": brief.artist_or_property or ""},
            {"Field": "Cities", "Value": ", ".join(brief.cities)},
            {"Field": "Number of Cities", "Value": str(brief.number_of_cities or "")},
            {"Field": "Audience", "Value": brief.audience_text or ""},
            {"Field": "Reach", "Value": brief.reach_text or ""},
            {"Field": "Footfall", "Value": brief.footfall_text or ""},
            {"Field": "Duration", "Value": brief.duration_text or ""},
            {"Field": "Sponsorship Roles", "Value": ", ".join(brief.sponsorship_roles)},
            {"Field": "Sponsorship Value Min (Cr)",
             "Value": str(brief.sponsorship_value_min_cr or "")},
            {"Field": "Sponsorship Value Max (Cr)",
             "Value": str(brief.sponsorship_value_max_cr or "")},
            {"Field": "Category Exclusivity", "Value": brief.category_exclusivity or ""},
            {"Field": "Deliverables", "Value": ", ".join(brief.deliverables)},
        ]),
        use_container_width=True,
        hide_index=True,
        disabled=["Field"],
        key="deck_analyzer_edit",
    )
    st.caption("You can edit any value above — corrections are used when "
               "generating recommendations on the Find Sponsors page.")

    st.subheader("Missing fields")
    if brief.missing_fields:
        st.warning("The following fields could not be extracted from the deck: "
                   + ", ".join(brief.missing_fields))
    else:
        st.success("All key fields were extracted.")


# ==============================================================================
# SECTION 20 — PAGE: CRM
# ==============================================================================
elif page == "🗂️ CRM":
    st.markdown("""
    <div class="siq-hero">
      <h1>🗂️ CRM</h1>
      <p>CRM pipeline view + staging list of brands you added from
         recommendations.</p>
    </div>
    """, unsafe_allow_html=True)

    # ----------------------------------------------------------------------
    # A · Staging list
    # ----------------------------------------------------------------------
    staged = st.session_state.get("crm_additions", [])
    st.subheader(f"Staging list — added from recommendations ({len(staged)})")

    if not staged:
        st.info("No brands added yet. Use the **Find Sponsors** page and click "
                "**Add to CRM** on any recommendation.")
    else:
        staged_df = pd.DataFrame(staged)

        expected_cols = ["Brand_ID", "Brand_Name", "Event", "Role",
                         "Match_Score", "Deal_Value_Cr", "Source", "Timestamp"]
        for c in expected_cols:
            if c not in staged_df.columns:
                staged_df[c] = None
        staged_df = staged_df[expected_cols]

        st.dataframe(staged_df, use_container_width=True, hide_index=True)

        c1, c2 = st.columns([3, 1])
        with c1:
            st.download_button(
                "⬇️ Download staging list (CSV)",
                data=staged_df.to_csv(index=False).encode("utf-8"),
                file_name="crm_staging.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with c2:
            if st.button("🗑️ Clear staging list", use_container_width=True):
                st.session_state["crm_additions"] = []
                st.success("Staging list cleared.")
                st.rerun()

    st.divider()

    # ----------------------------------------------------------------------
    # B · Internal CRM pipeline
    # ----------------------------------------------------------------------
    st.subheader("Internal CRM pipeline")

    if data.crm is None or data.crm.empty:
        st.info("CRM sheet not loaded.")
    else:
        df = data.crm.copy()

        try:
            if (not data.brands.empty
                    and "Brand_ID" in df.columns
                    and "Brand_ID" in data.brands.columns
                    and "Brand_Name" in data.brands.columns):
                brands_slim = (data.brands[["Brand_ID", "Brand_Name"]]
                               .dropna(subset=["Brand_ID"])
                               .drop_duplicates(subset=["Brand_ID"]))
                df = df.merge(brands_slim, on="Brand_ID", how="left")
        except Exception as e:
            st.warning(f"Could not attach brand names: {e}")

        if "Deal_Stage" in df.columns and df["Deal_Stage"].notna().any():
            stages = (df["Deal_Stage"]
                      .fillna("Unknown")
                      .astype(str)
                      .value_counts())
            st.markdown("**Pipeline by deal stage**")
            st.bar_chart(stages)
        else:
            st.info("No `Deal_Stage` values available for a pipeline chart.")

        all_cols = list(df.columns)
        default_cols = [c for c in
                        ["CRM_ID", "Brand_Name", "Brand_ID", "Deal_Stage",
                         "Deal_Value_Cr", "Probability", "Proposal_Status",
                         "Meeting_Date"] if c in all_cols]
        chosen = st.multiselect(
            "Columns to show",
            options=all_cols,
            default=default_cols or all_cols[:8],
        )
        st.dataframe(
            df[chosen] if chosen else df,
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "⬇️ Download CRM records (CSV)",
            data=(df[chosen] if chosen else df).to_csv(index=False).encode("utf-8"),
            file_name="crm_records.csv",
            mime="text/csv",
        )


# ==============================================================================
# SECTION 23 — FOOTER
# ==============================================================================
st.divider()
st.caption(
    "**SponsorIQ** · Closed-data sponsorship intelligence platform · "
    "Every value is sourced as Known-from-Dataset, Calculated, User-Input, "
    "AI-Extracted, AI-Suggested, or Data-Not-Available. "
    "Dataset mode uses no external data. AI mode is opt-in and every AI "
    "output is clearly labelled."
)