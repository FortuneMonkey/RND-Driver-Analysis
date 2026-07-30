import streamlit as st

ACCENT = "#D9631E"   # logbook / brand accent (burnt orange, readable on white)
CYAN = "#0E8C8A"     # GPS / secondary accent (deep teal)
AMBER = "#B8790F"    # warning / HIGH variance
GREEN = "#2E9E5B"    # OK / success
RED = "#C0392B"      # alert / STANDBY
DIM = "#6B7280"      # muted text/labels
GRID = "#E3E6EA"     # gridlines / borders
PANEL = "#FFFFFF"    # card/panel background
BG = "#F7F7F5"       # page background
TEXT = "#1F2328"     # primary text

FLAG_COLORS = {"OK": GREEN, "HIGH": AMBER, "STANDBY": RED, "NO-GPS": DIM, "NO-LOGBOOK": "#9333EA"}
FLAG_ICONS = {"OK": "✅", "HIGH": "⚠️", "STANDBY": "🚩", "NO-GPS": "❔", "NO-LOGBOOK": "👻"}


def apply_theme():
    """Inject the light theme CSS into the current Streamlit page. Call once, early."""
    st.markdown(f"""
<style>
    .stApp {{ background-color: {BG}; }}
    section[data-testid="stSidebar"] {{ background-color: #FFFFFF; border-right: 1px solid {GRID}; }}
    div[data-testid="stMetric"] {{
        background-color: {PANEL}; border: 1px solid {GRID}; border-radius: 10px;
        padding: 14px 16px 10px; border-left: 3px solid {ACCENT};
        box-shadow: 0 1px 2px rgba(20,20,20,0.04);
    }}
    div[data-testid="stMetricLabel"] {{ color: {DIM} !important; font-size: 11px; letter-spacing: .08em; text-transform: uppercase; }}
    div[data-testid="stMetricValue"] {{ font-family: 'Courier New', monospace; color: {TEXT} !important; font-size: 1.7rem !important; }}
    div[data-testid="stMetricDelta"] {{ font-family: 'Courier New', monospace; }}
    h1, h2, h3 {{ color: {TEXT} !important; }}
    p, span, label, .stCaption {{ color: #45494F; }}
    .stDataFrame {{ border: 1px solid {GRID}; border-radius: 10px; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 4px; border-bottom: 1px solid {GRID}; }}
    .stTabs [data-baseweb="tab"] {{
        background-color: #FBFBFA; border-radius: 8px 8px 0 0; padding: 8px 18px;
        color: {DIM}; border: 1px solid {GRID}; border-bottom: none;
    }}
    .stTabs [aria-selected="true"] {{
        color: {TEXT} !important; background-color: {PANEL} !important;
        border-top: 2px solid {ACCENT}; font-weight: 600;
    }}
    .insight-box {{
        background: linear-gradient(135deg, rgba(217,99,30,0.07), rgba(14,140,138,0.05));
        border: 1px solid {GRID}; border-left: 3px solid {ACCENT};
        border-radius: 10px; padding: 14px 18px; margin-bottom: 6px;
        font-size: 14px; line-height: 1.65; color: {TEXT};
        box-shadow: 0 1px 2px rgba(20,20,20,0.03);
    }}
    .insight-box b {{ color: {TEXT}; }}
    .badge {{
        display:inline-block; padding: 2px 10px; border-radius: 20px;
        font-size: 11px; font-weight: 600; font-family: 'Courier New', monospace;
    }}
    hr {{ border-color: {GRID} !important; }}
    div[data-testid="stFileUploaderDropzone"] {{ background-color: #FBFBFA; border-color: {GRID}; }}
    .stButton button, .stDownloadButton button {{
        border-radius: 8px; border: 1px solid {GRID};
    }}
    .stButton button[kind="primary"], .stDownloadButton button[kind="primary"] {{
        background-color: {ACCENT}; border-color: {ACCENT};
    }}
</style>
""", unsafe_allow_html=True)
