import re
import io
import os
import tempfile
import datetime as dt
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image
from fpdf import FPDF

st.set_page_config(page_title="RND Driver Analysis", layout="wide", page_icon="🛰️")

# ---------------------------------------------------------------------------
# Color palette — light "fleet ops / telemetry console" theme.
# Defined once here so the CSS below and every chart later in the file stay
# in sync with a single source of truth.
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Styling — light "fleet ops / telemetry console" theme
# ---------------------------------------------------------------------------
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
    .stDataFrame {{ border: 1px solid {GRID}; border-radius: 10px; overflow: hidden; }}
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


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
def norm_plate(s):
    """Extract a normalized plate code (e.g. 'BM 1356 CW') from a string."""
    if not isinstance(s, str) or not s.strip():
        return ""
    s2 = s.upper()
    m = re.search(r"[A-Z]{1,2}\s?\d{2,4}\s?[A-Z]{1,3}", s2)
    if m:
        return re.sub(r"\s+", " ", m.group(0)).strip()
    return s2.strip()


@st.cache_data(show_spinner=False)
def parse_logbook(file_bytes):
    xls = pd.ExcelFile(file_bytes)
    result = {}
    for sheet in xls.sheet_names:
        plate = norm_plate(sheet)
        df = pd.read_excel(xls, sheet_name=sheet, header=0, skiprows=[1])
        if df.shape[1] < 7:
            continue
        cols = ["No", "Tujuan", "Berangkat_Tgl", "Berangkat_KM", "Tiba_Tgl",
                "Tiba_KM", "Jumlah_Pemakaian", "Refuel_L", "Fuel_Consum", "Keterangan"]
        df.columns = cols[: df.shape[1]]
        df = df[pd.to_numeric(df["No"], errors="coerce").notna()].copy()
        if df.empty:
            continue
        df["Date"] = pd.to_datetime(df["Berangkat_Tgl"]).dt.date
        df["Jumlah_Pemakaian"] = pd.to_numeric(df["Jumlah_Pemakaian"], errors="coerce").fillna(0)
        df["Tujuan"] = df["Tujuan"].fillna("")
        result[plate] = df.reset_index(drop=True)
    return result


@st.cache_data(show_spinner=False)
def parse_gps(file_bytes):
    xls = pd.ExcelFile(file_bytes)
    sheet = next((s for s in xls.sheet_names if "report" in s.lower()), xls.sheet_names[0])
    raw = pd.read_excel(xls, sheet_name=sheet, header=None)

    header_row = None
    for i in range(min(10, len(raw))):
        row_str = raw.iloc[i].astype(str)
        if row_str.str.contains("GPS Time", na=False).any() and not row_str.str.contains("UTC", na=False).any():
            header_row = i
            break
    if header_row is None:
        header_row = 2

    df = pd.read_excel(xls, sheet_name=sheet, header=header_row)
    veh_fallback = norm_plate(str(raw.iloc[0, 1])) if raw.shape[0] > 0 and raw.shape[1] > 1 else ""

    if "Vehicle Code" in df.columns:
        df["Plate"] = df["Vehicle Code"].apply(lambda x: norm_plate(str(x)) or veh_fallback)
    else:
        df["Plate"] = veh_fallback

    df["GPS Time"] = pd.to_datetime(df["GPS Time"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    df["Date"] = df["GPS Time"].dt.date

    out = {}
    raw_out = {}
    for plate, g in df.groupby("Plate"):
        if not plate:
            continue
        daily = g.groupby("Date").agg(
            gps_km=("Mileage(KM)", lambda x: x.max() - x.min()),
            max_speed=("Speed (Km/hr)", "max"),
            acc_on=("ACC", lambda x: (x == "ON").sum()),
            pings=("Mileage(KM)", "count"),
        ).reset_index()
        out[plate] = daily
        raw_out[plate] = g[["GPS Time", "Speed (Km/hr)", "ACC", "Mileage(KM)"]].sort_values("GPS Time").reset_index(drop=True)
    return out, raw_out


def compute_reconciliation(lb_df, gps_df):
    if gps_df is not None:
        # outer join so GPS-only days (movement with no logbook entry at all)
        # are surfaced instead of silently dropped
        merged = lb_df.merge(gps_df, on="Date", how="outer")
    else:
        merged = lb_df.copy()

    if "gps_km" not in merged.columns:
        merged["gps_km"] = pd.NA
        merged["max_speed"] = pd.NA
        merged["acc_on"] = pd.NA

    merged = merged.sort_values("Date").reset_index(drop=True)
    merged["Tujuan"] = merged["Tujuan"].fillna("(no logbook entry)")
    # Jumlah_Pemakaian stays NaN (not 0) for GPS-only days — 0 means "logged as standby",
    # NaN means "never logged at all", and those are very different situations.

    merged["variance_km"] = merged["Jumlah_Pemakaian"] - merged["gps_km"]
    merged["variance_pct"] = merged.apply(
        lambda r: (r["variance_km"] / r["Jumlah_Pemakaian"])
        if pd.notna(r["Jumlah_Pemakaian"]) and r["Jumlah_Pemakaian"] != 0 and pd.notna(r["gps_km"])
        else float("nan"),  # proper NaN, not None — keeps this a numeric dtype column so
        axis=1,              # .abs() and other numeric ops downstream never break on it
    )
    merged["variance_pct"] = pd.to_numeric(merged["variance_pct"], errors="coerce")

    def flag(r):
        if pd.isna(r["Jumlah_Pemakaian"]):
            # No logbook entry at all for this date
            if pd.isna(r["gps_km"]):
                return "NO-GPS"  # no data from either source — nothing to say
            return "NO-LOGBOOK" if r["gps_km"] > 1 else "OK"  # GPS covers it: real unlogged trip, or genuinely idle (fine)
        if pd.isna(r["gps_km"]):
            return "NO-GPS"
        if r["Jumlah_Pemakaian"] == 0 and r["gps_km"] > 1:
            return "STANDBY"
        if r["variance_pct"] is not None and abs(r["variance_pct"]) > 0.1:
            return "HIGH"
        return "OK"

    merged["flag"] = merged.apply(flag, axis=1)
    return merged



def compute_fuel_overview(data, target_kml):
    """
    Cycle-free fuel overview. Instead of segmenting by refuel event (which
    assumes every fill-up tops the tank to exactly full — noisy with few
    fill-ups or partial fills), this just uses two totals over the whole
    period: total liters refueled and total km driven (GPS-preferred, else
    logbook). The only assumption left is the tank level at the very start
    and end of the period, a small, fixed edge-effect rather than compounding
    per-fill-up noise.

    Also returns a day-by-day cumulative comparison — cumulative liters
    actually used vs. cumulative liters implied by the target km/L for the
    distance driven so far — so you can see the gap trend smoothly over time
    instead of in noisy discrete per-cycle jumps.
    """
    d = data.sort_values("Date").reset_index(drop=True).copy()
    d["km_best"] = d["gps_km"].fillna(d["Jumlah_Pemakaian"])
    d["liters_refueled"] = d["Refuel_L"].fillna(0)

    total_liters = d["liters_refueled"].sum()
    total_km = d["km_best"].sum()
    avg_kml = (total_km / total_liters) if total_liters > 0 else None

    d["cum_liters"] = d["liters_refueled"].cumsum()
    d["cum_km"] = d["km_best"].cumsum()
    d["cum_liters_at_target"] = (d["cum_km"] / target_kml) if target_kml else None
    d["gap_l"] = d["cum_liters"] - d["cum_liters_at_target"] if target_kml else None

    return {
        "total_liters": total_liters,
        "total_km": total_km,
        "avg_kml": avg_kml,
        "daily": d[["Date", "liters_refueled", "km_best", "cum_liters", "cum_km", "cum_liters_at_target", "gap_l"]],
    }


def compute_idle_events(raw_df, threshold_minutes=10, speed_threshold=1.0):
    """
    Detect periods where the engine is ON (ACC=='ON') but the vehicle isn't
    moving (Speed <= speed_threshold) for at least threshold_minutes.
    Returns one row per qualifying idle event with start/end time and duration.
    """
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(columns=["date", "start", "end", "duration_min"])

    d = raw_df.sort_values("GPS Time").reset_index(drop=True)
    is_idle_ping = (d["ACC"] == "ON") & (d["Speed (Km/hr)"] <= speed_threshold)

    events = []
    start_idx = None
    for i, idle in enumerate(is_idle_ping):
        if idle and start_idx is None:
            start_idx = i
        elif not idle and start_idx is not None:
            start_t, end_t = d["GPS Time"].iloc[start_idx], d["GPS Time"].iloc[i - 1]
            dur_min = (end_t - start_t).total_seconds() / 60
            if dur_min >= threshold_minutes:
                events.append({"date": start_t.date(), "start": start_t, "end": end_t, "duration_min": dur_min})
            start_idx = None
    if start_idx is not None:  # trailing idle run through end of data
        start_t, end_t = d["GPS Time"].iloc[start_idx], d["GPS Time"].iloc[len(d) - 1]
        dur_min = (end_t - start_t).total_seconds() / 60
        if dur_min >= threshold_minutes:
            events.append({"date": start_t.date(), "start": start_t, "end": end_t, "duration_min": dur_min})

    return pd.DataFrame(events)


def compute_savings_estimate(target_kml, idle_minutes_total, idle_rate_lph=0.6, fuel_price=None, variance_km=None):
    """
    Estimate recoverable fuel from two sources, both independent of refuel-cycle data:
      1. Distance gap: if the logbook claims more km than GPS shows (variance_km > 0),
         that overstated distance — if used to plan or justify fuel — represents
         potentially over-claimed fuel, estimated at the target km/L. Only counted
         when the logbook overstates (GPS > logbook isn't converted to "savings").
      2. Idle waste: total idle time (engine on, not moving) x an assumed idle
         fuel-burn rate (liters/hour, editable — default 0.6 L/h is a common
         rule-of-thumb for a light diesel vehicle at idle).
    Returns a dict of both components, their sum, and (if a fuel price is given)
    the estimated cost of each.
    """
    result = {"distance_gap_l": 0.0, "idle_waste_l": 0.0, "total_l": 0.0}

    if variance_km is not None and target_kml:
        result["distance_gap_l"] = max(variance_km, 0.0) / target_kml

    result["idle_waste_l"] = (idle_minutes_total / 60.0) * idle_rate_lph
    result["total_l"] = result["distance_gap_l"] + result["idle_waste_l"]

    if fuel_price:
        result["distance_gap_cost"] = result["distance_gap_l"] * fuel_price
        result["idle_waste_cost"] = result["idle_waste_l"] * fuel_price
        result["total_cost"] = result["total_l"] * fuel_price

    return result


def build_insight(data, total_lb, total_gps, overall_var, flagged):
    """Auto-generate a short plain-language summary of the pattern."""
    n_days = len(data)
    if total_gps is None:
        return "No GPS data matched for this vehicle yet — upload a GPS export to see the reconciliation."

    lines = [
        f"Over <b>{n_days} days</b>, the logbook reports <b>{total_lb:,.0f} km</b> vs. "
        f"<b>{total_gps:,.0f} km</b> from GPS — an overall gap of <b>{overall_var*100:,.1f}%</b>."
    ]

    active = data[data["Jumlah_Pemakaian"] > 0]
    gap_dir = (active["variance_km"] > 0).sum()
    if len(active) > 0:
        if gap_dir == len(active):
            lines.append("GPS reads <b>lower</b> than the logbook on every active day — a consistent, one-directional pattern.")
        elif gap_dir == 0:
            lines.append("GPS reads <b>higher</b> than the logbook on every active day.")

    standby_issue = data[data["flag"] == "STANDBY"]
    no_logbook = data[data["flag"] == "NO-LOGBOOK"]

    if len(standby_issue):
        lines.append(f"⚠️ <b>{len(standby_issue)} 'standby' day(s)</b> show real GPS movement — worth checking for unlogged trips.")
    if len(no_logbook):
        km_missed = no_logbook["gps_km"].sum()
        lines.append(f"👻 <b>{len(no_logbook)} day(s)</b> show GPS movement ({km_missed:,.0f} km) with <b>no logbook entry at all</b> — more concerning than a standby mismatch, since the trip wasn't logged in any form.")
    if not len(standby_issue) and not len(no_logbook):
        lines.append("No unlogged movement detected — days marked idle (or missing from the logbook) show no meaningful GPS activity.")

    high = data[data["flag"] == "HIGH"]
    if len(high):
        worst = high.loc[high["variance_km"].abs().idxmax()]
        lines.append(
            f"<b>{len(high)} day(s)</b> exceed the 10% variance threshold — biggest gap on "
            f"<b>{worst['Date']}</b> ({worst['Jumlah_Pemakaian']:.0f} km logged vs {worst['gps_km']:.1f} km GPS)."
        )

    return " ".join(lines)


# ---------------------------------------------------------------------------
# Report generation (PDF) — print-friendly light theme, separate from the
# dark on-screen dashboard.
# ---------------------------------------------------------------------------
RPT_LB = ACCENT      # logbook (matches on-screen brand accent)
RPT_GPS = CYAN        # gps (matches on-screen brand accent)
RPT_GRID = "#dddddd"
RPT_TEXT = "#222222"


def strip_html(s):
    return re.sub(r"<[^>]+>", "", s)


_EMOJI_RE = re.compile(
    "["
    "\U0001F300-\U0001FAFF"  # symbols & pictographs, emoticons, supplemental symbols
    "\U00002600-\U000027BF"  # misc symbols & dingbats (includes ✅⚠️)
    "\U0000FE0F"             # variation selector (emoji presentation)
    "]+",
    flags=re.UNICODE,
)


def pdf_safe(s):
    """Strip HTML tags, replace common Unicode punctuation with ASCII equivalents,
    drop emoji, and fall back to dropping anything else non-Latin-1 (the core PDF
    font only supports Latin-1) so report generation never crashes on stray characters."""
    s = strip_html(str(s))
    s = _EMOJI_RE.sub("", s)
    replacements = {
        "—": "-", "–": "-", "→": "->", "\u2019": "'", "\u2018": "'",
        "\u201c": '"', "\u201d": '"', "±": "+/-", "×": "x",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    s = s.encode("latin-1", "ignore").decode("latin-1")
    return re.sub(r"\s{2,}", " ", s).strip()


def _fig_to_png_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def render_daily_chart_png(dates_str, lb_vals, gps_vals):
    fig, ax = plt.subplots(figsize=(7.5, 3.2))
    x = range(len(dates_str))
    w = 0.38
    bars_lb = ax.bar([i - w / 2 for i in x], lb_vals, width=w, color=RPT_LB, label="Logbook KM")
    bars_gps = ax.bar([i + w / 2 for i in x], gps_vals, width=w, color=RPT_GPS, label="GPS KM")
    ax.bar_label(bars_lb, fmt="%.0f", fontsize=6.5, padding=1)
    ax.bar_label(bars_gps, fmt=lambda v: f"{v:.0f}" if v == v else "", fontsize=6.5, padding=1)  # v==v filters NaN
    ax.set_xticks(list(x))
    ax.set_xticklabels(dates_str, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("KM")
    ax.margins(y=0.15)
    ax.grid(axis="y", color=RPT_GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    ax.legend(frameon=False, fontsize=8, loc="upper right")
    fig.tight_layout()
    return _fig_to_png_bytes(fig)


def render_cumulative_chart_png(dates_str, cum_lb, cum_gps):
    fig, ax = plt.subplots(figsize=(7.5, 3.0))
    ax.plot(dates_str, cum_lb, color=RPT_LB, marker="o", markersize=3, linewidth=2, label="Logbook (cumulative)")
    ax.plot(dates_str, cum_gps, color=RPT_GPS, marker="o", markersize=3, linewidth=2, label="GPS (cumulative)")
    for i, v in enumerate(cum_lb):
        ax.annotate(f"{v:.0f}", (i, v), textcoords="offset points", xytext=(0, 5), fontsize=6, ha="center", color=RPT_LB)
    for i, v in enumerate(cum_gps):
        ax.annotate(f"{v:.0f}", (i, v), textcoords="offset points", xytext=(0, -9), fontsize=6, ha="center", color=RPT_GPS)
    ax.set_xticks(range(len(dates_str)))
    ax.set_xticklabels(dates_str, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Cumulative KM")
    ax.margins(y=0.15)
    ax.grid(color=RPT_GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    ax.legend(frameon=False, fontsize=8, loc="upper left")
    fig.tight_layout()
    return _fig_to_png_bytes(fig)


def render_fuel_gap_chart_png(dates_str, daily_gap_l, cum_gap_l):
    fig, ax = plt.subplots(figsize=(7.5, 3.0))
    bars = ax.bar(range(len(dates_str)), daily_gap_l, color=RPT_LB, label="Daily fuel-equivalent gap (L)")
    ax.bar_label(bars, fmt=lambda v: f"{v:.1f}" if v == v and v != 0 else "", fontsize=6.5, padding=1)
    ax2 = ax.twinx()
    ax2.plot(range(len(dates_str)), cum_gap_l, color=RPT_GPS, marker="o", markersize=3, linewidth=2,
              linestyle="--", label="Cumulative (L)")
    ax.set_xticks(range(len(dates_str)))
    ax.set_xticklabels(dates_str, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Daily fuel-equivalent (L)")
    ax2.set_ylabel("Cumulative (L)")
    ax.margins(y=0.15)
    ax.grid(axis="y", color=RPT_GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    for spine in ["top"]:
        ax.spines[spine].set_visible(False)
        ax2.spines[spine].set_visible(False)
    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, frameon=False, fontsize=8, loc="upper left")
    fig.tight_layout()
    return _fig_to_png_bytes(fig)


def render_flag_donut_png(flag_counts):
    """flag_counts: dict like {'OK': 10, 'HIGH': 3, ...}"""
    colors_map = {"OK": GREEN, "HIGH": AMBER, "STANDBY": RED, "NO-GPS": DIM, "NO-LOGBOOK": "#9333EA"}
    labels = list(flag_counts.keys())
    values = list(flag_counts.values())
    fig, ax = plt.subplots(figsize=(4.2, 3.4))
    wedges, _ = ax.pie(
        values, colors=[colors_map.get(k, RPT_GRID) for k in labels],
        startangle=90, wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
    )
    ax.legend(wedges, [f"{k} ({v})" for k, v in zip(labels, values)],
              loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False, fontsize=8)
    ax.text(0, 0, f"{sum(values)}\ndays", ha="center", va="center", fontsize=12, color=RPT_TEXT)
    fig.tight_layout()
    return _fig_to_png_bytes(fig)


def render_idle_daily_chart_png(dates_str, idle_minutes):
    fig, ax = plt.subplots(figsize=(7.5, 2.8))
    bars = ax.bar(range(len(dates_str)), idle_minutes, color=AMBER)
    ax.bar_label(bars, fmt="%.0f", fontsize=7, padding=2)
    ax.set_xticks(range(len(dates_str)))
    ax.set_xticklabels(dates_str, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Idle minutes")
    ax.margins(y=0.15)
    ax.grid(axis="y", color=RPT_GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    return _fig_to_png_bytes(fig)


def render_fleet_bar_png(vehicle_names, values_a, values_b=None, label_a="", label_b="", ylabel=""):
    fig, ax = plt.subplots(figsize=(7.5, 3.2))
    x = range(len(vehicle_names))
    if values_b is not None:
        w = 0.38
        bars_a = ax.bar([i - w/2 for i in x], values_a, width=w, color=RPT_LB, label=label_a)
        bars_b = ax.bar([i + w/2 for i in x], values_b, width=w, color=RPT_GPS, label=label_b)
        ax.bar_label(bars_a, fmt="%.0f", fontsize=6.5, padding=1)
        ax.bar_label(bars_b, fmt=lambda v: f"{v:.0f}" if v == v else "", fontsize=6.5, padding=1)
        ax.legend(frameon=False, fontsize=8, loc="upper right")
    else:
        bars = ax.bar(list(x), values_a, color=AMBER)
        ax.bar_label(bars, fmt="%.0f", fontsize=7, padding=2)
    ax.set_xticks(list(x))
    ax.set_xticklabels(vehicle_names, fontsize=8, rotation=15, ha="right")
    ax.set_ylabel(ylabel)
    ax.margins(y=0.15)
    ax.grid(axis="y", color=RPT_GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
    fig.tight_layout()
    return _fig_to_png_bytes(fig)


class ReportPDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._tmp_image_paths = []

    def header(self):
        if getattr(self, "suppress_header", False):
            return
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(140, 140, 140)
        self.cell(0, 8, "Driver Logbook <-> GPS Tracker Report", align="L")
        self.cell(0, 8, dt.datetime.now().strftime("Generated %Y-%m-%d %H:%M"), align="R", ln=True)
        self.set_draw_color(220, 220, 220)
        self.line(10, 18, 200, 18)
        self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")

    def section_title(self, text):
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(30, 30, 30)
        self.cell(0, 9, pdf_safe(text), ln=True)
        self.set_draw_color(230, 230, 230)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(3)

    def kpi_row(self, kpis):
        self.set_font("Helvetica", "", 9)
        n = len(kpis)
        col_w = 190 / n
        y0 = self.get_y()
        for label, value in kpis:
            x = self.get_x()
            self.set_fill_color(245, 246, 247)
            self.rect(x, y0, col_w - 4, 16, style="F")
            self.set_xy(x + 2, y0 + 2)
            self.set_font("Helvetica", "", 7.5)
            self.set_text_color(120, 120, 120)
            self.cell(col_w - 6, 4, pdf_safe(label))
            self.set_xy(x + 2, y0 + 7)
            self.set_font("Helvetica", "B", 11)
            self.set_text_color(30, 30, 30)
            self.cell(col_w - 6, 6, pdf_safe(str(value)))
            self.set_xy(x + col_w - 2, y0)
        self.set_xy(10, y0 + 20)

    def body_text(self, text):
        self.set_font("Helvetica", "", 9.5)
        self.set_text_color(50, 50, 50)
        self.multi_cell(0, 5, pdf_safe(text))
        self.ln(2)

    def image_from_bytes(self, png_bytes, w=190):
        # Write to a temp file and pass a plain path string — the most
        # version-independent input fpdf2 accepts (avoids BytesIO/PIL-Image
        # support differences and object-hashing quirks across versions).
        # Cleanup is deferred until after pdf.output() (see cleanup_temp_files),
        # in case the library reads image data lazily rather than immediately.
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        tmp.write(png_bytes)
        tmp.close()
        self._tmp_image_paths.append(tmp.name)
        self.image(tmp.name, w=w)
        self.ln(2)

    def cleanup_temp_files(self):
        for p in self._tmp_image_paths:
            try:
                os.remove(p)
            except OSError:
                pass

    def data_table(self, headers, rows, col_widths=None, font_size=7.5):
        if col_widths is None:
            col_widths = [190 / len(headers)] * len(headers)
        self.set_font("Helvetica", "B", font_size)
        self.set_fill_color(235, 237, 239)
        self.set_text_color(40, 40, 40)
        for h, w in zip(headers, col_widths):
            self.cell(w, 6, pdf_safe(str(h)), border=1, fill=True, align="C")
        self.ln()
        self.set_font("Helvetica", "", font_size)
        fill = False
        for row in rows:
            if self.get_y() > 270:
                self.add_page()
            self.set_fill_color(250, 250, 250)
            for val, w in zip(row, col_widths):
                self.cell(w, 5.5, pdf_safe(str(val)), border=1, fill=fill, align="C")
            self.ln()
            fill = not fill


def build_vehicle_report_section(pdf: "ReportPDF", plate, data, target_kml, tolerance_pct,
                                  raw_df=None, idle_rate_lph=0.6, fuel_price=None):
    dates_str = data["Date"].astype(str)
    total_lb = data["Jumlah_Pemakaian"].sum()
    has_gps_rows = data["gps_km"].notna()
    total_gps = data.loc[has_gps_rows, "gps_km"].sum() if has_gps_rows.any() else None
    overall_var = (total_lb - total_gps) / total_lb if total_gps is not None and total_lb > 0 else None
    flagged = data["flag"].isin(["HIGH", "STANDBY", "NO-LOGBOOK"]).sum()
    max_speed = data["max_speed"].max() if data["max_speed"].notna().any() else None

    pdf.section_title(f"Vehicle: {plate}")
    pdf.kpi_row([
        ("LOGBOOK TOTAL", f"{total_lb:,.0f} km"),
        ("GPS TOTAL", f"{total_gps:,.0f} km" if total_gps is not None else "-"),
        ("VARIANCE", f"{overall_var*100:,.1f}%" if overall_var is not None else "-"),
        ("FLAGGED DAYS", f"{flagged} / {len(data)}"),
        ("MAX SPEED", f"{max_speed:,.0f} km/h" if max_speed is not None else "-"),
    ])

    pdf.body_text(strip_html(build_insight(data, total_lb, total_gps, overall_var, flagged)))

    if total_gps is not None:
        png = render_daily_chart_png(dates_str.str[5:], data["Jumlah_Pemakaian"], data["gps_km"])
        pdf.image_from_bytes(png)

        cum_lb = data["Jumlah_Pemakaian"].cumsum()
        cum_gps = data["gps_km"].fillna(0).cumsum()
        png2 = render_cumulative_chart_png(dates_str.str[5:], cum_lb, cum_gps)
        pdf.image_from_bytes(png2)

    pdf.add_page()
    pdf.section_title(f"{plate} — Day-by-Day Detail")
    rows = []
    for _, r in data.iterrows():
        rows.append([
            str(r["Date"]),
            f"{r['Jumlah_Pemakaian']:.0f}" if pd.notna(r["Jumlah_Pemakaian"]) else "-",
            f"{r['gps_km']:.1f}" if pd.notna(r["gps_km"]) else "-",
            f"{abs(r['variance_pct'])*100:.0f}%" if pd.notna(r["variance_pct"]) else "-",
            f"{r['max_speed']:.0f}" if pd.notna(r["max_speed"]) else "-",
            r["flag"],
        ])
    pdf.data_table(
        ["Date", "Logbook KM", "GPS KM", "Var %", "Max Spd", "Flag"],
        rows, col_widths=[30, 30, 28, 25, 27, 50],
    )
    pdf.ln(2)
    pdf.body_text("Note: Var % is shown as an absolute value. Logbook KM x Var % = the gap between Logbook and GPS, in km. See the KM columns for which side is higher.")

    if total_gps is not None:
        pdf.add_page()
        pdf.section_title(f"{plate} — Fuel Analysis")
        km_gap = total_lb - total_gps
        est_savings_l = max(km_gap, 0) / target_kml if target_kml else 0.0
        total_refuel_l = data["Refuel_L"].fillna(0).sum()

        pdf.kpi_row([
            ("TOTAL REFUELED", f"{total_refuel_l:,.0f} L"),
            ("KM GAP", f"{km_gap:,.0f} km"),
            ("TARGET", f"{target_kml:.1f} km/L"),
            ("EST. FUEL IMPACT", f"{est_savings_l:,.0f} L"),
        ])

        fuel_insight = (
            f"Simple estimate: the logbook-vs-GPS KM gap is {km_gap:,.0f} km "
            f"({total_lb:,.0f} km logged vs {total_gps:,.0f} km GPS). At a target of {target_kml:.1f} km/L, "
            f"that gap is worth roughly {est_savings_l:,.0f} L of fuel"
            + (" - worth checking if that distance is being fueled but not actually driven."
               if km_gap > 0 else ", though GPS shows more distance than logged here, so there's no fuel implied by this gap.")
        )
        pdf.body_text(fuel_insight)

        daily_gap_km = (data["Jumlah_Pemakaian"] - data["gps_km"]).clip(lower=0)
        daily_gap_l = (daily_gap_km / target_kml) if target_kml else daily_gap_km * 0
        cum_gap_l = daily_gap_l.cumsum()
        dates_str_fuel = data["Date"].astype(str).str[5:]

        png3 = render_fuel_gap_chart_png(dates_str_fuel, daily_gap_l, cum_gap_l)
        pdf.image_from_bytes(png3)
        pdf.body_text("Estimate = max(Logbook KM - GPS KM, 0) / Target km/L. No refuel-cycle assumptions, no actual liters tracked - just what the distance gap would cost in fuel at the target efficiency.")

    idle_events = compute_idle_events(raw_df, 10, 1.0) if raw_df is not None else pd.DataFrame()
    variance_km_val = (total_lb - total_gps) if total_gps is not None else None
    savings = compute_savings_estimate(
        target_kml,
        idle_events["duration_min"].sum() if not idle_events.empty else 0.0,
        idle_rate_lph, fuel_price, variance_km_val,
    )

    pdf.add_page()
    pdf.section_title(f"{plate} — Idle Time & Savings Estimate")

    if idle_events.empty:
        pdf.body_text(f"No idle events of 10+ minutes detected for this vehicle over the period.")
    else:
        total_idle_min = idle_events["duration_min"].sum()
        longest = idle_events.loc[idle_events["duration_min"].idxmax()]
        days_with_idle = idle_events["date"].nunique()

        pdf.body_text(
            f"{len(idle_events)} idle event(s) of 10+ minutes across {days_with_idle} day(s), totaling "
            f"{total_idle_min/60:,.1f} hours of engine-on, not-moving time. The longest single event was "
            f"{longest['duration_min']:.0f} minutes on {longest['date']} (starting {longest['start'].strftime('%H:%M')})."
        )

        pdf.kpi_row([
            ("TOTAL IDLE TIME", f"{total_idle_min/60:,.1f} hrs"),
            ("IDLE EVENTS", f"{len(idle_events)}"),
            ("DAYS AFFECTED", f"{days_with_idle}"),
            ("LONGEST EVENT", f"{longest['duration_min']:.0f} min"),
        ])

        daily_idle = idle_events.groupby("date")["duration_min"].sum().reset_index()
        png_idle = render_idle_daily_chart_png(daily_idle["date"].astype(str).str[5:], daily_idle["duration_min"])
        pdf.image_from_bytes(png_idle)

        rows = [[str(r["date"]), r["start"].strftime("%H:%M"), r["end"].strftime("%H:%M"), f"{r['duration_min']:.0f}"]
                for _, r in idle_events.iterrows()]
        pdf.data_table(["Date", "Start", "End", "Duration (min)"], rows, col_widths=[45, 45, 45, 55])

    pdf.ln(3)
    distance_gap_desc = (
        f"({variance_km_val:,.0f} km / {target_kml:.1f} km/L target)" if variance_km_val is not None
        else "(no GPS data to compute a distance gap for this vehicle)"
    )
    savings_line = (
        f"Estimated recoverable fuel: {savings['distance_gap_l']:,.0f} L from the logbook-vs-GPS distance gap "
        f"{distance_gap_desc}, "
        f"plus {savings['idle_waste_l']:,.0f} L from idle time (at {idle_rate_lph:.1f} L/hour assumed idle burn), "
        f"for a total estimate of {savings['total_l']:,.0f} L."
    )
    if fuel_price:
        savings_line += f" At the given fuel price, that's approximately Rp.{savings.get('total_cost', 0):,.0f}."
    savings_line += " This is only estimation based on the assumptions above."
    pdf.body_text(savings_line)


def _pdf_to_bytes(pdf):
    """
    Modern fpdf2's output() takes no arguments and always returns a bytearray
    of the complete PDF. The old, unmaintained 'fpdf' package (which installs
    under the same 'fpdf' module name, so it's easy to end up with it instead)
    uses an older API where output() needs dest='S' to return the buffer as a
    string — without it, behavior varies by version and can produce something
    that isn't a valid, complete PDF. We detect which API is present and call
    it correctly, then verify the result actually looks like a real PDF before
    handing it back, so a bad install fails loudly here rather than producing
    a file that downloads but won't open.
    """
    import inspect

    try:
        try:
            sig = inspect.signature(pdf.output)
            if "dest" in sig.parameters:
                raw = pdf.output(dest="S")  # old-style 'fpdf' API
            else:
                raw = pdf.output()  # modern fpdf2
        except (TypeError, ValueError):
            raw = pdf.output()
    finally:
        pdf.cleanup_temp_files()

    data = raw.encode("latin-1") if isinstance(raw, str) else bytes(raw)

    if not data.startswith(b"%PDF-") or b"%%EOF" not in data[-2048:]:
        raise RuntimeError(
            "Generated PDF looks incomplete or invalid — this usually means the "
            "old 'fpdf' package is installed instead of 'fpdf2'. Please run:\n"
            "  pip uninstall fpdf\n"
            "  pip install fpdf2\n"
            "and try again."
        )
    return data


def build_single_vehicle_pdf(plate, data, target_kml, tolerance_pct,
                              raw_df=None, idle_rate_lph=0.6, fuel_price=None):
    pdf = ReportPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(20, 20, 20)
    pdf.cell(0, 12, pdf_safe(f"Reconciliation Report: {plate}"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 6, pdf_safe(f"Period: {data['Date'].min()} to {data['Date'].max()}"), ln=True)
    pdf.ln(4)

    build_vehicle_report_section(pdf, plate, data, target_kml, tolerance_pct,
                                  raw_df, idle_rate_lph, fuel_price)
    return _pdf_to_bytes(pdf)


def build_fleet_pdf(vehicle_payloads):
    """vehicle_payloads: list of dicts with keys plate, data, target_kml, tolerance_pct"""
    pdf = ReportPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=18)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(20, 20, 20)
    pdf.cell(0, 14, pdf_safe("Fleet-Wide Reconciliation Report"), ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 6, pdf_safe(f"{len(vehicle_payloads)} vehicle(s) included"), ln=True)
    pdf.ln(4)

    pdf.section_title("Fleet Overview")
    rows = []
    chart_names, chart_lb, chart_gps, chart_savings = [], [], [], []
    total_var_km_sum, total_idle_hrs_sum, total_savings_sum = 0.0, 0.0, 0.0
    for vp in vehicle_payloads:
        d = vp["data"]
        total_lb = d["Jumlah_Pemakaian"].sum()
        has_gps = d["gps_km"].notna()
        total_gps = d.loc[has_gps, "gps_km"].sum() if has_gps.any() else None
        var_km_val = (total_lb - total_gps) if total_gps is not None else None
        var = (var_km_val / total_lb) if var_km_val is not None and total_lb else None
        flagged = d["flag"].isin(["HIGH", "STANDBY", "NO-LOGBOOK"]).sum()
        fuel_overview = compute_fuel_overview(d, vp["target_kml"]) if d["Refuel_L"].fillna(0).sum() > 0 else None
        liters = fuel_overview["total_liters"] if fuel_overview else None

        idle_evts = compute_idle_events(vp.get("raw_df"), 10, 1.0) if vp.get("raw_df") is not None else pd.DataFrame()
        idle_hrs = idle_evts["duration_min"].sum() / 60.0 if not idle_evts.empty else 0.0
        savings = compute_savings_estimate(
            vp["target_kml"], idle_evts["duration_min"].sum() if not idle_evts.empty else 0.0,
            vp.get("idle_rate_lph", 0.6), vp.get("fuel_price"), var_km_val,
        )

        rows.append([
            vp["plate"],
            f"{total_lb:,.0f}",
            f"{total_gps:,.0f}" if total_gps is not None else "-",
            f"{var*100:.1f}%" if var is not None else "-",
            f"{flagged}/{len(d)}",
            f"{liters:,.0f}" if liters is not None else "-",
            f"{idle_hrs:.1f}",
            f"{savings['total_l']:.0f}",
        ])
        chart_names.append(vp["plate"])
        chart_lb.append(total_lb)
        chart_gps.append(total_gps if total_gps is not None else 0)
        chart_savings.append(savings["total_l"])
        total_var_km_sum += abs(var_km_val) if var_km_val is not None else 0
        total_idle_hrs_sum += idle_hrs
        total_savings_sum += savings["total_l"]

    pdf.data_table(
        ["Vehicle", "Logbook KM", "GPS KM", "Variance", "Flagged", "Fuel (L)", "Idle Hrs", "Est. Save (L)"],
        rows, col_widths=[32, 24, 22, 22, 20, 22, 20, 26],
    )

    pdf.ln(3)
    fuel_price_any = next((vp.get("fuel_price") for vp in vehicle_payloads if vp.get("fuel_price")), None)
    fleet_insight = (
        f"Across {len(vehicle_payloads)} vehicle(s), total logbook-vs-GPS variance is {total_var_km_sum:,.0f} km "
        f"and total idle time is {total_idle_hrs_sum:,.1f} hours. Estimated total recoverable fuel across the "
        f"fleet: {total_savings_sum:,.0f} L"
        + (f" (~{total_savings_sum*fuel_price_any:,.0f} at the given fuel price)" if fuel_price_any else "")
        + " - combining the fuel implied by the distance gap (variance km / target km/L) and idle time "
          "(idle hours x assumed burn rate). This is an estimate, not an audit finding; use it to prioritize "
          "which vehicles to look at first."
    )
    pdf.body_text(fleet_insight)

    png_var = render_fleet_bar_png(chart_names, chart_lb, chart_gps, "Logbook KM", "GPS KM", "KM")
    pdf.image_from_bytes(png_var)
    png_save = render_fleet_bar_png(chart_names, chart_savings, ylabel="Estimated liters")
    pdf.image_from_bytes(png_save)

    for vp in vehicle_payloads:
        pdf.add_page()
        build_vehicle_report_section(
            pdf, vp["plate"], vp["data"], vp["target_kml"], vp["tolerance_pct"],
            vp.get("raw_df"), vp.get("idle_rate_lph", 0.6), vp.get("fuel_price"),
        )

    return _pdf_to_bytes(pdf)


# ---------------------------------------------------------------------------
# Sidebar — data in
# ---------------------------------------------------------------------------
st.sidebar.markdown("### 📥 Data In")
logbook_file = st.sidebar.file_uploader("Logbook workbook (one sheet per vehicle)", type=["xlsx", "xls"])
gps_files = st.sidebar.file_uploader(
    "GPS history export(s) — select multiple", type=["xlsx", "xls"], accept_multiple_files=True
)

logbook_data = parse_logbook(logbook_file) if logbook_file else {}
gps_data = {}
gps_raw_data = {}
if gps_files:
    for f in gps_files:
        parsed, raw_parsed = parse_gps(f)
        gps_data.update(parsed)
        gps_raw_data.update(raw_parsed)

if logbook_file:
    st.sidebar.success(f"{len(logbook_data)} vehicle sheet(s) loaded", icon="📘")
if gps_files:
    st.sidebar.success(f"{len(gps_files)} GPS file(s) loaded", icon="🛰️")

vehicles = sorted(set(logbook_data.keys()) | set(gps_data.keys()))

st.title("🛰️ Driver Logbook ↔ GPS Tracker")
st.caption("Upload a driver logbook and one or more GPS history exports — vehicles are matched automatically by plate number.")

if not vehicles:
    st.info("⬅ Upload a logbook workbook and matching GPS export(s) in the sidebar to build the dashboard.")
    st.stop()

# ---------------------------------------------------------------------------
# Sidebar — vehicle picker
# ---------------------------------------------------------------------------
st.sidebar.markdown("### 🚚 Vehicles")
for v in vehicles:
    has_lb, has_gps = v in logbook_data, v in gps_data
    dot = "🟢" if (has_lb and has_gps) else ("🟡" if has_lb else "🔵")
    status = "matched" if (has_lb and has_gps) else ("no gps" if has_lb else "no logbook")
    st.sidebar.markdown(f"{dot} **{v}** — *{status}*")

plate = st.sidebar.selectbox("Select vehicle to inspect", vehicles, label_visibility="collapsed")

lb_df = logbook_data.get(plate)
gps_df = gps_data.get(plate)
gps_raw_df = gps_raw_data.get(plate)

if lb_df is None:
    st.warning(f"No logbook rows found for **{plate}** (GPS-only vehicle).")
    st.stop()

data = compute_reconciliation(lb_df, gps_df)
dates_str = data["Date"].astype(str).str[5:]

# ---------------------------------------------------------------------------
# KPI row
# ---------------------------------------------------------------------------
total_lb = data["Jumlah_Pemakaian"].sum()
has_gps_rows = data["gps_km"].notna()
total_gps = data.loc[has_gps_rows, "gps_km"].sum() if has_gps_rows.any() else None
overall_var = (total_lb - total_gps) / total_lb if total_gps is not None and total_lb > 0 else None
flagged = (data["flag"].isin(["HIGH", "STANDBY", "NO-LOGBOOK"])).sum()
max_speed = data["max_speed"].max() if data["max_speed"].notna().any() else None

st.subheader(f"📋 {plate}")

st.markdown(f"<div class='insight-box'>💡 {build_insight(data, total_lb, total_gps, overall_var, flagged)}</div>", unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("📘 Logbook Total", f"{total_lb:,.0f} km")
c2.metric("🛰️ GPS Total", f"{total_gps:,.0f} km" if total_gps is not None else "—")
c3.metric(
    "📊 Overall Variance", f"{overall_var*100:,.1f}%" if overall_var is not None else "—",
    delta="within tolerance" if overall_var is not None and abs(overall_var) <= 0.10 else ("check needed" if overall_var is not None else None),
    delta_color="normal" if overall_var is not None and abs(overall_var) <= 0.10 else "inverse",
)
c4.metric(
    "🚩 Flagged Days", f"{flagged} / {len(data)}",
    delta="clean" if flagged == 0 else f"{flagged} to review",
    delta_color="normal" if flagged == 0 else "inverse",
)
c5.metric("⚡ Max Speed", f"{max_speed:,.0f} km/h" if max_speed is not None else "—")

st.write("")

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
    ["📈 Daily Pattern", "📉 Cumulative Trend", "📋 Detail Table", "⛽ Fuel Analysis",
     "⏱️ Idle Time", "🚚 Fleet Overview", "📄 Reports"]
)

with tab1:
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("**Logbook vs. GPS distance per day**")
        st.caption("Bars show the exact KM reported by each source; amber line = variance % (right axis), drawn continuously through standby days. Flagged days are marked ⚠️ / 🚩 above the bars.")

        fig1 = go.Figure()
        fig1.add_bar(
            x=dates_str, y=data["Jumlah_Pemakaian"], name="Logbook KM", marker_color=ACCENT,
            text=data["Jumlah_Pemakaian"].apply(lambda v: f"{v:.0f}" if pd.notna(v) else ""), textposition="outside",
            textfont=dict(size=11, color=TEXT),
            hovertemplate="Logbook: %{y:.0f} km<extra></extra>",
        )
        fig1.add_bar(
            x=dates_str, y=data["gps_km"], name="GPS KM", marker_color=CYAN,
            text=data["gps_km"].apply(lambda v: f"{v:.0f}" if pd.notna(v) else ""), textposition="outside",
            textfont=dict(size=11, color=TEXT),
            hovertemplate="GPS: %{y:.1f} km<extra></extra>",
        )

        var_pct_display = data["variance_pct"].apply(lambda v: abs(v) * 100 if pd.notna(v) else None)
        fig1.add_trace(go.Scatter(
            x=dates_str, y=var_pct_display, name="Variance %", mode="lines+markers",
            line=dict(color=AMBER, width=2), marker=dict(size=6),
            yaxis="y2", connectgaps=True,
            hovertemplate="Variance: %{y:.1f}%<extra></extra>",
        ))

        # Annotate flagged days above their bars
        top_val = data[["Jumlah_Pemakaian", "gps_km"]].max(axis=1).fillna(data["Jumlah_Pemakaian"])
        for i, row in data.iterrows():
            if row["flag"] in ("HIGH", "STANDBY", "NO-LOGBOOK"):
                icon = FLAG_ICONS[row["flag"]]
                fig1.add_annotation(
                    x=dates_str.iloc[i], y=top_val.iloc[i], text=icon,
                    showarrow=False, yshift=26, font=dict(size=14),
                )

        fig1.update_layout(
            barmode="group", template="plotly_white",
            paper_bgcolor=PANEL, plot_bgcolor=PANEL,
            legend=dict(orientation="h", y=1.15, font=dict(color=DIM)),
            margin=dict(l=10, r=10, t=30, b=10), height=420,
            xaxis=dict(gridcolor=GRID, type="category"), yaxis=dict(title="KM", gridcolor=GRID),
            yaxis2=dict(title="Variance % (abs)", overlaying="y", side="right", ticksuffix="%", showgrid=False, rangemode="tozero"),
            hovermode="x unified",
        )
        st.plotly_chart(fig1, width="stretch")
        st.caption("Note: Logbook KM x Variance % = the gap between Logbook and GPS, in km. To see which side is higher on a given day, compare the two bars directly.")

    with col2:
        st.markdown("**Flag distribution**")
        st.caption("Share of days by reconciliation outcome")
        counts = data["flag"].value_counts()
        fig2 = go.Figure(go.Pie(
            labels=[f"{FLAG_ICONS.get(k,'')} {k}" for k in counts.index], values=counts.values, hole=0.62,
            marker=dict(colors=[FLAG_COLORS.get(k, DIM) for k in counts.index], line=dict(color=PANEL, width=3)),
            textinfo="value+percent", textfont=dict(color=TEXT, size=12),
        ))
        fig2.add_annotation(text=f"{len(data)}<br>days", showarrow=False, font=dict(size=18, color=TEXT))
        fig2.update_layout(
            template="plotly_white", paper_bgcolor=PANEL,
            margin=dict(l=10, r=10, t=30, b=10), height=420,
            legend=dict(font=dict(color=DIM)),
        )
        st.plotly_chart(fig2, width="stretch")

with tab2:
    st.markdown("**Logbook vs. GPS cumulative distance + daily gap**")
    st.caption("Solid lines = running total (left axis); dashed amber line = that day's Logbook−GPS gap in KM (right axis)")

    cum_lb = data["Jumlah_Pemakaian"].cumsum()
    cum_gps = data["gps_km"].fillna(0).cumsum()
    daily_gap = data["variance_km"]

    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(x=dates_str, y=cum_lb, name="Logbook (cumulative KM)",
                               mode="lines+markers", line=dict(color=ACCENT, width=2),
                               fill="tozeroy", fillcolor="rgba(255,122,51,0.08)"))
    fig3.add_trace(go.Scatter(x=dates_str, y=cum_gps, name="GPS (cumulative KM)",
                               mode="lines+markers", line=dict(color=CYAN, width=2),
                               fill="tozeroy", fillcolor="rgba(53,194,193,0.08)"))
    fig3.add_trace(go.Scatter(x=dates_str, y=daily_gap, name="Daily Gap (KM)",
                               mode="lines+markers", line=dict(color=AMBER, width=2, dash="dash"),
                               yaxis="y2", connectgaps=True))

    fig3.update_layout(
        template="plotly_white", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
        legend=dict(orientation="h", y=1.12, font=dict(color=DIM)),
        margin=dict(l=10, r=10, t=20, b=10), height=440,
        xaxis=dict(gridcolor=GRID, type="category"), yaxis=dict(title="Cumulative KM", gridcolor=GRID),
        yaxis2=dict(title="Daily Gap (KM)", overlaying="y", side="right", showgrid=False),
        hovermode="x unified",
    )
    st.plotly_chart(fig3, width="stretch")

    final_gap = daily_gap.sum()
    st.caption(f"Final cumulative gap after {len(data)} days: **{final_gap:,.1f} km** "
               f"({'logbook ahead' if final_gap > 0 else 'GPS ahead'}).")

with tab3:
    st.markdown("**Day-by-day detail**")
    only_flagged = st.checkbox("Show only flagged days", value=False)

    table_df = data.copy()
    if only_flagged:
        table_df = table_df[table_df["flag"].isin(["HIGH", "STANDBY", "NO-LOGBOOK"])]

    display_df = table_df[["Date", "Tujuan", "Jumlah_Pemakaian", "gps_km", "variance_pct", "max_speed", "acc_on", "flag"]].copy()
    display_df.columns = ["Date", "Route", "Logbook KM", "GPS KM", "Variance", "Max Speed", "ACC-On Pings", "Flag"]
    display_df["Variance"] = pd.to_numeric(display_df["Variance"], errors="coerce").abs() * 100  # absolute value, as % points
    display_df["Flag"] = display_df["Flag"].apply(lambda f: f"{FLAG_ICONS.get(f,'')} {f}")

    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,
        column_config={
            "Logbook KM": st.column_config.NumberColumn(format="%.0f km"),
            "GPS KM": st.column_config.NumberColumn(format="%.1f km"),
            "Variance": st.column_config.ProgressColumn(
                format="%.0f%%", min_value=0.0, max_value=100.0,
                help="|Logbook−GPS| as a share of logbook KM (always shown positive)",
            ),
            "Max Speed": st.column_config.NumberColumn(format="%.0f km/h"),
            "ACC-On Pings": st.column_config.NumberColumn(format="%d"),
        },
    )
    st.caption(f"Showing {len(display_df)} of {len(data)} days. Logbook KM x Variance % = the gap between Logbook and GPS, in km. Compare the KM columns directly to see which side is higher.")

with tab4:
    total_lb = data["Jumlah_Pemakaian"].sum()
    has_gps_rows = data["gps_km"].notna()
    total_gps = data.loc[has_gps_rows, "gps_km"].sum() if has_gps_rows.any() else None

    if total_gps is None:
        st.info("No GPS data matched for this vehicle yet — the fuel estimate needs a KM gap (Logbook vs GPS) to work from.")
    else:
        st.markdown("**Set a target fuel efficiency for this vehicle**")
        st.caption("Defaults to 9.0 km/L for every vehicle — override it with the manufacturer-rated or fleet-standard km/L for this vehicle class if you have one.")

        target_key = f"target_kml_{plate}"
        if target_key not in st.session_state:
            st.session_state[target_key] = 9.0
        target_kml = st.number_input("Target km/L", min_value=1.0, max_value=50.0, step=0.5, key=target_key)

        km_gap = total_lb - total_gps
        est_savings_l = max(km_gap, 0) / target_kml if target_kml else 0.0

        st.markdown(
            f"<div class='insight-box'>⛽ Simple estimate: the logbook-vs-GPS <b>KM gap</b> is <b>{km_gap:,.0f} km</b> "
            f"({total_lb:,.0f} km logged vs {total_gps:,.0f} km GPS). At a target of <b>{target_kml:.1f} km/L</b>, "
            f"that gap is worth roughly <b>{est_savings_l:,.0f} L</b> of fuel"
            + (" — worth checking if that distance is being fueled but not actually driven." if km_gap > 0 else ", though GPS shows more distance than logged here, so there's no fuel implied by this gap.")
            + "</div>",
            unsafe_allow_html=True,
        )
        st.caption("Estimate = max(Logbook KM − GPS KM, 0) ÷ Target km/L. Simple and transparent — no refuel-cycle assumptions, no actual liters tracked, just what the distance gap would cost in fuel at your target efficiency.")

        total_refuel_l = data["Refuel_L"].fillna(0).sum()

        fc1, fc2, fc3, fc4 = st.columns(4)
        fc1.metric("⛽ Total Refueled", f"{total_refuel_l:,.0f} L")
        fc2.metric("📏 KM Gap", f"{km_gap:,.0f} km")
        fc3.metric("🎯 Target", f"{target_kml:.1f} km/L")
        fc4.metric("💧 Est. Fuel Impact", f"{est_savings_l:,.0f} L")

        st.write("")
        st.markdown("**Daily KM gap → fuel-equivalent**")
        st.caption("Bars = each day's KM gap converted to liters at your target km/L; dashed line = the running cumulative total.")

        daily_gap_km = (data["Jumlah_Pemakaian"] - data["gps_km"]).clip(lower=0)
        daily_gap_l = daily_gap_km / target_kml if target_kml else daily_gap_km * 0
        cum_gap_l = daily_gap_l.cumsum()
        dates_str_fuel = data["Date"].astype(str).str[5:]

        fig4 = go.Figure()
        fig4.add_bar(
            x=dates_str_fuel, y=daily_gap_l, name="Daily fuel-equivalent gap", marker_color=ACCENT,
            text=daily_gap_l.apply(lambda v: f"{v:.1f}" if pd.notna(v) else ""), textposition="outside",
            textfont=dict(size=11, color=TEXT),
        )
        fig4.add_trace(go.Scatter(x=dates_str_fuel, y=cum_gap_l, name="Cumulative (L)",
                                   mode="lines+markers", line=dict(color=CYAN, width=2, dash="dash"),
                                   yaxis="y2"))
        fig4.update_layout(
            template="plotly_white", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
            legend=dict(orientation="h", y=1.15, font=dict(color=DIM)),
            margin=dict(l=10, r=10, t=30, b=10), height=420,
            xaxis=dict(gridcolor=GRID, type="category"), yaxis=dict(title="Daily fuel-equivalent (L)", gridcolor=GRID),
            yaxis2=dict(title="Cumulative (L)", overlaying="y", side="right", showgrid=False),
            hovermode="x unified",
        )
        st.plotly_chart(fig4, width="stretch")

        st.write("")
        st.markdown("**Daily detail**")
        fuel_display = pd.DataFrame({
            "Date": data["Date"], "Logbook KM": data["Jumlah_Pemakaian"], "GPS KM": data["gps_km"],
            "KM Gap": daily_gap_km, "Fuel-Equivalent (L)": daily_gap_l, "Cumulative (L)": cum_gap_l,
        })
        st.dataframe(
            fuel_display, width="stretch", hide_index=True,
            column_config={
                "Logbook KM": st.column_config.NumberColumn(format="%.0f km"),
                "GPS KM": st.column_config.NumberColumn(format="%.1f km"),
                "KM Gap": st.column_config.NumberColumn(format="%.1f km"),
                "Fuel-Equivalent (L)": st.column_config.NumberColumn(format="%.1f L"),
                "Cumulative (L)": st.column_config.NumberColumn(format="%.1f L"),
            },
        )

with tab5:
    st.markdown("**Idle time detection**")
    st.caption("Periods where the engine was ON but the vehicle wasn't moving, for at least the threshold below. Long idling burns fuel with zero distance to show for it.")

    idle_col1, idle_col2 = st.columns([1, 3])
    idle_threshold_key = f"idle_threshold_{plate}"
    idle_speed_key = f"idle_speed_{plate}"
    if idle_threshold_key not in st.session_state:
        st.session_state[idle_threshold_key] = 10
    if idle_speed_key not in st.session_state:
        st.session_state[idle_speed_key] = 1.0

    idle_threshold = idle_col1.number_input(
        "Idle threshold (minutes)", min_value=1, max_value=120, step=1, key=idle_threshold_key,
        help="Minimum continuous ACC-ON, not-moving duration to count as an idle event",
    )
    idle_speed_thresh = idle_col2.slider(
        "\"Not moving\" speed cutoff (km/h)", min_value=0.0, max_value=5.0, step=0.5, key=idle_speed_key,
        help="Speed at or below this counts as stationary (accounts for GPS jitter)",
    )

    if gps_raw_df is None:
        st.info("No raw GPS ping data available for this vehicle — idle detection needs a matched GPS export.")
        idle_events = pd.DataFrame()
    else:
        idle_events = compute_idle_events(gps_raw_df, idle_threshold, idle_speed_thresh)

        if idle_events.empty:
            st.markdown(f"<div class='insight-box'>✅ No idle events of {idle_threshold}+ minutes detected for this vehicle over the period.</div>", unsafe_allow_html=True)
        else:
            total_idle_min = idle_events["duration_min"].sum()
            longest = idle_events.loc[idle_events["duration_min"].idxmax()]
            days_with_idle = idle_events["date"].nunique()

            st.markdown(
                f"<div class='insight-box'>⏱️ <b>{len(idle_events)} idle event(s)</b> of {idle_threshold}+ minutes "
                f"across <b>{days_with_idle} day(s)</b>, totaling <b>{total_idle_min/60:,.1f} hours</b> of engine-on, "
                f"not-moving time. The longest single event was <b>{longest['duration_min']:.0f} minutes</b> on "
                f"<b>{longest['date']}</b> (starting {longest['start'].strftime('%H:%M')}).</div>",
                unsafe_allow_html=True,
            )

            ic1, ic2, ic3 = st.columns(3)
            ic1.metric("⏱️ Total Idle Time", f"{total_idle_min/60:,.1f} hrs")
            ic2.metric("🔢 Idle Events", f"{len(idle_events)}")
            ic3.metric("📅 Days Affected", f"{days_with_idle} / {data['Date'].nunique()}")

            st.write("")
            st.markdown("**Idle minutes per day**")
            daily_idle = idle_events.groupby("date")["duration_min"].sum().reset_index()
            fig5 = go.Figure()
            fig5.add_bar(
                x=daily_idle["date"].astype(str).str[5:], y=daily_idle["duration_min"],
                marker_color=AMBER,
                text=daily_idle["duration_min"].apply(lambda v: f"{v:.0f}"), textposition="outside",
                textfont=dict(size=11, color=TEXT),
            )
            fig5.update_layout(
                template="plotly_white", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
                margin=dict(l=10, r=10, t=20, b=10), height=340, showlegend=False,
                xaxis=dict(gridcolor=GRID, type="category"), yaxis=dict(gridcolor=GRID, title="Idle minutes"),
            )
            st.plotly_chart(fig5, width="stretch")

            st.write("")
            st.markdown("**Idle event detail**")
            idle_display = idle_events.copy()
            idle_display["start"] = idle_display["start"].dt.strftime("%Y-%m-%d %H:%M")
            idle_display["end"] = idle_display["end"].dt.strftime("%H:%M")
            idle_display = idle_display[["date", "start", "end", "duration_min"]]
            idle_display.columns = ["Date", "Start", "End", "Duration (min)"]
            st.dataframe(
                idle_display, width="stretch", hide_index=True,
                column_config={"Duration (min)": st.column_config.NumberColumn(format="%.0f min")},
            )

with tab6:
    st.markdown("**Fleet-wide comparison**")
    st.caption("Every vehicle currently loaded, side by side — total distance, variance, fuel efficiency vs. target, idle time, and an estimate of recoverable fuel.")

    fleet_price_key = "fleet_fuel_price"
    fleet_idle_rate_key = "fleet_idle_rate"
    if fleet_price_key not in st.session_state:
        st.session_state[fleet_price_key] = 16000.0
    if fleet_idle_rate_key not in st.session_state:
        st.session_state[fleet_idle_rate_key] = 1.0

    fp1, fp2 = st.columns(2)
    fuel_price = fp1.number_input(
        "Fuel price per liter (optional — for cost estimate, 0 = skip)", min_value=0.0, step=500.0,
        key=fleet_price_key, help="Leave at 0 to see savings in liters only",
    )
    idle_rate = fp2.number_input(
        "Assumed idle fuel burn (L/hour)", min_value=0.1, max_value=5.0, step=0.1,
        key=fleet_idle_rate_key, help="Rule-of-thumb for a light diesel vehicle at idle is ~0.5-0.8 L/hour",
    )

    with st.spinner("Computing fleet overview..."):
        overview_rows = []
        for v in vehicles:
            v_lb = logbook_data.get(v)
            if v_lb is None:
                continue
            v_gps = gps_data.get(v)
            v_raw = gps_raw_data.get(v)
            v_data = compute_reconciliation(v_lb, v_gps)
            v_target = st.session_state.get(f"target_kml_{v}", 9.0)

            v_fuel = compute_fuel_overview(v_data, v_target) if v_data["Refuel_L"].fillna(0).sum() > 0 else None
            avg_kml = v_fuel["avg_kml"] if v_fuel else None

            v_idle_events = compute_idle_events(v_raw, 10, 1.0) if v_raw is not None else pd.DataFrame()
            v_idle_min = v_idle_events["duration_min"].sum() if not v_idle_events.empty else 0.0

            total_lb = v_data["Jumlah_Pemakaian"].sum()
            has_gps_rows = v_data["gps_km"].notna()
            total_gps = v_data.loc[has_gps_rows, "gps_km"].sum() if has_gps_rows.any() else None
            var_km = (total_lb - total_gps) if total_gps is not None else None
            var_pct = (var_km / total_lb) if var_km is not None and total_lb else None

            v_savings = compute_savings_estimate(v_target, v_idle_min, idle_rate, fuel_price or None, var_km)

            overview_rows.append({
                "Vehicle": v, "Logbook KM": total_lb, "GPS KM": total_gps,
                "Variance KM": var_km, "Variance %": var_pct,
                "Avg km/L": avg_kml, "Target km/L": v_target,
                "Idle Hours": v_idle_min / 60.0,
                "Est. Savings (L)": v_savings["total_l"],
            })

        overview_df = pd.DataFrame(overview_rows)
        # Coerce to numeric so a mix of real numbers and Python None (from GPS-less
        # vehicles) becomes proper NaN rather than staying as object dtype — abs()
        # and other numeric ops fail on None but work fine on NaN.
        for col in ["Logbook KM", "GPS KM", "Variance KM", "Variance %", "Avg km/L", "Target km/L", "Idle Hours", "Est. Savings (L)"]:
            overview_df[col] = pd.to_numeric(overview_df[col], errors="coerce")

    if overview_df.empty:
        st.info("No vehicles with logbook data loaded yet.")
    else:
        total_savings_l = overview_df["Est. Savings (L)"].sum()
        total_var_km = overview_df["Variance KM"].abs().sum()
        total_idle_hrs = overview_df["Idle Hours"].sum()
        fleet_msg = (
            f"<div class='insight-box'>🚚 Across <b>{len(overview_df)} vehicle(s)</b>, total logbook-vs-GPS variance "
            f"is <b>{total_var_km:,.0f} km</b> and total idle time is <b>{total_idle_hrs:,.1f} hours</b>. "
            f"Estimated total recoverable fuel across the fleet: <b>{total_savings_l:,.0f} L</b>"
            + (f" (~ Rp.{total_savings_l*fuel_price:,.0f} at your fuel price)" if fuel_price else "")
            + " — combining the fuel implied by the distance gap (variance km / target km/L) and idle time "
              "(idle hours x assumed burn rate). This is an estimate, not an audit finding; use it to prioritize "
              "which vehicles to look at first.</div>"
        )
        st.markdown(fleet_msg, unsafe_allow_html=True)

        fo1, fo2, fo3, fo4 = st.columns(4)
        fo1.metric("🚚 Vehicles", f"{len(overview_df)}")
        fo2.metric("📏 Total Variance", f"{total_var_km:,.0f} km")
        fo3.metric("⏱️ Total Idle Time", f"{total_idle_hrs:,.1f} hrs")
        fo4.metric(
            "💧 Est. Total Savings", f"{total_savings_l:,.0f} L",
            delta=f"~Rp.{total_savings_l*fuel_price:,.0f} at set price" if fuel_price else None,
        )

        st.write("")
        st.markdown("**Variance by vehicle**")
        fig6 = go.Figure()
        fig6.add_bar(x=overview_df["Vehicle"], y=overview_df["Logbook KM"], name="Logbook KM", marker_color=ACCENT,
                     text=overview_df["Logbook KM"].apply(lambda v: f"{v:.0f}"), textposition="outside",
                     textfont=dict(size=11, color=TEXT))
        fig6.add_bar(x=overview_df["Vehicle"], y=overview_df["GPS KM"], name="GPS KM", marker_color=CYAN,
                     text=overview_df["GPS KM"].apply(lambda v: f"{v:.0f}" if pd.notna(v) else ""), textposition="outside",
                     textfont=dict(size=11, color=TEXT))
        fig6.update_layout(
            barmode="group", template="plotly_white", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
            legend=dict(orientation="h", y=1.15, font=dict(color=DIM)),
            margin=dict(l=10, r=10, t=30, b=10), height=380,
            xaxis=dict(gridcolor=GRID, type="category"), yaxis=dict(gridcolor=GRID, title="KM"),
        )
        st.plotly_chart(fig6, width="stretch")

        st.write("")
        st.markdown("**Estimated recoverable fuel by vehicle**")
        fig7 = go.Figure()
        fig7.add_bar(x=overview_df["Vehicle"], y=overview_df["Est. Savings (L)"], marker_color=AMBER,
                     text=overview_df["Est. Savings (L)"].apply(lambda v: f"{v:.0f}"), textposition="outside",
                     textfont=dict(size=11, color=TEXT))
        fig7.update_layout(
            template="plotly_white", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
            margin=dict(l=10, r=10, t=20, b=10), height=340, showlegend=False,
            xaxis=dict(gridcolor=GRID, type="category"), yaxis=dict(gridcolor=GRID, title="Estimated liters"),
        )
        st.plotly_chart(fig7, width="stretch")

        st.write("")
        st.markdown("**Fleet overview table**")
        disp = overview_df.copy()
        disp["Variance %"] = disp["Variance %"].apply(lambda v: f"{v*100:.1f}%" if pd.notna(v) else "-")
        st.dataframe(
            disp, width="stretch", hide_index=True,
            column_config={
                "Logbook KM": st.column_config.NumberColumn(format="%.0f"),
                "GPS KM": st.column_config.NumberColumn(format="%.0f"),
                "Variance KM": st.column_config.NumberColumn(format="%.0f"),
                "Avg km/L": st.column_config.NumberColumn(format="%.1f"),
                "Target km/L": st.column_config.NumberColumn(format="%.1f"),
                "Idle Hours": st.column_config.NumberColumn(format="%.1f"),
                "Est. Savings (L)": st.column_config.NumberColumn(format="%.0f"),
            },
        )
        st.caption("Est. Savings = (Variance KM / Target km/L) + (Idle Hours x assumed idle burn rate). A rough prioritization estimate, not a precise audit figure.")

with tab7:
    st.markdown("**Generate a report for this vehicle**")
    st.caption("A print-ready PDF with KPIs, the daily pattern, cumulative trend, day-by-day detail, fuel analysis, and idle time / savings estimate for the currently selected vehicle.")

    this_target_key = f"target_kml_{plate}"
    this_tol_key = f"tolerance_{plate}"
    default_target = st.session_state.get(this_target_key, 9.0)
    default_tol = st.session_state.get(this_tol_key, 15)

    report_idle_rate = st.session_state.get("fleet_idle_rate", 0.6)
    report_fuel_price = st.session_state.get("fleet_fuel_price", 0.0) or None

    if st.button(f"📄 Generate report for {plate}", type="primary"):
        try:
            with st.spinner("Building PDF..."):
                pdf_bytes = build_single_vehicle_pdf(
                    plate, data, default_target, default_tol,
                    gps_raw_df, report_idle_rate, report_fuel_price,
                )
            st.download_button(
                "⬇ Download vehicle report (PDF)", data=pdf_bytes,
                file_name=f"{plate.replace(' ', '_')}_report.pdf", mime="application/pdf",
            )
            st.success("Report ready — click above to download.")
        except RuntimeError as e:
            st.error(str(e))

    st.divider()
    st.markdown("**Generate a combined report for all vehicles**")
    st.caption(f"One PDF covering all {len(vehicles)} vehicle(s) currently loaded, each with its own full section — same content as above, one after another, plus a fleet overview table up front.")

    if st.button("📄 Generate fleet-wide report (all vehicles)", type="secondary"):
        try:
            with st.spinner(f"Building fleet report for {len(vehicles)} vehicle(s)..."):
                payloads = []
                for v in vehicles:
                    v_lb = logbook_data.get(v)
                    if v_lb is None:
                        continue  # GPS-only vehicles have nothing to reconcile against
                    v_gps = gps_data.get(v)
                    v_raw = gps_raw_data.get(v)
                    v_data = compute_reconciliation(v_lb, v_gps)
                    v_target = st.session_state.get(f"target_kml_{v}", 9.0)
                    v_tol = st.session_state.get(f"tolerance_{v}", 15)
                    payloads.append({
                        "plate": v, "data": v_data,
                        "target_kml": v_target, "tolerance_pct": v_tol,
                        "raw_df": v_raw, "idle_rate_lph": report_idle_rate, "fuel_price": report_fuel_price,
                    })
                fleet_pdf_bytes = build_fleet_pdf(payloads)
            st.download_button(
                "⬇ Download fleet report (PDF)", data=fleet_pdf_bytes,
                file_name="fleet_reconciliation_report.pdf", mime="application/pdf",
            )
            st.success(f"Fleet report ready ({len(payloads)} vehicle(s) included) — click above to download.")
        except RuntimeError as e:
            st.error(str(e))