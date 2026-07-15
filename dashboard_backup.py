"""
Fleet Logbook <-> GPS Reconciliation Dashboard (Streamlit)
-----------------------------------------------------------
Run locally with:
    pip install streamlit pandas openpyxl plotly
    streamlit run fleet_dashboard_streamlit.py

Upload your logbook workbook (one sheet per vehicle) and one or more GPS
history exports. Vehicles are matched automatically by plate number.
"""

import re
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Fleet Telemetry Reconciliation", layout="wide", page_icon="🛰️")

# ---------------------------------------------------------------------------
# Styling — dark "fleet ops / telemetry console" theme
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .stApp { background-color: #12161a; }
    section[data-testid="stSidebar"] { background-color: #171d23; border-right: 1px solid #2a323c; }
    div[data-testid="stMetric"] {
        background-color: #1a2027; border: 1px solid #2a323c; border-radius: 8px;
        padding: 14px 16px 10px; border-left: 3px solid #ff7a33;
    }
    div[data-testid="stMetricLabel"] { color: #8b98a3 !important; font-size: 11px; letter-spacing: .08em; text-transform: uppercase; }
    div[data-testid="stMetricValue"] { font-family: 'Courier New', monospace; color: #e7ecf0 !important; font-size: 1.7rem !important; }
    div[data-testid="stMetricDelta"] { font-family: 'Courier New', monospace; }
    h1, h2, h3 { color: #e7ecf0 !important; }
    p, span, label, .stCaption { color: #c3ccd3; }
    .stDataFrame { border: 1px solid #2a323c; border-radius: 8px; }
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #1a2027; border-radius: 6px 6px 0 0; padding: 8px 18px;
        color: #8b98a3; border: 1px solid #2a323c; border-bottom: none;
    }
    .stTabs [aria-selected="true"] { color: #e7ecf0 !important; border-top: 2px solid #ff7a33; }
    .insight-box {
        background: linear-gradient(135deg, rgba(255,122,51,0.10), rgba(53,194,193,0.06));
        border: 1px solid #2a323c; border-left: 3px solid #ff7a33;
        border-radius: 8px; padding: 14px 18px; margin-bottom: 6px;
        font-size: 14px; line-height: 1.65;
    }
    .insight-box b { color: #e7ecf0; }
    .badge {
        display:inline-block; padding: 2px 10px; border-radius: 20px;
        font-size: 11px; font-weight: 600; font-family: 'Courier New', monospace;
    }
    hr { border-color: #2a323c !important; }
    div[data-testid="stFileUploaderDropzone"] { background-color: #1a2027; border-color: #2a323c; }
</style>
""", unsafe_allow_html=True)

ACCENT = "#ff7a33"
CYAN = "#35c2c1"
AMBER = "#f2a93b"
GREEN = "#4caf6d"
RED = "#e1523d"
DIM = "#8b98a3"
GRID = "#2a323c"
PANEL = "#1a2027"
FLAG_COLORS = {"OK": GREEN, "HIGH": AMBER, "STANDBY": RED, "NO-GPS": DIM}
FLAG_ICONS = {"OK": "✅", "HIGH": "⚠️", "STANDBY": "🚩", "NO-GPS": "❔"}


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
    return out


def compute_reconciliation(lb_df, gps_df):
    merged = lb_df.merge(gps_df, on="Date", how="left") if gps_df is not None else lb_df.copy()
    if "gps_km" not in merged.columns:
        merged["gps_km"] = pd.NA
        merged["max_speed"] = pd.NA
        merged["acc_on"] = pd.NA

    merged["variance_km"] = merged["Jumlah_Pemakaian"] - merged["gps_km"]
    merged["variance_pct"] = merged.apply(
        lambda r: (r["variance_km"] / r["Jumlah_Pemakaian"])
        if r["Jumlah_Pemakaian"] and pd.notna(r["Jumlah_Pemakaian"]) and r["Jumlah_Pemakaian"] != 0
        else None,
        axis=1,
    )

    def flag(r):
        if pd.isna(r["gps_km"]):
            return "NO-GPS"
        if r["Jumlah_Pemakaian"] == 0 and r["gps_km"] > 1:
            return "STANDBY"
        if r["variance_pct"] is not None and abs(r["variance_pct"]) > 0.15:
            return "HIGH"
        return "OK"

    merged["flag"] = merged.apply(flag, axis=1)
    return merged


def compute_fuel_segments(data):
    """
    Compute distance-per-refuel ('full-to-full') fuel efficiency segments.
    Each segment = km driven since the previous refuel, divided by the liters
    added at the closing refuel. The trailing km since the last refuel (with
    no closing fill yet) is returned separately as 'pending'.
    """
    d = data.sort_values("Date").reset_index(drop=True)
    segments = []
    seg_lb_km, seg_gps_km, seg_gps_valid = 0.0, 0.0, True
    period_start = d["Date"].iloc[0] if len(d) else None

    for _, row in d.iterrows():
        seg_lb_km += row["Jumlah_Pemakaian"] or 0
        if pd.notna(row.get("gps_km")):
            seg_gps_km += row["gps_km"]
        else:
            seg_gps_valid = False

        refuel = row.get("Refuel_L")
        if pd.notna(refuel) and refuel > 0:
            segments.append({
                "period": f"{period_start} → {row['Date']}",
                "liters": refuel,
                "km_logbook": seg_lb_km,
                "km_gps": seg_gps_km if seg_gps_valid else None,
                "kml_logbook": seg_lb_km / refuel if refuel else None,
                "kml_gps": (seg_gps_km / refuel) if (refuel and seg_gps_valid) else None,
            })
            seg_lb_km, seg_gps_km, seg_gps_valid = 0.0, 0.0, True
            period_start = None  # set on next iteration below

        if period_start is None:
            period_start = row["Date"]

    pending_km = seg_lb_km if seg_lb_km > 0 else None
    seg_df = pd.DataFrame(segments)

    if not seg_df.empty:
        # Prefer GPS-based efficiency where available, else fall back to logbook
        seg_df["kml_best"] = seg_df["kml_gps"].where(seg_df["kml_gps"].notna(), seg_df["kml_logbook"])

    return seg_df, pending_km


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
    if len(standby_issue):
        lines.append(f"⚠️ <b>{len(standby_issue)} 'standby' day(s)</b> show real GPS movement — worth checking for unlogged trips.")
    else:
        lines.append("No standby-day misuse detected — days marked idle show no meaningful GPS movement.")

    high = data[data["flag"] == "HIGH"]
    if len(high):
        worst = high.loc[high["variance_km"].abs().idxmax()]
        lines.append(
            f"<b>{len(high)} day(s)</b> exceed the 15% variance threshold — biggest gap on "
            f"<b>{worst['Date']}</b> ({worst['Jumlah_Pemakaian']:.0f} km logged vs {worst['gps_km']:.1f} km GPS)."
        )

    return " ".join(lines)


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
if gps_files:
    for f in gps_files:
        parsed = parse_gps(f)
        gps_data.update(parsed)

if logbook_file:
    st.sidebar.success(f"{len(logbook_data)} vehicle sheet(s) loaded", icon="📘")
if gps_files:
    st.sidebar.success(f"{len(gps_files)} GPS file(s) loaded", icon="🛰️")

vehicles = sorted(set(logbook_data.keys()) | set(gps_data.keys()))

st.markdown(
    f"<span style='color:{ACCENT}; font-family:monospace; font-size:12px; letter-spacing:.15em;'>"
    "// TELEMETRY / RECONCILIATION CONSOLE</span>",
    unsafe_allow_html=True,
)
st.title("🛰️ Fleet Logbook ↔ GPS Reconciliation")
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
flagged = (data["flag"].isin(["HIGH", "STANDBY"])).sum()
max_speed = data["max_speed"].max() if data["max_speed"].notna().any() else None

st.subheader(f"📋 {plate}")

st.markdown(f"<div class='insight-box'>💡 {build_insight(data, total_lb, total_gps, overall_var, flagged)}</div>", unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("📘 Logbook Total", f"{total_lb:,.0f} km")
c2.metric("🛰️ GPS Total", f"{total_gps:,.0f} km" if total_gps is not None else "—")
c3.metric(
    "📊 Overall Variance", f"{overall_var*100:,.1f}%" if overall_var is not None else "—",
    delta="within tolerance" if overall_var is not None and abs(overall_var) <= 0.15 else ("check needed" if overall_var is not None else None),
    delta_color="normal" if overall_var is not None and abs(overall_var) <= 0.15 else "inverse",
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
tab1, tab2, tab3, tab4 = st.tabs(["📈 Daily Pattern", "📉 Cumulative Trend", "📋 Detail Table", "⛽ Fuel Analysis"])

with tab1:
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown("**Logbook vs. GPS distance per day**")
        st.caption("Bars show the exact KM reported by each source; amber line = variance % (right axis), drawn continuously through standby days. Flagged days are marked ⚠️ / 🚩 above the bars.")

        fig1 = go.Figure()
        fig1.add_bar(
            x=dates_str, y=data["Jumlah_Pemakaian"], name="Logbook KM", marker_color=ACCENT,
            text=data["Jumlah_Pemakaian"].apply(lambda v: f"{v:.0f}"), textposition="outside",
            textfont=dict(size=11, color="#e7ecf0"),
        )
        fig1.add_bar(
            x=dates_str, y=data["gps_km"], name="GPS KM", marker_color=CYAN,
            text=data["gps_km"].apply(lambda v: f"{v:.0f}" if pd.notna(v) else ""), textposition="outside",
            textfont=dict(size=11, color="#e7ecf0"),
        )

        var_pct_display = data["variance_pct"].apply(lambda v: v * 100 if pd.notna(v) else None)
        fig1.add_trace(go.Scatter(
            x=dates_str, y=var_pct_display, name="Variance %", mode="lines+markers",
            line=dict(color=AMBER, width=2), marker=dict(size=6),
            yaxis="y2", connectgaps=True,
        ))

        # Annotate flagged days above their bars
        top_val = data[["Jumlah_Pemakaian", "gps_km"]].max(axis=1).fillna(data["Jumlah_Pemakaian"])
        for i, row in data.iterrows():
            if row["flag"] in ("HIGH", "STANDBY"):
                icon = FLAG_ICONS[row["flag"]]
                fig1.add_annotation(
                    x=dates_str.iloc[i], y=top_val.iloc[i], text=icon,
                    showarrow=False, yshift=26, font=dict(size=14),
                )

        fig1.update_layout(
            barmode="group", template="plotly_dark",
            paper_bgcolor=PANEL, plot_bgcolor=PANEL,
            legend=dict(orientation="h", y=1.15, font=dict(color=DIM)),
            margin=dict(l=10, r=10, t=30, b=10), height=420,
            xaxis=dict(gridcolor=GRID), yaxis=dict(title="KM", gridcolor=GRID),
            yaxis2=dict(title="Variance %", overlaying="y", side="right", ticksuffix="%", showgrid=False),
            hovermode="x unified",
        )
        st.plotly_chart(fig1, width="stretch")

    with col2:
        st.markdown("**Flag distribution**")
        st.caption("Share of days by reconciliation outcome")
        counts = data["flag"].value_counts()
        fig2 = go.Figure(go.Pie(
            labels=[f"{FLAG_ICONS.get(k,'')} {k}" for k in counts.index], values=counts.values, hole=0.62,
            marker=dict(colors=[FLAG_COLORS.get(k, DIM) for k in counts.index], line=dict(color=PANEL, width=3)),
            textinfo="value+percent", textfont=dict(color="#e7ecf0", size=12),
        ))
        fig2.add_annotation(text=f"{len(data)}<br>days", showarrow=False, font=dict(size=18, color="#e7ecf0"))
        fig2.update_layout(
            template="plotly_dark", paper_bgcolor=PANEL,
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
        template="plotly_dark", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
        legend=dict(orientation="h", y=1.12, font=dict(color=DIM)),
        margin=dict(l=10, r=10, t=20, b=10), height=440,
        xaxis=dict(gridcolor=GRID), yaxis=dict(title="Cumulative KM", gridcolor=GRID),
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
        table_df = table_df[table_df["flag"].isin(["HIGH", "STANDBY"])]

    display_df = table_df[["Date", "Tujuan", "Jumlah_Pemakaian", "gps_km", "variance_pct", "max_speed", "acc_on", "flag"]].copy()
    display_df.columns = ["Date", "Route", "Logbook KM", "GPS KM", "Variance", "Max Speed", "ACC-On Pings", "Flag"]
    display_df["Variance"] = display_df["Variance"] * 100  # convert fraction to percentage points
    display_df["Flag"] = display_df["Flag"].apply(lambda f: f"{FLAG_ICONS.get(f,'')} {f}")

    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,
        column_config={
            "Logbook KM": st.column_config.NumberColumn(format="%.0f km"),
            "GPS KM": st.column_config.NumberColumn(format="%.1f km"),
            "Variance": st.column_config.ProgressColumn(
                format="%.0f%%", min_value=-100.0, max_value=100.0,
                help="Logbook−GPS variance as a share of logbook KM",
            ),
            "Max Speed": st.column_config.NumberColumn(format="%.0f km/h"),
            "ACC-On Pings": st.column_config.NumberColumn(format="%d"),
        },
    )
    st.caption(f"Showing {len(display_df)} of {len(data)} days.")

with tab4:
    seg_df, pending_km = compute_fuel_segments(data)

    if seg_df.empty:
        st.info("No refuel entries found in the logbook for this vehicle — nothing to analyze yet.")
    else:
        total_liters = seg_df["liters"].sum()
        total_km_lb = seg_df["km_logbook"].sum()
        total_km_gps = seg_df["km_gps"].sum() if seg_df["km_gps"].notna().any() else None
        overall_kml_lb = total_km_lb / total_liters if total_liters else None
        overall_kml_gps = (total_km_gps / total_liters) if (total_km_gps and total_liters) else None
        overall_kml_best = overall_kml_gps or overall_kml_lb

        st.markdown("**Set a target fuel efficiency for this vehicle**")
        st.caption("e.g. the manufacturer-rated or fleet-standard km/L for this vehicle class. Defaults to this vehicle's own observed average — override it with a known spec if you have one.")

        target_key = f"target_kml_{plate}"
        if target_key not in st.session_state:
            st.session_state[target_key] = round(overall_kml_best, 1) if overall_kml_best else 8.0
        tol_key = f"tolerance_{plate}"
        if tol_key not in st.session_state:
            st.session_state[tol_key] = 15

        tcol1, tcol2 = st.columns([1, 2])
        target_kml = tcol1.number_input("Target km/L", min_value=1.0, max_value=50.0, step=0.5, key=target_key)
        tolerance_pct = tcol2.slider("Tolerance band (±%)", min_value=5, max_value=40, step=5, key=tol_key,
                                       help="Cycles within this % of the target are considered normal")

        tol = tolerance_pct / 100.0
        seg_df["deviation_pct"] = (seg_df["kml_best"] - target_kml) / target_kml
        seg_df["fuel_flag"] = seg_df["deviation_pct"].apply(
            lambda v: "BELOW TARGET" if v < -tol else ("ABOVE TARGET" if v > tol else "OK")
        )
        below_count = (seg_df["fuel_flag"] == "BELOW TARGET").sum()
        above_count = (seg_df["fuel_flag"] == "ABOVE TARGET").sum()

        overall_vs_target = (overall_kml_best - target_kml) / target_kml if overall_kml_best else None

        st.markdown(
            f"<div class='insight-box'>⛽ Across <b>{len(seg_df)} refuel cycle(s)</b> totaling "
            f"<b>{total_liters:,.0f} L</b>, this vehicle averages <b>{overall_kml_best:,.1f} km/L</b> "
            f"against a target of <b>{target_kml:.1f} km/L</b> "
            f"({'above' if overall_vs_target and overall_vs_target > 0 else 'below'} target by {abs(overall_vs_target*100):,.0f}%). "
            + (f"<b>{below_count} cycle(s)</b> fall more than {tolerance_pct}% below target — worth checking against fuel receipts. " if below_count else "No cycles fall meaningfully below target. ")
            + (f"<b>{above_count} cycle(s)</b> run more than {tolerance_pct}% above target, which can mean genuinely efficient driving, a partial (not full) refuel, or under-reported distance." if above_count else "")
            + (f" There's also <b>{pending_km:.0f} km</b> driven since the last refuel with no closing fill-up yet." if pending_km else "")
            + "</div>",
            unsafe_allow_html=True,
        )
        st.caption("⚠️ Per-cycle figures still assume each refuel tops the tank to full — if fill-ups aren't always complete, individual cycles will be noisy. The overall average above (total km ÷ total liters) doesn't depend on that assumption and is the more reliable number.")

        fc1, fc2, fc3 = st.columns(3)
        fc1.metric("⛽ Total Refueled", f"{total_liters:,.0f} L")
        fc2.metric(
            "📏 Avg vs Target", f"{overall_kml_best:,.1f} km/L",
            delta=f"{overall_vs_target*100:+.0f}% vs {target_kml:.1f} target" if overall_vs_target is not None else None,
            delta_color="normal" if overall_vs_target is not None and overall_vs_target >= -tol else "inverse",
        )
        fc3.metric(
            "🚩 Below Target", f"{below_count} / {len(seg_df)}",
            delta="clean" if below_count == 0 else f"{below_count} to review",
            delta_color="normal" if below_count == 0 else "inverse",
        )

        st.write("")
        st.markdown("**Fuel efficiency per refuel cycle vs. target**")
        st.caption("Bars = km/L per cycle (GPS-based where available, else logbook); dashed line = your target")

        kml_col = "kml_gps" if seg_df["kml_gps"].notna().any() else "kml_logbook"
        flag_bar_colors = {"BELOW TARGET": RED, "ABOVE TARGET": AMBER, "OK": CYAN}
        bar_colors = [flag_bar_colors[f] for f in seg_df["fuel_flag"]]

        fig4 = go.Figure()
        fig4.add_bar(
            x=seg_df["period"], y=seg_df[kml_col], marker_color=bar_colors,
            text=seg_df[kml_col].apply(lambda v: f"{v:.1f}"), textposition="outside",
            textfont=dict(size=11, color="#e7ecf0"), name="km/L",
        )
        fig4.add_hline(y=target_kml, line_dash="dash", line_color="#e7ecf0",
                        annotation_text=f"target {target_kml:.1f} km/L", annotation_font_color="#e7ecf0")
        fig4.update_layout(
            template="plotly_dark", paper_bgcolor=PANEL, plot_bgcolor=PANEL,
            margin=dict(l=10, r=10, t=20, b=10), height=380, showlegend=False,
            xaxis=dict(gridcolor=GRID, title="Refuel cycle"), yaxis=dict(gridcolor=GRID, title="km/L"),
        )
        st.plotly_chart(fig4, width="stretch")

        st.write("")
        st.markdown("**Refuel cycle detail**")
        seg_display = seg_df[["period", "liters", "km_logbook", "km_gps", "kml_logbook", "kml_gps", "deviation_pct", "fuel_flag"]].copy()
        seg_display.columns = ["Period", "Liters", "Logbook KM", "GPS KM", "km/L (Logbook)", "km/L (GPS)", "vs Target", "Flag"]
        seg_display["vs Target"] = seg_display["vs Target"] * 100
        flag_icon = {"BELOW TARGET": "🔻", "ABOVE TARGET": "🔺", "OK": "✅"}
        seg_display["Flag"] = seg_display["Flag"].apply(lambda f: f"{flag_icon[f]} {f}")
        st.dataframe(
            seg_display, width="stretch", hide_index=True,
            column_config={
                "Liters": st.column_config.NumberColumn(format="%.0f L"),
                "Logbook KM": st.column_config.NumberColumn(format="%.0f km"),
                "GPS KM": st.column_config.NumberColumn(format="%.1f km"),
                "km/L (Logbook)": st.column_config.NumberColumn(format="%.1f"),
                "km/L (GPS)": st.column_config.NumberColumn(format="%.1f"),
                "vs Target": st.column_config.NumberColumn(format="%+.0f%%"),
            },
        )
        if pending_km:
            st.caption(f"Not shown: {pending_km:.0f} km driven since the last refuel, awaiting a closing fill-up.")