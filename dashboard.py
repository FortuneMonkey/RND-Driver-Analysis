import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from modules.theme import apply_theme, ACCENT, CYAN, AMBER, DIM, GRID, PANEL, TEXT, FLAG_COLORS, FLAG_ICONS
from modules.parsing import parse_logbook, parse_gps
from modules.reconciliation import (
    compute_reconciliation, compute_fuel_overview, compute_idle_events,
    compute_savings_estimate, build_insight,
)
from modules.pdf_report import build_single_vehicle_pdf, build_fleet_pdf

st.set_page_config(page_title="RND Driver Analysis", layout="wide", page_icon="🛰️")
apply_theme()


# ---------------------------------------------------------------------------
# Sidebar — data in
# ---------------------------------------------------------------------------
st.sidebar.markdown("### 📥 Data In")
logbook_files = st.sidebar.file_uploader(
    "Logbook workbook(s) — Excel with one sheet per vehicle, or one CSV per vehicle",
    type=["xlsx", "xls", "csv"], accept_multiple_files=True,
)
gps_files = st.sidebar.file_uploader(
    "GPS history export(s) — select multiple", type=["xlsx", "xls", "csv"], accept_multiple_files=True
)

logbook_data = {}
if logbook_files:
    for f in logbook_files:
        logbook_data.update(parse_logbook(f))
gps_data = {}
gps_raw_data = {}
if gps_files:
    for f in gps_files:
        parsed, raw_parsed = parse_gps(f)
        gps_data.update(parsed)
        gps_raw_data.update(raw_parsed)

# if logbook_files:
#     st.sidebar.success(f"{len(logbook_files)} file(s) -> {len(logbook_data)} vehicle(s) loaded", icon="📘")
# if gps_files:
#     st.sidebar.success(f"{len(gps_files)} GPS file(s) loaded", icon="🛰️")

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

# ---------------------------------------------------------------------------
# Sidebar — configuration (per-vehicle: flag threshold, target, idle threshold;
# fleet-wide: fuel price, idle burn rate)
# ---------------------------------------------------------------------------
st.sidebar.markdown("### ⚙️ Configuration")

flag_threshold_key = f"flag_threshold_{plate}"
target_key = f"target_kml_{plate}"
idle_threshold_key = f"idle_threshold_{plate}"
st.session_state.setdefault(flag_threshold_key, 10)
st.session_state.setdefault(target_key, 9.0)
st.session_state.setdefault(idle_threshold_key, 10)
st.session_state.setdefault("fleet_fuel_price", 16000.0)
st.session_state.setdefault("fleet_idle_rate", 1.0)

st.sidebar.caption(f"For **{plate}**")
flag_threshold_km = st.sidebar.number_input(
    "Flag threshold (km gap)", min_value=1, max_value=200, step=1, key=flag_threshold_key,
    help="Days with a Logbook-GPS gap bigger than this are flagged HIGH",
)
target_kml = st.sidebar.number_input(
    "Target km/L", min_value=1.0, max_value=50.0, step=0.5, key=target_key,
    help="Manufacturer-rated or fleet-standard efficiency for this vehicle",
)
idle_threshold_min = st.sidebar.number_input(
    "Idle threshold (minutes)", min_value=1, max_value=120, step=1, key=idle_threshold_key,
    help="Minimum continuous ACC-ON, not-moving duration to count as an idle event",
)

st.sidebar.caption("Fleet-wide")
fuel_price = st.sidebar.number_input(
    "Fuel price per liter (Rp)", min_value=0.0, step=500.0, key="fleet_fuel_price",
    help="Leave at 0 to see savings in liters only",
)
idle_rate = st.sidebar.number_input(
    "Assumed idle fuel burn (L/hour)", min_value=0.1, max_value=5.0, step=0.1, key="fleet_idle_rate",
    help="Diesel vehicle at idle is ~0.8-1.2 L/hour",
)

data = compute_reconciliation(lb_df, gps_df, flag_threshold_km)
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
        st.caption(f"Bars show the exact KM reported by each source; amber line = variance % (right axis), drawn continuously through standby days. Days with more than a {flag_threshold_km:.0f} km gap between Logbook and GPS are marked ⚠️ / 🚩 above the bars.")

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

    detail_event = st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
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
    st.caption(f"Showing {len(display_df)} of {len(data)} days. Logbook KM x Variance % = the gap between Logbook and GPS, in km. Compare the KM columns directly to see which side is higher. Days flagged HIGH have a gap of more than {flag_threshold_km:.0f} km.")

    # Per-row raw GPS ping download — click a row above, then download only
    # that day's raw pings (not the whole vehicle's GPS history).
    selected_rows = detail_event["selection"]["rows"] if detail_event and detail_event["selection"] else []
    if selected_rows:
        sel_date = table_df.iloc[selected_rows[0]]["Date"]
        if gps_raw_df is not None:
            day_pings = gps_raw_df[gps_raw_df["GPS Time"].dt.date == sel_date]
            if not day_pings.empty:
                st.download_button(
                    f"⬇ Download raw GPS pings for {sel_date} ({len(day_pings)} pings)",
                    data=day_pings.to_csv(index=False).encode("utf-8"),
                    file_name=f"{plate.replace(' ', '_')}_{sel_date}_gps_pings.csv",
                    mime="text/csv",
                    key=f"dl_detail_{plate}_{sel_date}",
                )
            else:
                st.caption(f"No raw GPS pings recorded for {sel_date}.")
        else:
            st.caption("No raw GPS data available for this vehicle to download.")
    else:
        st.caption("💡 Click a row above to download that day's raw GPS pings only.")

with tab4:
    total_lb = data["Jumlah_Pemakaian"].sum()
    has_gps_rows = data["gps_km"].notna()
    total_gps = data.loc[has_gps_rows, "gps_km"].sum() if has_gps_rows.any() else None

    if total_gps is None:
        st.info("No GPS data matched for this vehicle yet — the fuel estimate needs a KM gap (Logbook vs GPS) to work from.")
    else:
        st.caption(f"Using target of **{target_kml:.1f} km/L** — adjust in the sidebar Configuration section.")

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
    st.caption(f"Periods where the engine was ON but the vehicle wasn't moving, for at least **{idle_threshold_min} minutes** (adjust in the sidebar Configuration section). Long idling burns fuel with zero distance to show for it.")

    IDLE_SPEED_CUTOFF = 1.0  # km/h — speed at or below this counts as stationary (accounts for GPS jitter)

    if gps_raw_df is None:
        st.info("No raw GPS ping data available for this vehicle — idle detection needs a matched GPS export.")
        idle_events = pd.DataFrame()
    else:
        idle_events = compute_idle_events(gps_raw_df, idle_threshold_min, IDLE_SPEED_CUTOFF)

        if idle_events.empty:
            st.markdown(f"<div class='insight-box'>✅ No idle events of {idle_threshold_min}+ minutes detected for this vehicle over the period.</div>", unsafe_allow_html=True)
        else:
            total_idle_min = idle_events["duration_min"].sum()
            longest = idle_events.loc[idle_events["duration_min"].idxmax()]
            days_with_idle = idle_events["date"].nunique()

            st.markdown(
                f"<div class='insight-box'>⏱️ <b>{len(idle_events)} idle event(s)</b> of {idle_threshold_min}+ minutes "
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
            idle_event = st.dataframe(
                idle_display, width="stretch", hide_index=True,
                on_select="rerun", selection_mode="single-row",
                column_config={"Duration (min)": st.column_config.NumberColumn(format="%.0f min")},
            )

            # Per-row raw GPS ping download — click an idle event above, then
            # download only the raw pings that fall inside that event's time window.
            idle_selected_rows = idle_event["selection"]["rows"] if idle_event and idle_event["selection"] else []
            if idle_selected_rows:
                ev = idle_events.iloc[idle_selected_rows[0]]
                if gps_raw_df is not None:
                    event_pings = gps_raw_df[
                        (gps_raw_df["GPS Time"] >= ev["start"]) & (gps_raw_df["GPS Time"] <= ev["end"])
                    ]
                    if not event_pings.empty:
                        st.download_button(
                            f"⬇ Download raw GPS pings for {ev['start'].strftime('%Y-%m-%d %H:%M')}–{ev['end'].strftime('%H:%M')} "
                            f"({len(event_pings)} pings)",
                            data=event_pings.to_csv(index=False).encode("utf-8"),
                            file_name=f"{plate.replace(' ', '_')}_{ev['start'].strftime('%Y%m%d_%H%M')}_idle_gps_pings.csv",
                            mime="text/csv",
                            key=f"dl_idle_{plate}_{idle_selected_rows[0]}",
                        )
                    else:
                        st.caption("No raw GPS pings found in this event's exact time window.")
                else:
                    st.caption("No raw GPS data available for this vehicle to download.")
            else:
                st.caption("💡 Click an idle event above to download the raw GPS pings for that time window only.")

with tab6:
    st.markdown("**Fleet-wide comparison**")
    st.caption(f"Every vehicle currently loaded, side by side — total distance, variance, fuel efficiency vs. target, idle time, and an estimate of recoverable fuel. Using fuel price **Rp {fuel_price:,.0f}/L** and idle burn **{idle_rate:.1f} L/hour** — adjust in the sidebar Configuration section.")

    with st.spinner("Computing fleet overview..."):
        overview_rows = []
        for v in vehicles:
            v_lb = logbook_data.get(v)
            if v_lb is None:
                continue
            v_gps = gps_data.get(v)
            v_raw = gps_raw_data.get(v)
            v_flag_threshold = st.session_state.get(f"flag_threshold_{v}", 10)
            v_data = compute_reconciliation(v_lb, v_gps, v_flag_threshold)
            v_target = st.session_state.get(f"target_kml_{v}", 9.0)

            v_fuel = compute_fuel_overview(v_data, v_target) if v_data["Refuel_L"].fillna(0).sum() > 0 else None
            avg_kml = v_fuel["avg_kml"] if v_fuel else None

            v_idle_threshold = st.session_state.get(f"idle_threshold_{v}", 10)
            v_idle_events = compute_idle_events(v_raw, v_idle_threshold, 1.0) if v_raw is not None else pd.DataFrame()
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

    default_target = target_kml
    default_tol = st.session_state.get(f"tolerance_{plate}", 15)

    report_idle_rate = idle_rate
    report_fuel_price = fuel_price or None

    if st.button(f"📄 Generate report for {plate}", type="primary"):
        try:
            with st.spinner("Building PDF..."):
                pdf_bytes = build_single_vehicle_pdf(
                    plate, data, default_target, default_tol,
                    gps_raw_df, report_idle_rate, report_fuel_price, idle_threshold_min,
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
                    v_flag_threshold = st.session_state.get(f"flag_threshold_{v}", 10)
                    v_data = compute_reconciliation(v_lb, v_gps, v_flag_threshold)
                    v_target = st.session_state.get(f"target_kml_{v}", 9.0)
                    v_tol = st.session_state.get(f"tolerance_{v}", 15)
                    v_idle_threshold = st.session_state.get(f"idle_threshold_{v}", 10)
                    payloads.append({
                        "plate": v, "data": v_data,
                        "target_kml": v_target, "tolerance_pct": v_tol,
                        "raw_df": v_raw, "idle_rate_lph": report_idle_rate, "fuel_price": report_fuel_price,
                        "idle_threshold_min": v_idle_threshold,
                    })
                fleet_pdf_bytes = build_fleet_pdf(payloads)
            st.download_button(
                "⬇ Download fleet report (PDF)", data=fleet_pdf_bytes,
                file_name="fleet_reconciliation_report.pdf", mime="application/pdf",
            )
            st.success(f"Fleet report ready ({len(payloads)} vehicle(s) included) — click above to download.")
        except RuntimeError as e:
            st.error(str(e))