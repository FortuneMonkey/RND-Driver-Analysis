import re
import io
import os
import inspect
import tempfile
import datetime as dt
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from PIL import Image
from fpdf import FPDF

from modules.theme import ACCENT, CYAN, AMBER, GREEN, RED, DIM, GRID
from modules.reconciliation import (
    build_insight, compute_idle_events, compute_savings_estimate, compute_fuel_overview,
)

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


FLAG_MARKERS = {"HIGH": ("▲", AMBER), "STANDBY": ("●", RED), "NO-LOGBOOK": ("◆", "#9333EA")}


def render_daily_chart_png(dates_str, lb_vals, gps_vals, flags=None):
    fig, ax = plt.subplots(figsize=(7.5, 3.2))
    x = range(len(dates_str))
    w = 0.38
    bars_lb = ax.bar([i - w / 2 for i in x], lb_vals, width=w, color=RPT_LB, label="Logbook KM")
    bars_gps = ax.bar([i + w / 2 for i in x], gps_vals, width=w, color=RPT_GPS, label="GPS KM")
    ax.bar_label(bars_lb, fmt="%.0f", fontsize=4, padding=1)
    ax.bar_label(bars_gps, fmt=lambda v: f"{v:.0f}" if v == v else "", fontsize=4, padding=1)  # v==v filters NaN
    ax.set_xticks(list(x))
    ax.set_xticklabels(dates_str, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("KM")
    ax.margins(y=0.22)
    ax.grid(axis="y", color=RPT_GRID, linewidth=0.6)
    ax.set_axisbelow(True)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)

    if flags is not None:
        for i, f in enumerate(flags):
            if f in FLAG_MARKERS:
                marker, color = FLAG_MARKERS[f]
                top_val = max(
                    lb_vals[i] if pd.notna(lb_vals[i]) else 0,
                    gps_vals[i] if pd.notna(gps_vals[i]) else 0,
                )
                ax.annotate(marker, (i, top_val), textcoords="offset points", xytext=(0, 12),
                            fontsize=7, ha="center", color=color)
        # small legend for the flag markers actually present
        present = sorted(set(f for f in flags if f in FLAG_MARKERS), key=list(FLAG_MARKERS).index)
        if present:
            legend_text = "   ".join(f"{FLAG_MARKERS[f][0]} {f}" for f in present)
            ax.text(0.5, -0.32, legend_text, transform=ax.transAxes, ha="center", fontsize=7, color=RPT_TEXT)

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
        png = render_daily_chart_png(dates_str.str[5:], data["Jumlah_Pemakaian"], data["gps_km"], data["flag"].tolist())
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
    pdf.body_text("Note: Logbook KM x Var % = the gap between Logbook and GPS, in km. See the KM columns for which side is higher.")

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


