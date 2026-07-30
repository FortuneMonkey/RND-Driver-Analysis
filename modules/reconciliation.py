import pandas as pd


def compute_reconciliation(lb_df, gps_df, flag_threshold_km=10):
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
        if pd.notna(r["variance_km"]) and abs(r["variance_km"]) > flag_threshold_km:
            return "HIGH"
        return "OK"

    merged["flag"] = merged.apply(flag, axis=1)
    return merged


def compute_fuel_overview(data, target_kml):

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


def build_insight(data, total_lb, total_gps, overall_var, flagged, flag_threshold_km=10):
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
            f"<b>{len(high)} day(s)</b> have a gap of more than {flag_threshold_km} km between Logbook and GPS — biggest gap on "
            f"<b>{worst['Date']}</b> ({worst['Jumlah_Pemakaian']:.0f} km logged vs {worst['gps_km']:.1f} km GPS)."
        )

    return " ".join(lines)
