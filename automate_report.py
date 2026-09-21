import io
import glob
import os
import sys
import tempfile
import datetime as dt

# parse_logbook/parse_gps are decorated with @st.cache_data. Outside a running
# Streamlit app, its hasher tries to fingerprint file-like objects by opening
# their `.name` attribute as a real path on disk -- which breaks here, since
# we set `.name` on an in-memory BytesIO just for extension detection, not
# because it points to a file in the script's working directory. There's no
# benefit to caching in a one-shot script anyway, so disable it before the
# decorated functions are imported.
import streamlit as st


def _no_cache(*args, **kwargs):
    if args and callable(args[0]):
        return args[0]
    return lambda func: func


st.cache_data = _no_cache

from modules.parsing import parse_logbook, parse_gps
from modules.reconciliation import compute_reconciliation
from modules.pdf_report import build_single_vehicle_pdf, build_fleet_pdf
from modules.cleaning import build_clean_workbook
from modules.screenshot import render_pdf_page
from modules.emailer import send_report_email

# ---------------------------------------------------------------------------
# Config — edit here, not inline below
# ---------------------------------------------------------------------------
TARGET_VEHICLES = ["BM 1356 CW", "BM 8079 CK", "BM 8499 CK", "BM 8988 CJ"]

RAW_LOGBOOK_ROOT = r"S:\Temporary\Alvin\1. R&D Car\1. Logbook"
GPS_TRACKER_ROOT = r"S:\Temporary\GPS Tracker\gps_history"
CLEAN_LOGBOOK_DIR = r"S:\Temporary\GPS Tracker\logbook_data"
REPORT_ROOT = r"S:\Temporary\GPS Tracker\gps_and_logbook_report"

GPS_FILE_EXTS = (".xlsx", ".xls", ".csv")

# Same defaults the dashboard falls back to for a fresh vehicle
DEFAULT_CONFIG = {
    "flag_threshold_km": 10,
    "target_kml": 9.0,
    "idle_threshold_min": 10,
    "tolerance_pct": 15,
    "fuel_price": 16000.0,
    "idle_rate_lph": 1.0,
}

EMAIL_CONFIG = {
    "enabled": True,          # set False to skip the email step entirely
    "send": True,            # False = open as a draft in Outlook for review; True = send immediately
    "to": ["Betty_AndrianySirait@aprilasia.com", "kira_theresa@aprilasia.com", "Muhamar_Prayogi@aprilasia.com", "Husni_Mubarok@aprilasia.com", "Sahat_Manimbo@globalnetlcl.com", "Halimah_Tanjung@globalnetlcl.com", "jessika_sembiring@globalnetlcl.com"],   # TODO: real recipients
    "cc": ["Srikumar@aprilasia.com", "muhammad_yuliarto@aprilasia.com", "iswandi@aprilasia.com", "susanna_chitraresmi@aprilasia.com", "Alvaro_Duran@aprilasia.com", "sabar_siregar@aprilasia.com"],
    "subject_template": "R&D Vehicle Report - {month_name} {year}",
}


# ---------------------------------------------------------------------------
# Month/period arithmetic
# ---------------------------------------------------------------------------
def shift_month(year, month, delta):
    m = month - 1 + delta
    y = year + m // 12
    m = m % 12 + 1
    return y, m


def target_period(today=None):
    """
    Returns everything derived from "today": the target (previous) month,
    the 26th-to-25th logbook cycle, and the 25th-to-25th GPS folder window.
    """
    today = today or dt.date.today()
    ty, tm = shift_month(today.year, today.month, -1)          # target month = previous month
    py, pm = shift_month(ty, tm, -1)                            # month before that

    logbook_start = dt.date(py, pm, 26)
    logbook_end = dt.date(ty, tm, 25)
    gps_start = dt.date(py, pm, 25)
    gps_end = dt.date(ty, tm, 25)

    return {
        "year": ty, "month": tm,
        "month_name": dt.date(ty, tm, 1).strftime("%B"),
        "logbook_start": logbook_start, "logbook_end": logbook_end,
        "gps_start": gps_start, "gps_end": gps_end,
    }


# ---------------------------------------------------------------------------
# File discovery
# ---------------------------------------------------------------------------
def find_raw_logbook_file(period):
    folder = os.path.join(RAW_LOGBOOK_ROOT, str(period["year"]))
    stem = f"{period['month']}. Logbook {period['month_name']}"
    for ext in (".xlsx", ".xls"):
        candidate = os.path.join(folder, stem + ext)
        if os.path.exists(candidate):
            return candidate
    # fall back to a case/spacing-tolerant search in case the file was saved slightly differently
    if os.path.isdir(folder):
        stem_lower = stem.lower()
        for name in os.listdir(folder):
            base, ext = os.path.splitext(name)
            if ext.lower() in (".xlsx", ".xls") and base.strip().lower() == stem_lower:
                return os.path.join(folder, name)
    raise FileNotFoundError(
        f"Could not find '{stem}.xlsx' (or .xls) in {folder}. "
        "Check RAW_LOGBOOK_ROOT and the month-name convention in that folder."
    )


def find_gps_folder(period):
    folder_name = f"{period['gps_start']:%Y%m%d}-{period['gps_end']:%Y%m%d}"
    folder = os.path.join(GPS_TRACKER_ROOT, folder_name)
    if not os.path.isdir(folder):
        raise FileNotFoundError(f"GPS folder not found: {folder}")
    return folder


def find_gps_files(gps_folder):
    files = []
    for ext in GPS_FILE_EXTS:
        files.extend(glob.glob(os.path.join(gps_folder, f"*{ext}")))
    return files


# ---------------------------------------------------------------------------
# Small helper: give a plain file path the .name/.seek() shape that
# parse_logbook/parse_gps expect (they were written for Streamlit's
# uploaded-file objects).
# ---------------------------------------------------------------------------
def _as_uploadlike(path):
    with open(path, "rb") as f:
        buf = io.BytesIO(f.read())
    buf.name = os.path.basename(path)
    return buf


# ---------------------------------------------------------------------------
# Steps
# ---------------------------------------------------------------------------
def clean_and_save_logbook(period):
    raw_path = find_raw_logbook_file(period)
    print(f"[1/6] Raw logbook: {raw_path}")

    wb, found, warnings = build_clean_workbook(
        raw_path, TARGET_VEHICLES, period["logbook_start"], period["logbook_end"],
    )
    for w in warnings:
        print(f"      ! {w}")

    missing = [p for p, ok in found.items() if not ok]
    if missing:
        print(f"      ! Vehicles missing from the raw workbook: {missing}")

    os.makedirs(CLEAN_LOGBOOK_DIR, exist_ok=True)
    out_path = os.path.join(
        CLEAN_LOGBOOK_DIR, f"logbook_{period['month_name']}_{period['year']}.xlsx"
    )
    wb.save(out_path)
    print(f"      Clean logbook saved -> {out_path}")
    return out_path


def load_logbook_data(clean_path):
    logbook_data = parse_logbook(_as_uploadlike(clean_path))
    return {plate: df for plate, df in logbook_data.items() if plate in TARGET_VEHICLES}


def load_gps_data(period):
    gps_folder = find_gps_folder(period)
    files = find_gps_files(gps_folder)
    print(f"[2/6] GPS folder: {gps_folder} ({len(files)} file(s))")

    gps_data, gps_raw_data = {}, {}
    for path in files:
        try:
            parsed, raw_parsed = parse_gps(_as_uploadlike(path))
        except Exception as e:
            print(f"      ! Skipped {os.path.basename(path)}: {e}")
            continue
        gps_data.update(parsed)
        gps_raw_data.update(raw_parsed)

    gps_data = {p: d for p, d in gps_data.items() if p in TARGET_VEHICLES}
    gps_raw_data = {p: d for p, d in gps_raw_data.items() if p in TARGET_VEHICLES}
    return gps_data, gps_raw_data


def build_reports(period, report_dir, logbook_data, gps_data, gps_raw_data):
    print(f"[3/6] Per-vehicle reports -> {report_dir}")

    cfg = DEFAULT_CONFIG
    written = []
    for plate in TARGET_VEHICLES:
        lb_df = logbook_data.get(plate)
        if lb_df is None:
            print(f"      ! {plate}: no logbook rows, skipping report.")
            continue

        gps_df = gps_data.get(plate)
        gps_raw_df = gps_raw_data.get(plate)
        if gps_df is None:
            print(f"      ! {plate}: no GPS match this period (report will show logbook-only).")

        data = compute_reconciliation(lb_df, gps_df, cfg["flag_threshold_km"])
        try:
            pdf_bytes = build_single_vehicle_pdf(
                plate, data, cfg["target_kml"], cfg["tolerance_pct"],
                gps_raw_df, cfg["idle_rate_lph"], cfg["fuel_price"], cfg["idle_threshold_min"],
            )
        except RuntimeError as e:
            print(f"      ! {plate}: failed to build PDF ({e})")
            continue

        out_path = os.path.join(report_dir, f"{plate.replace(' ', '_')}_report.pdf")
        with open(out_path, "wb") as f:
            f.write(pdf_bytes)
        written.append(out_path)
        print(f"      {plate} -> {out_path}")

    return written


def build_fleet_report(period, report_dir, logbook_data, gps_data, gps_raw_data):
    """One combined PDF covering every target vehicle, mirroring the dashboard's
    'Generate fleet-wide report' button (modules/dashboard.py tab7)."""
    cfg = DEFAULT_CONFIG
    payloads = []
    for plate in TARGET_VEHICLES:
        lb_df = logbook_data.get(plate)
        if lb_df is None:
            continue  # nothing to reconcile for a vehicle with no logbook rows
        gps_df = gps_data.get(plate)
        raw_df = gps_raw_data.get(plate)
        data = compute_reconciliation(lb_df, gps_df, cfg["flag_threshold_km"])
        payloads.append({
            "plate": plate, "data": data,
            "target_kml": cfg["target_kml"], "tolerance_pct": cfg["tolerance_pct"],
            "raw_df": raw_df, "idle_rate_lph": cfg["idle_rate_lph"],
            "fuel_price": cfg["fuel_price"], "idle_threshold_min": cfg["idle_threshold_min"],
        })

    if not payloads:
        print("      ! No vehicles had logbook data -- skipping fleet report.")
        return None

    try:
        pdf_bytes = build_fleet_pdf(payloads)
    except RuntimeError as e:
        print(f"      ! Fleet report failed to build ({e})")
        return None

    out_path = os.path.join(report_dir, f"Fleet_Report_{period['month_name']}_{period['year']}.pdf")
    with open(out_path, "wb") as f:
        f.write(pdf_bytes)
    print(f"      Fleet report ({len(payloads)} vehicle(s)) -> {out_path}")
    return out_path


def email_fleet_report(period, fleet_pdf_path, vehicle_pdf_paths, report_dir):
    if not EMAIL_CONFIG["enabled"]:
        return
    if not fleet_pdf_path:
        print("      ! No fleet PDF to email, skipping.")
        return

    # Written to the local temp folder, not the network share: it's a
    # throwaway used only to embed inline in the email body, and writing to
    # S:\ here has been unreliable (locked/blocked files, AV scanning, etc.)
    # while local disk just works.
    screenshot_path = os.path.join(
        tempfile.gettempdir(), f"fleet_report_page1_{os.getpid()}.png"
    )
    render_pdf_page(fleet_pdf_path, screenshot_path, page_number=0, dpi=150)

    subject = EMAIL_CONFIG["subject_template"].format(
        month_name=period["month_name"], year=period["year"]
    )
    try:
        send_report_email(
            to_recipients=EMAIL_CONFIG["to"],
            cc_recipients=EMAIL_CONFIG["cc"],
            subject=subject,
            screenshot_path=screenshot_path,
            attachments=[fleet_pdf_path] + vehicle_pdf_paths,
            period_label=f"{period['month_name']} {period['year']}",
            send=EMAIL_CONFIG["send"],
        )
    finally:
        try:
            os.remove(screenshot_path)
        except OSError:
            pass  # best-effort cleanup of the temp file; not worth failing the run over

    mode = "sent" if EMAIL_CONFIG["send"] else "opened as a draft for review"
    print(f"      Email {mode} ({subject}).")


def main():
    period = target_period()
    print(
        f"[0/6] Target period: {period['month_name']} {period['year']} "
        f"(logbook rows {period['logbook_start']}..{period['logbook_end']}, "
        f"GPS folder {period['gps_start']:%Y%m%d}-{period['gps_end']:%Y%m%d})"
    )

    report_dir = os.path.join(REPORT_ROOT, f"{period['year']}_{period['month']:02d}")
    os.makedirs(report_dir, exist_ok=True)

    clean_path = clean_and_save_logbook(period)
    logbook_data = load_logbook_data(clean_path)
    gps_data, gps_raw_data = load_gps_data(period)

    vehicle_pdfs = build_reports(period, report_dir, logbook_data, gps_data, gps_raw_data)

    print("[4/6] Fleet-wide report")
    fleet_pdf = build_fleet_report(period, report_dir, logbook_data, gps_data, gps_raw_data)

    print("[5/6] Email")
    email_fleet_report(period, fleet_pdf, vehicle_pdfs, report_dir)

    print(f"[6/6] Done — {len(vehicle_pdfs)} vehicle report(s) + fleet report written.")


if __name__ == "__main__":
    try:
        main()
    except FileNotFoundError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)