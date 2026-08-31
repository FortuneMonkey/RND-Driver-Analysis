import re
import os
import pandas as pd
import streamlit as st


def norm_plate(s):
    """Extract a normalized plate code (e.g. 'BM 1356 CW') from a string."""
    if not isinstance(s, str) or not s.strip():
        return ""
    s2 = s.upper()
    m = re.search(r"[A-Z]{1,2}\s?\d{2,4}\s?[A-Z]{1,3}", s2)
    if m:
        return re.sub(r"\s+", " ", m.group(0)).strip()
    return s2.strip()


def _file_ext(file_obj):
    """Return the lowercase file extension, e.g. '.csv' or '.xlsx'."""
    name = getattr(file_obj, "name", "") or ""
    return os.path.splitext(name)[1].lower()


def _is_csv(file_obj):
    return _file_ext(file_obj) == ".csv"


def _plate_from_filename(file_obj):
    """CSV files have no 'sheet name' to pull a plate from — fall back to the filename."""
    name = getattr(file_obj, "name", "") or ""
    base = os.path.splitext(name)[0]
    return norm_plate(base)


def _seek0(file_obj):
    """Reset the read cursor so the same uploaded file can be read more than once."""
    if hasattr(file_obj, "seek"):
        file_obj.seek(0)


def _clean_logbook_df(df):
    """Shared cleanup applied to a raw logbook table, regardless of CSV or Excel source."""
    if df.shape[1] < 7:
        return None
    cols = ["No", "Tujuan", "Berangkat_Tgl", "Berangkat_KM", "Tiba_Tgl",
            "Tiba_KM", "Jumlah_Pemakaian", "Refuel_L", "Fuel_Consum", "Keterangan"]
    df.columns = cols[: df.shape[1]]
    df = df[pd.to_numeric(df["No"], errors="coerce").notna()].copy()
    if df.empty:
        return None
    df["Date"] = pd.to_datetime(df["Berangkat_Tgl"]).dt.date
    df["Jumlah_Pemakaian"] = pd.to_numeric(df["Jumlah_Pemakaian"], errors="coerce").fillna(0)
    df["Tujuan"] = df["Tujuan"].fillna("")
    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def parse_logbook(file_bytes):
    """
    Parse a logbook file into {plate: DataFrame}.
    - .xlsx/.xls: one sheet per vehicle, plate taken from the sheet name (as before).
    - .csv: a single vehicle's logbook, plate taken from the filename (e.g.
      'BM 1356 CW.csv' -> 'BM 1356 CW'), since CSV has no concept of sheets.
      Upload one CSV per vehicle (the app already supports multiple GPS files
      the same way — do the same for multiple logbook CSVs if needed).
    """
    result = {}

    if _is_csv(file_bytes):
        plate = _plate_from_filename(file_bytes)
        _seek0(file_bytes)
        df = pd.read_csv(file_bytes, header=0, skiprows=[1])
        cleaned = _clean_logbook_df(df)
        if cleaned is not None:
            result[plate] = cleaned
        return result

    # Excel: one sheet per vehicle
    xls = pd.ExcelFile(file_bytes)
    for sheet in xls.sheet_names:
        plate = norm_plate(sheet)
        df = pd.read_excel(xls, sheet_name=sheet, header=0, skiprows=[1])
        cleaned = _clean_logbook_df(df)
        if cleaned is not None:
            result[plate] = cleaned
    return result


def _find_gps_header_row(raw):
    """Locate the row containing the real column headers (has 'GPS Time' but not 'GPS Time UTC')."""
    for i in range(min(10, len(raw))):
        row_str = raw.iloc[i].astype(str)
        if row_str.str.contains("GPS Time", na=False).any() and not row_str.str.contains("UTC", na=False).any():
            return i
    return 2  # fall back to the known default position if detection fails


@st.cache_data(show_spinner=False)
def parse_gps(file_bytes):
    """
    Parse a GPS history export into (daily_aggregates, raw_pings), both dicts
    keyed by plate. Supports .xlsx/.xls and .csv — same column layout and
    header-row auto-detection either way.
    """
    is_csv = _is_csv(file_bytes)

    if is_csv:
        _seek0(file_bytes)
        raw = pd.read_csv(file_bytes, header=None)
    else:
        xls = pd.ExcelFile(file_bytes)
        sheet = next((s for s in xls.sheet_names if "report" in s.lower()), xls.sheet_names[0])
        raw = pd.read_excel(xls, sheet_name=sheet, header=None)

    header_row = _find_gps_header_row(raw)

    if is_csv:
        _seek0(file_bytes)
        df = pd.read_csv(file_bytes, header=header_row)
    else:
        df = pd.read_excel(xls, sheet_name=sheet, header=header_row)

    veh_fallback = norm_plate(str(raw.iloc[0, 1])) if raw.shape[0] > 0 and raw.shape[1] > 1 else ""
    if not veh_fallback:
        veh_fallback = _plate_from_filename(file_bytes)  # e.g. 'History_BM_1356_CW.csv'

    if "Vehicle Code" in df.columns:
        df["Plate"] = df["Vehicle Code"].apply(lambda x: norm_plate(str(x)) or veh_fallback)
    else:
        df["Plate"] = veh_fallback

    df["GPS Time"] = pd.to_datetime(df["GPS Time"], format="%d/%m/%Y %H:%M:%S", errors="coerce")
    df["Date"] = df["GPS Time"].dt.date

    # Some GPS exports include an address/location column (name varies by
    # tracker platform, e.g. "Location", "Location(Get All Locations)").
    # Detect it by name rather than assuming an exact header, and normalize
    # it to a plain "Location" column so downstream code has one name to rely on.
    loc_col = next((c for c in df.columns if "location" in str(c).lower()), None)
    if loc_col is not None and loc_col != "Location":
        df["Location"] = df[loc_col]
    raw_cols = ["GPS Time", "Speed (Km/hr)", "ACC", "Mileage(KM)"]
    if "Location" in df.columns:
        raw_cols.append("Location")

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
        raw_out[plate] = g[raw_cols].sort_values("GPS Time").reset_index(drop=True)
    return out, raw_out