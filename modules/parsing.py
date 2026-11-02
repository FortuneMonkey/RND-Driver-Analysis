import logging
import os
import re

import pandas as pd
import streamlit as st

logger = logging.getLogger(__name__)

# Encodings tried in order for CSV files. utf-8-sig also handles a BOM.
# latin-1 can decode any byte, so it is the last resort and never fails.
_CSV_ENCODINGS = ("utf-8-sig", "cp1252", "latin-1")

_GPS_TIME_FORMATS = (
    "%d/%m/%Y %H:%M:%S",   # original tracker export
    "%d/%m/%Y %H:%M",      # Excel-reformatted rows (no seconds, no zero padding)
    "%Y-%m-%d %H:%M:%S",   # ISO, e.g. when Excel stores real datetimes
    "%Y-%m-%d %H:%M",
)

_TEXT_DATE_RE = re.compile(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})")


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def _warn(message):
    """Log a data-quality warning, and show it in the dashboard when running under Streamlit."""
    logger.warning(message)
    try:
        st.warning(message)
    except Exception:
        pass


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


def _read_csv_any_encoding(file_obj, **kwargs):
    """
    pd.read_csv that tolerates non-UTF-8 exports. Files re-saved from Excel as
    plain "CSV" are Windows-1252, which crashes pandas' default UTF-8 decoder
    as soon as one special character (°, é, –, ...) appears.
    """
    _seek0(file_obj)
    head = file_obj.read(2)
    if isinstance(head, bytes) and head in (b"\xff\xfe", b"\xfe\xff"):
        encodings = ("utf-16",)
    else:
        encodings = _CSV_ENCODINGS

    last_err = None
    for enc in encodings:
        _seek0(file_obj)
        try:
            return pd.read_csv(file_obj, encoding=enc, **kwargs)
        except UnicodeError as e:
            last_err = e
    raise last_err


def _read_excel_file(file_obj):
    """pd.ExcelFile with a clear message when the .xls reader (xlrd) is missing."""
    _seek0(file_obj)
    try:
        return pd.ExcelFile(file_obj)
    except ImportError as e:
        if _file_ext(file_obj) == ".xls":
            raise ImportError(
                "Reading .xls files needs the 'xlrd' package. Install it with: pip install xlrd"
            ) from e
        raise


def _coerce_dates(series):
    """
    Convert a logbook date column to date objects. Handles real datetimes and
    text like 'Kamis. 30/07/26' (day name + dd/mm/yy). Unparseable cells become NaT.
    """
    def one(value):
        if isinstance(value, pd.Timestamp):
            return value.normalize()
        if hasattr(value, "year") and hasattr(value, "month"):
            return pd.Timestamp(value).normalize()
        if isinstance(value, str):
            m = _TEXT_DATE_RE.search(value)
            if m:
                d, mo, y = (int(x) for x in m.groups())
                y = y + 2000 if y < 100 else y
                try:
                    return pd.Timestamp(year=y, month=mo, day=d)
                except ValueError:
                    return pd.NaT
        return pd.NaT

    return pd.to_datetime(series.map(one), errors="coerce")


def _parse_gps_time(series):
    """
    Parse the 'GPS Time' column, tolerating rows that Excel reformatted
    (e.g. '1/10/2026 4:21' instead of '01/10/2026 04:21:00').
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    s = series.astype(str).str.strip()
    result = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for fmt in _GPS_TIME_FORMATS:
        missing = result.isna()
        if not missing.any():
            break
        result = result.fillna(pd.to_datetime(s[missing], format=fmt, errors="coerce"))
    return result


# ---------------------------------------------------------------------------
# Logbook
# ---------------------------------------------------------------------------
def _clean_logbook_df(df):
    """Shared cleanup applied to a raw logbook table, regardless of CSV or Excel source."""
    if df.shape[1] < 7:
        return None
    cols = ["No", "Tujuan", "Berangkat_Tgl", "Berangkat_KM", "Tiba_Tgl",
            "Tiba_KM", "Jumlah_Pemakaian", "Refuel_L", "Fuel_Consum", "Keterangan"]
    df = df.iloc[:, : len(cols)].copy()
    df.columns = cols[: df.shape[1]]
    df = df[pd.to_numeric(df["No"], errors="coerce").notna()].copy()
    if df.empty:
        return None

    parsed_dates = _coerce_dates(df["Berangkat_Tgl"])
    bad = parsed_dates.isna()
    if bad.any():
        _warn(f"{int(bad.sum())} logbook row(s) had an unreadable date and were skipped.")
        df = df[~bad].copy()
        parsed_dates = parsed_dates[~bad]
        if df.empty:
            return None

    df["Date"] = parsed_dates.dt.date
    df["Jumlah_Pemakaian"] = pd.to_numeric(df["Jumlah_Pemakaian"], errors="coerce").fillna(0)
    df["Tujuan"] = df["Tujuan"].fillna("")
    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def parse_logbook(file_bytes):
    """
    Parse a logbook file into {plate: DataFrame}.
    - .xlsx/.xls: one sheet per vehicle, plate taken from the sheet name.
    - .csv: a single vehicle's logbook, plate taken from the filename (e.g.
      'BM 1356 CW.csv' -> 'BM 1356 CW'), since CSV has no concept of sheets.
      Upload one CSV per vehicle.
    """
    result = {}

    if _is_csv(file_bytes):
        plate = _plate_from_filename(file_bytes)
        df = _read_csv_any_encoding(file_bytes, header=0, skiprows=[1])
        cleaned = _clean_logbook_df(df)
        if cleaned is not None:
            result[plate] = cleaned
        return result

    # Excel: one sheet per vehicle
    xls = _read_excel_file(file_bytes)
    for sheet in xls.sheet_names:
        plate = norm_plate(sheet)
        df = pd.read_excel(xls, sheet_name=sheet, header=0, skiprows=[1])
        cleaned = _clean_logbook_df(df)
        if cleaned is not None:
            result[plate] = cleaned
    return result


# ---------------------------------------------------------------------------
# GPS
# ---------------------------------------------------------------------------
def _find_gps_header_row(raw):
    """
    Locate the row holding the real column headers: the first row (within the
    top 10) that has a cell exactly equal to 'GPS Time'. The 'GPS Time (UTC)'
    column sits in the same row, so it must not be used to rule a row out.
    """
    for i in range(min(10, len(raw))):
        cells = raw.iloc[i].astype(str).str.strip()
        if (cells == "GPS Time").any():
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
        raw = _read_csv_any_encoding(file_bytes, header=None)
    else:
        xls = _read_excel_file(file_bytes)
        sheet = next((s for s in xls.sheet_names if "report" in s.lower()), xls.sheet_names[0])
        raw = pd.read_excel(xls, sheet_name=sheet, header=None)

    header_row = _find_gps_header_row(raw)

    if is_csv:
        df = _read_csv_any_encoding(file_bytes, header=header_row)
    else:
        df = pd.read_excel(xls, sheet_name=sheet, header=header_row)

    veh_fallback = norm_plate(str(raw.iloc[0, 1])) if raw.shape[0] > 0 and raw.shape[1] > 1 else ""
    if not veh_fallback:
        veh_fallback = _plate_from_filename(file_bytes)  # e.g. 'History_BM_1356_CW.csv'

    if "Vehicle Code" in df.columns:
        df["Plate"] = df["Vehicle Code"].apply(lambda x: norm_plate(str(x)) or veh_fallback)
    else:
        df["Plate"] = veh_fallback

    df["GPS Time"] = _parse_gps_time(df["GPS Time"])
    n_bad = int(df["GPS Time"].isna().sum())
    if n_bad:
        name = getattr(file_bytes, "name", "GPS file")
        _warn(f"{name}: {n_bad} of {len(df)} GPS row(s) had an unreadable time and were ignored.")
    df = df[df["GPS Time"].notna()].copy()
    df["Date"] = df["GPS Time"].dt.date

    # Excel re-saves can turn numeric columns into text; coerce so aggregation never breaks.
    for col in ("Mileage(KM)", "Speed (Km/hr)"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    if "ACC" in df.columns:
        df["ACC"] = df["ACC"].astype(str).str.strip().str.upper()

    # Some GPS exports include an address/location column (name varies by
    # tracker platform, e.g. "Location", "Location Get All Locations").
    # Normalize it to a plain "Location" column for downstream code.
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


def compute_avg_moving_speed(raw_df, min_speed=20.0, dates=None):
    """
    Average speed (km/h) over raw GPS pings, counting only pings at or above
    `min_speed` (default 20 km/h) so parked / crawling / GPS-jitter pings don't
    drag the average down. Optionally restrict to a collection of dates.
    Returns None if there is no usable data.
    """
    if raw_df is None or raw_df.empty or "Speed (Km/hr)" not in raw_df.columns:
        return None
    df = raw_df
    if dates is not None:
        df = df[df["GPS Time"].dt.date.isin(set(dates))]
    speeds = pd.to_numeric(df["Speed (Km/hr)"], errors="coerce")
    speeds = speeds[speeds >= min_speed]
    return float(speeds.mean()) if not speeds.empty else None