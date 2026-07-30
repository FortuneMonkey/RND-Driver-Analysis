import re
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
