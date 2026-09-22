import re
import datetime as dt

import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy as copy_style

from modules.parsing import norm_plate

HEADER_ROWS = 5     # rows 1-5: the fixed "Vechicle Log Book" info block
DATA_START_ROW = 6  # row 6 onward: the actual trip rows
N_DATA_COLS = 10     # No, Tujuan, Berangkat_Tgl, Berangkat_KM, Tiba_Tgl,
                      # Tiba_KM, Jumlah_Pemakaian, Refuel_L, Fuel_Consum, Keterangan

_DATE_RE = re.compile(r"(\d{1,2})[/.](\d{1,2})[/.](\d{2,4})")


def _coerce_date(value):
    """Accepts a real datetime/date, or messy text like 'Kamis. 30/07/26'."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        m = _DATE_RE.search(value)
        if m:
            d, mo, y = (int(x) for x in m.groups())
            y = y + 2000 if y < 100 else y
            try:
                return dt.date(y, mo, d)
            except ValueError:
                return None
    return None


def _coerce_number(value, default=None):
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s in ("", "-"):
        return default
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return default


def _copy_cell_style(src_cell, dst_cell):
    dst_cell.font = copy_style(src_cell.font)
    dst_cell.fill = copy_style(src_cell.fill)
    dst_cell.border = copy_style(src_cell.border)
    dst_cell.alignment = copy_style(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format


def _copy_header_block(src_ws, dst_ws):
    """Copy rows 1..HEADER_ROWS verbatim (values + basic styling + column widths)."""
    for row in range(1, HEADER_ROWS + 1):
        for col in range(1, N_DATA_COLS + 1):
            src_cell = src_ws.cell(row=row, column=col)
            dst_cell = dst_ws.cell(row=row, column=col, value=src_cell.value)
            _copy_cell_style(src_cell, dst_cell)
    for col in range(1, N_DATA_COLS + 1):
        letter = get_column_letter(col)
        if letter in src_ws.column_dimensions:
            dst_ws.column_dimensions[letter].width = src_ws.column_dimensions[letter].width


def clean_vehicle_rows(src_ws, period_start, period_end, warnings):
    """
    Reads data rows (row 6+) from src_ws and returns a list of cleaned rows:
    [No, Tujuan, Berangkat_Tgl, Berangkat_KM, Tiba_Tgl, Tiba_KM,
     Jumlah_Pemakaian, Refuel_L, Fuel_Consum, Keterangan]

    - Keeps only rows whose Berangkat date falls within [period_start, period_end]
      (the 26th-to-25th billing cycle).
    - Renumbers `No` sequentially from 1.
    - For a zero-usage row (Jumlah_Pemakaian == 0) with a placeholder/zero KM
      reading, both KM cells are replaced with the previous row's ending KM.
    """
    cleaned = []
    last_km = None
    no_counter = 0

    for row in src_ws.iter_rows(min_row=DATA_START_ROW, max_col=N_DATA_COLS, values_only=True):
        if row is None:
            continue
        row = list(row) + [None] * (N_DATA_COLS - len(row))
        (_no, tujuan, ber_tgl, ber_km, tib_tgl, tib_km,
         jumlah, refuel, fuelc, ket) = row[:N_DATA_COLS]

        ber_date = _coerce_date(ber_tgl)
        if ber_date is None:
            continue  # not a real data row
        if not (period_start <= ber_date <= period_end):
            continue  # outside this billing cycle

        jumlah_num = _coerce_number(jumlah, default=0.0)
        ber_km_num = _coerce_number(ber_km)
        tib_km_num = _coerce_number(tib_km)

        if jumlah_num == 0:
            if last_km is not None:
                ber_km_num, tib_km_num = last_km, last_km
            else:
                warnings.append(
                    f"{src_ws.title}: zero-usage row on {ber_date} is the first row in "
                    "the period, so there's no prior KM to carry forward — left as-is."
                )

        if tib_km_num is not None:
            last_km = tib_km_num
        elif ber_km_num is not None:
            last_km = ber_km_num

        no_counter += 1
        cleaned.append([
            no_counter, tujuan or "-", ber_date, ber_km_num, ber_date, tib_km_num,
            jumlah_num, refuel, fuelc, ket,
        ])

    return cleaned


def build_clean_workbook(raw_path, target_vehicles, period_start, period_end):
    """
    Reads `raw_path`, keeps only the sheets matching `target_vehicles`
    (matched against either the sheet name or the 'No Plat Pol.' cell,
    both normalized with norm_plate), cleans each one, and returns
    (openpyxl.Workbook, {plate: found_bool}, warnings_list).
    """
    src_wb = openpyxl.load_workbook(raw_path, data_only=True)
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)  # drop the default blank sheet

    found = {plate: False for plate in target_vehicles}
    warnings = []
    target_set = set(target_vehicles)

    for sheet_name in src_wb.sheetnames:
        src_ws = src_wb[sheet_name]
        plate_cell = str(src_ws.cell(row=1, column=6).value or "").lstrip(":").strip()
        candidates = {norm_plate(sheet_name), norm_plate(plate_cell)}
        match = candidates & target_set
        if not match:
            continue
        plate = next(iter(match))
        if found[plate]:
            warnings.append(f"Multiple sheets matched {plate}; keeping the first one found.")
            continue
        found[plate] = True

        dst_ws = out_wb.create_sheet(title=plate)
        _copy_header_block(src_ws, dst_ws)

        rows = clean_vehicle_rows(src_ws, period_start, period_end, warnings)
        style_row = DATA_START_ROW if src_ws.max_row >= DATA_START_ROW else None
        for i, row_vals in enumerate(rows):
            r = DATA_START_ROW + i
            for c, val in enumerate(row_vals, start=1):
                cell = dst_ws.cell(row=r, column=c, value=val)
                if style_row:
                    _copy_cell_style(src_ws.cell(row=style_row, column=c), cell)
            # dates as real datetimes so Excel/pandas read them cleanly either way
            dst_ws.cell(row=r, column=3).number_format = "dd/mm/yyyy"
            dst_ws.cell(row=r, column=5).number_format = "dd/mm/yyyy"

        if not rows:
            warnings.append(f"{plate}: no data rows fell inside {period_start}..{period_end}.")

    for plate, ok in found.items():
        if not ok:
            warnings.append(f"{plate}: no matching sheet found in {raw_path}.")

    return out_wb, found, warnings
