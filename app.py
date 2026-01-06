from fastapi import FastAPI, Request, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from pathlib import Path
import pandas as pd
import uvicorn
from datetime import datetime, timedelta
import csv
import os
import warnings
from io import StringIO, BytesIO

warnings.filterwarnings("ignore")

# =========================================================
# RENDER / PATH CONFIG (WORKS: Render + Local Windows)
# =========================================================
IS_RENDER = os.getenv("RENDER", "").lower() == "true"

# Render Disk default: /mnt/data  (set in render.yaml)
DATA_DIR = Path(os.getenv("DATA_DIR", "/mnt/data" if IS_RENDER else ".")).resolve()

# Optional folders (good for uploads)
INPUT_DIR = Path(os.getenv("INPUT_DIR", str(DATA_DIR / "input"))).resolve()
EXPORT_DIR = Path(os.getenv("EXPORT_DIR", str(DATA_DIR / "exports"))).resolve()

# Repo folder (where app.py exists on Render build)
BASE_DIR = Path(__file__).resolve().parent

# IMPORTANT:
# Your .xls files may be in GitHub repo root OR uploaded to Render Disk.
# We will search BOTH.
CANDIDATE_INPUT_DIRS = [
    INPUT_DIR,    # /mnt/data/input  (upload API)
    BASE_DIR,     # repo folder (GitHub committed files)
]

# DATE FILE CONFIG
DATE_FILE = Path(os.getenv("DATE_FILE", str(DATA_DIR / "Date.xlsx"))).resolve()
if not DATE_FILE.exists():
    DATE_FILE = BASE_DIR / "Date.xlsx"

# =========================================================
# YOUR ORIGINAL SETTINGS (UNCHANGED)
# =========================================================
FILE_MAPPINGS = {
    "_AMRAVATI_STOCK_Summary_report.xls": "AMT",
    "_CHIKHALI_STOCK_Summary_report.xls": "CHI",
    "_NAGPUR_WARDHAMANNGR_STOCK_Summary_report.xls": "CITY",
    "_NAGPUR_KAMPTHEEROAD_STOCK_Summary_report.xls": "HO",
    "_YAVATMAL_STOCK_Summary_report.xls": "YAT",
    "_WAGHOLI_STOCK_Summary_report.xls": "WAG",
    "_CHAUFULA_SZZ_STOCK_Summary_report.xls": "CHA",
    "_SHIKRAPUR_SZS_STOCK_Summary_report.xls": "SHI",
    "_KOLHAPUR_WS_STOCK_Summary_report.xls": "KOL",
}

SHEET_NAMES = [
    "Overall_Division_Pivot",
    "Cat_Spares",
    "ABC_Analysis",
    "RIS_Analysis",
    "Cat_Accessories",
    "Cat_Battery",
    "Cat_CONSUMABLE",
    "Cat_Local Accessories",
    "Cat_Local Parts",
    "Cat_Maxicare Products",
    "Cat_Maxiclean",
    "Cat_Maximile Coolant",
    "Cat_Maximile Grease",
    "Cat_Maximile Oil",
    "Cat_Maximile Oil Filter",
    "Cat_OIL & AC GAS",
    "Cat_Tools",
    "Cat_Tyres",
]

RIS_CATEGORY_SHEETS = ["Cat_Spares", "Cat_Accessories", "Cat_Local Accessories"]

_data_cache = None
_date_cache = None

app = FastAPI(title="Unnati Stock Kundali (Render Ready)")

# =========================================================
# DATE FUNCTIONS (LOAD ONLY - NO SAVE)
# =========================================================

def validate_date_format(date_str: str) -> bool:
    """
    Validate date is in DD-MM-YYYY format
    
    Conditions checked:
    - Format must be DD-MM-YYYY
    - DD must be 01-31
    - MM must be 01-12
    - YYYY must be valid year
    
    Returns: True if valid, False otherwise
    """
    try:
        if not date_str or len(str(date_str).strip()) == 0:
            return False
        
        date_str = str(date_str).strip()
        
        if len(date_str) != 10:
            return False
        
        parts = date_str.split('-')
        if len(parts) != 3:
            return False
        
        day, month, year = parts
        
        # Check format
        if not day.isdigit() or not month.isdigit() or not year.isdigit():
            return False
        
        day_int = int(day)
        month_int = int(month)
        year_int = int(year)
        
        # Validate ranges
        if day_int < 1 or day_int > 31:
            return False
        if month_int < 1 or month_int > 12:
            return False
        if year_int < 2000 or year_int > 2099:
            return False
        
        # Try to parse as date
        pd.to_datetime(date_str, format="%d-%m-%Y")
        return True
    except Exception:
        return False


def load_dates():
    """
    Load dates from Date.xlsx on app startup
    
    Conditions:
    1. Check if file exists
    2. Check if Sheet1 exists
    3. Check if columns exist
    4. Validate date values
    5. Convert to DD-MM-YYYY format
    6. Fall back to today's date if any error
    
    Returns: dict with dates
    """
    global _date_cache
    
    try:
        if not DATE_FILE.exists():
            print(f"[DATE WARNING] Date file not found: {DATE_FILE}")
            raise FileNotFoundError(f"Date file not found at {DATE_FILE}")
        
        # Read Excel file
        df = pd.read_excel(DATE_FILE, sheet_name='Sheet1')
        
        # Check if has data
        if len(df) == 0:
            print("[DATE ERROR] Date file is empty")
            raise ValueError("Date file is empty")
        
        # Check if columns exist
        if 'Open Stock Date' not in df.columns or 'Close Stock Date' not in df.columns:
            print("[DATE ERROR] Required columns not found. Expected: 'Open Stock Date', 'Close Stock Date'")
            raise KeyError("Required columns not found in Date.xlsx")
        
        # Get first row values
        open_date = df.iloc[0]['Open Stock Date']
        close_date = df.iloc[0]['Close Stock Date']
        
        # Convert to datetime if needed
        if not isinstance(open_date, pd.Timestamp):
            open_date = pd.to_datetime(open_date)
        if not isinstance(close_date, pd.Timestamp):
            close_date = pd.to_datetime(close_date)
        
        # Format as DD-MM-YYYY
        open_date_str = open_date.strftime("%d-%m-%Y")
        close_date_str = close_date.strftime("%d-%m-%Y")
        
        # Validate format
        if not validate_date_format(open_date_str) or not validate_date_format(close_date_str):
            print(f"[DATE ERROR] Date format validation failed")
            raise ValueError("Invalid date format in file")
        
        # Validate range
        open_dt = pd.to_datetime(open_date_str, format="%d-%m-%Y")
        close_dt = pd.to_datetime(close_date_str, format="%d-%m-%Y")
        if close_dt < open_dt:
            print(f"[DATE ERROR] Close date is before open date")
            raise ValueError("Close date must be >= Open date")
        
        _date_cache = {
            "open_date": open_date_str,
            "close_date": close_date_str,
            "open_date_raw": open_date,
            "close_date_raw": close_date,
        }
        
        print(f"[DATE LOADED ✓] Open: {_date_cache['open_date']}, Close: {_date_cache['close_date']}")
        return _date_cache
        
    except Exception as e:
        print(f"[DATE ERROR] {str(e)}")
        # Fallback: use today's date
        today = datetime.now()
        today_str = today.strftime("%d-%m-%Y")
        _date_cache = {
            "open_date": today_str,
            "close_date": today_str,
            "open_date_raw": today,
            "close_date_raw": today,
        }
        print(f"[DATE FALLBACK] Using today's date: {today_str}")
        return _date_cache


def get_dates():
    """
    Get current dates from cache or load from file
    
    Conditions:
    1. Check if cache exists
    2. Load from file if not cached
    3. Return date dict
    """
    global _date_cache
    if _date_cache is None:
        return load_dates()
    return _date_cache


# =========================================================
# FILE READING (UNCHANGED LOGIC)
# =========================================================
def detect_format(file_path: str) -> str:
    try:
        with open(file_path, "rb") as f:
            header = f.read(2000)

        header_str = header.decode("utf-8", errors="ignore").lower()
        if "<html" in header_str or "<table" in header_str or "<!doctype" in header_str:
            return "HTML"

        if b"\t" in header:
            return "TSV"

        if header[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
            return "XLS"

        if header[:2] == b"PK":
            return "XLSX"

        return "UNKNOWN"
    except:
        return "UNKNOWN"


def read_html_file(file_path: str):
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            html_content = f.read()
        tables = pd.read_html(StringIO(html_content))
        if tables and len(tables) > 0 and len(tables[0]) > 0:
            return tables[0], "HTML"
    except:
        pass
    return None, None


def read_tsv_file(file_path: str):
    try:
        df = pd.read_csv(file_path, sep="\t")
        if df is not None and len(df) > 0:
            return df, "TSV"
    except:
        pass
    return None, None


def read_xls_file(file_path: str):
    for engine in ["xlrd", "openpyxl", None]:
        try:
            df = pd.read_excel(file_path, engine=engine)
            if df is not None and len(df) > 0:
                return df, f"XLS ({engine or 'default'})"
        except:
            pass
    return None, None


def read_file(file_path: str):
    file_format = detect_format(file_path)

    if file_format == "HTML":
        return read_html_file(file_path)
    if file_format == "TSV":
        return read_tsv_file(file_path)
    if file_format in ["XLS", "XLSX"]:
        return read_xls_file(file_path)

    for df, method in [read_tsv_file(file_path), read_xls_file(file_path), read_html_file(file_path)]:
        if df is not None and len(df) > 0:
            return df, method

    return None, None


def _find_input_file(filename: str):
    for d in CANDIDATE_INPUT_DIRS:
        p = d / filename
        if p.exists():
            return p
    return None


def load_files():
    global _data_cache

    print("\n" + "=" * 80)
    print("LOADING AND MERGING FILES (Render Ready)")
    print("=" * 80)
    print(f"DATA_DIR  : {DATA_DIR}")
    print(f"INPUT_DIR : {INPUT_DIR}")
    print(f"BASE_DIR  : {BASE_DIR}")
    print("=" * 80)

    all_data = []
    file_count = 0

    for idx, (filename, division) in enumerate(FILE_MAPPINGS.items(), 1):
        file_path = _find_input_file(filename)

        if file_path is None:
            print(f"[{idx}/{len(FILE_MAPPINGS)}]  {filename}... [MISSING]")
            continue

        df, method = read_file(str(file_path))
        if df is None or len(df) == 0:
            print(f"[{idx}/{len(FILE_MAPPINGS)}]  {filename}... [FAILED]")
            continue

        df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        df.insert(0, "Division", division)

        all_data.append(df)
        file_count += 1
        print(f"[{idx}/{len(FILE_MAPPINGS)}]  {filename}... [OK] ({method}, {len(df):,} rows)")

    if not all_data:
        _data_cache = None
        print("\n[WARNING] No valid files loaded.")
        return None

    combined_df = pd.concat(all_data, ignore_index=True)
    _data_cache = combined_df
    print(f"\n[OK] Total merged rows: {len(combined_df):,}")
    print(f"[OK] Total columns: {len(combined_df.columns)}")
    return combined_df


def get_data():
    global _data_cache
    if _data_cache is None:
        return load_files()
    return _data_cache


# =========================================================
# HELPERS (UNCHANGED + LOCATION SUPPORT)
# =========================================================
def format_indian_number(num, decimal_places=0):
    if pd.isna(num):
        return ""
    try:
        num = float(num)
    except:
        return ""

    is_negative = num < 0
    num = abs(num)

    if decimal_places == 0:
        integer_part = f"{num:.0f}"
        decimal_part = ""
    else:
        s = f"{num:.{decimal_places}f}"
        if "." in s:
            integer_part, decimal_part = s.split(".")
        else:
            integer_part, decimal_part = s, "0" * decimal_places

    if len(integer_part) <= 3:
        formatted_integer = integer_part
    else:
        last_three = integer_part[-3:]
        remaining = integer_part[:-3]
        parts = []
        while len(remaining) > 2:
            parts.insert(0, remaining[-2:])
            remaining = remaining[:-2]
        if remaining:
            parts.insert(0, remaining)
        formatted_integer = ",".join(parts) + "," + last_three

    result = formatted_integer if decimal_places == 0 else f"{formatted_integer}.{decimal_part}"
    if is_negative:
        result = "-" + result
    return result


def get_divisions(sheet_name=None):
    df = get_data()
    if df is None or "PRODCT_DIVSN_DESC" not in df.columns:
        return []

    if sheet_name and sheet_name.startswith("Cat_"):
        category_name = sheet_name.replace("Cat_", "").replace("_", " ")
        if "PART_CATGRY_DESC" in df.columns:
            df = df[df["PART_CATGRY_DESC"].astype(str).str.contains(category_name, case=False, na=False)]

    return sorted(df["PRODCT_DIVSN_DESC"].dropna().astype(str).unique().tolist())


def get_locations():
    df = get_data()
    if df is None or "Division" not in df.columns:
        return sorted(set(FILE_MAPPINGS.values()))
    return sorted(df["Division"].dropna().astype(str).unique().tolist())


def filter_data(divisions=None, locations=None):
    df = get_data()
    if df is None:
        return None

    if divisions and len(divisions) > 0 and divisions != ["All"]:
        if "PRODCT_DIVSN_DESC" in df.columns:
            df = df[df["PRODCT_DIVSN_DESC"].astype(str).isin([str(x) for x in divisions])]

    if locations and len(locations) > 0 and locations != ["All"]:
        if "Division" in df.columns:
            df = df[df["Division"].astype(str).isin([str(x) for x in locations])]

    return df


# =========================================================
# PIVOT LOGIC (UNCHANGED)
# =========================================================
def create_standard_pivot(df):
    if df is None or len(df) == 0 or "Division" not in df.columns:
        return None

    value_cols = [
        "OPEN_QTY", "OPEN_VAL",
        "RECPT_QTY", "RECPT_VAL",
        "ISS_QTY", "ISS_VAL",
        "CLOSE_QTY", "CLOSE_VAL",
    ]
    available_cols = [c for c in value_cols if c in df.columns]
    if not available_cols:
        return None

    pivot_data = []
    for division in sorted(df["Division"].dropna().astype(str).unique()):
        div_data = df[df["Division"].astype(str) == division]
        row = {"Division": division}

        for col in available_cols:
            if "QTY" in col:
                count = (pd.to_numeric(div_data[col], errors="coerce").fillna(0) != 0).sum()
                row[col.replace("QTY", "Count")] = int(count)
            else:
                row[col] = pd.to_numeric(div_data[col], errors="coerce").fillna(0).sum()

        if {"OPEN_QTY", "RECPT_QTY", "ISS_QTY"}.issubset(div_data.columns):
            oq = pd.to_numeric(div_data["OPEN_QTY"], errors="coerce").fillna(0)
            rq = pd.to_numeric(div_data["RECPT_QTY"], errors="coerce").fillna(0)
            iq = pd.to_numeric(div_data["ISS_QTY"], errors="coerce").fillna(0)
            row["Part Sold From Open Stock Count"] = int(((oq > 0) & (rq == 0) & (iq > 0)).sum())

        if {"OPEN_VAL", "RECPT_VAL", "ISS_VAL"}.issubset(div_data.columns):
            ov = pd.to_numeric(div_data["OPEN_VAL"], errors="coerce").fillna(0)
            rv = pd.to_numeric(div_data["RECPT_VAL"], errors="coerce").fillna(0)
            iv = pd.to_numeric(div_data["ISS_VAL"], errors="coerce").fillna(0)
            row["Part Sold From Open Stock Value"] = max(0, iv[(ov > 0) & (rv == 0) & (iv > 0)].sum())

        if {"RECPT_QTY", "ISS_QTY"}.issubset(div_data.columns):
            rq = pd.to_numeric(div_data["RECPT_QTY"], errors="coerce").fillna(0)
            iq = pd.to_numeric(div_data["ISS_QTY"], errors="coerce").fillna(0)
            row["Part Sold From Recpt Count"] = int(((rq > 0) & (iq > 0)).sum())

        if {"RECPT_VAL", "ISS_VAL"}.issubset(div_data.columns):
            rv = pd.to_numeric(div_data["RECPT_VAL"], errors="coerce").fillna(0)
            iv = pd.to_numeric(div_data["ISS_VAL"], errors="coerce").fillna(0)
            row["Part Sold From Recpt Value"] = max(0, iv[(rv > 0) & (iv > 0)].sum())

        pivot_data.append(row)

    pivot_df = pd.DataFrame(pivot_data)

    if "OPEN_Count" in pivot_df.columns and "CLOSE_Count" in pivot_df.columns:
        pivot_df["Change Count"] = pivot_df["CLOSE_Count"] - pivot_df["OPEN_Count"]
    if "OPEN_VAL" in pivot_df.columns and "CLOSE_VAL" in pivot_df.columns:
        pivot_df["Change Value"] = pivot_df["CLOSE_VAL"] - pivot_df["OPEN_VAL"]

    total_row = {"Division": "Total"}
    for col in pivot_df.columns:
        if col != "Division":
            total_row[col] = pd.to_numeric(pivot_df[col], errors="coerce").fillna(0).sum()
    pivot_df = pd.concat([pivot_df, pd.DataFrame([total_row])], ignore_index=True)

    return pivot_df


def create_abc_analysis(df):
    if df is None or len(df) == 0:
        return None
    if "ABC_IND" not in df.columns:
        return create_standard_pivot(df)

    abc_values = ["A", "B", "C1", "C2"]
    result = {"type": "multiple_tables", "tables": []}

    for abc_val in abc_values:
        abc_data = df[df["ABC_IND"].astype(str).str.strip() == abc_val]
        if len(abc_data) > 0:
            pivot = create_standard_pivot(abc_data)
            if pivot is not None:
                result["tables"].append({"title": f"ABC {abc_val}", "data": pivot})

    return result if result["tables"] else create_standard_pivot(df)


def create_ris_analysis(df):
    if df is None or len(df) == 0:
        return None
    if "RIS_IND" not in df.columns:
        return create_standard_pivot(df)

    ris_values = [
        {"code": "R", "display": "Regular (R)"},
        {"code": "I", "display": "Intermediate (I)"},
        {"code": "S", "display": "Slow (S)"},
    ]
    result = {"type": "multiple_tables", "tables": []}

    for ris_info in ris_values:
        ris_data = df[df["RIS_IND"].astype(str).str.strip() == ris_info["code"]]
        if len(ris_data) > 0:
            pivot = create_standard_pivot(ris_data)
            if pivot is not None:
                result["tables"].append({"title": ris_info["display"], "data": pivot})

    return result if result["tables"] else create_standard_pivot(df)


def create_category_ris_tables(df, sheet_name):
    if df is None or len(df) == 0:
        return None
    if "RIS_IND" not in df.columns:
        return create_standard_pivot(df)

    ris_values = [
        {"code": "R", "display": "Regular (R)"},
        {"code": "I", "display": "Intermediate (I)"},
        {"code": "S", "display": "Slow (S)"},
    ]
    result = {"type": "multiple_tables", "tables": []}

    all_pivot = create_standard_pivot(df)
    if all_pivot is not None:
        result["tables"].append({"title": f"{sheet_name.replace('_',' ')} - All", "data": all_pivot})

    for ris_info in ris_values:
        ris_data = df[df["RIS_IND"].astype(str).str.strip() == ris_info["code"]]
        if len(ris_data) > 0:
            pivot = create_standard_pivot(ris_data)
            if pivot is not None:
                result["tables"].append({"title": ris_info["display"], "data": pivot})

    return result if result["tables"] else create_standard_pivot(df)


def create_pivot(df, sheet_name):
    if df is None or len(df) == 0:
        return None

    if sheet_name == "Overall_Division_Pivot":
        return create_standard_pivot(df)

    if sheet_name == "ABC_Analysis":
        return create_abc_analysis(df)

    if sheet_name == "RIS_Analysis":
        return create_ris_analysis(df)

    if sheet_name.startswith("Cat_"):
        category_name = sheet_name.replace("Cat_", "").replace("_", " ")
        if "PART_CATGRY_DESC" in df.columns:
            df = df[df["PART_CATGRY_DESC"].astype(str).str.contains(category_name, case=False, na=False)]

        if sheet_name in RIS_CATEGORY_SHEETS:
            return create_category_ris_tables(df, sheet_name)

        return create_standard_pivot(df)

    return create_standard_pivot(df)


# =========================================================
# TABLE HTML (UNCHANGED STYLE + SAFE FORMAT)
# =========================================================
def format_value(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (int, float)):
        if value < 0:
            return f'<span class="neg">({format_indian_number(abs(value), 0)})</span>'
        return format_indian_number(value, 0)
    return str(value)


def table_to_html(pivot_result):
    try:
        if isinstance(pivot_result, dict) and pivot_result.get("type") == "multiple_tables":
            html_parts = []
            for table_info in pivot_result["tables"]:
                title = table_info["title"]
                df = table_info["data"]
                if df is None or len(df) == 0:
                    continue

                html_parts.append('<section class="tbl-group">')
                html_parts.append(f'<div class="tbl-title">{title}</div>')
                html_parts.append('<div class="tbl-wrap">')
                html_parts.append('<table class="pivot">')

                html_parts.append('<thead><tr>')
                for col in df.columns:
                    html_parts.append(f'<th>{col}</th>')
                html_parts.append('</tr></thead>')

                html_parts.append('<tbody>')
                for _, row in df.iterrows():
                    is_total = str(row.iloc[0]).strip().lower() == "total"
                    tr_class = "total" if is_total else ""
                    html_parts.append(f'<tr class="{tr_class}">')
                    for col_idx, value in enumerate(row):
                        td_class = "rowlabel" if col_idx == 0 else "num"
                        html_parts.append(f'<td class="{td_class}">{format_value(value)}</td>')
                    html_parts.append("</tr>")
                html_parts.append("</tbody>")

                html_parts.append("</table></div></section>")

            return "".join(html_parts) if html_parts else '<div class="empty">No data available</div>'

        df = pivot_result
        if df is None or len(df) == 0:
            return '<div class="empty">No data available</div>'

        html = ['<div class="tbl-wrap"><table class="pivot">']
        html.append("<thead><tr>")
        for col in df.columns:
            html.append(f"<th>{col}</th>")
        html.append("</tr></thead><tbody>")

        for _, row in df.iterrows():
            is_total = str(row.iloc[0]).strip().lower() == "total"
            tr_class = "total" if is_total else ""
            html.append(f'<tr class="{tr_class}">')
            for col_idx, value in enumerate(row):
                td_class = "rowlabel" if col_idx == 0 else "num"
                html.append(f'<td class="{td_class}">{format_value(value)}</td>')
            html.append("</tr>")

        html.append("</tbody></table></div>")
        return "".join(html)

    except Exception as e:
        return f'<div class="empty">Error: {e}</div>'


def safe_title(sheet: str) -> str:
    return sheet.replace("_", " ").replace(" Pivot", "").strip()


# =========================================================
# EXPORT HELPERS (PIVOT CSV + RAW CSV)
# =========================================================
def pivot_to_csv_bytes(pivot_result) -> bytes:
    buf = StringIO()
    writer = csv.writer(buf)

    if isinstance(pivot_result, dict) and pivot_result.get("type") == "multiple_tables":
        for i, table_info in enumerate(pivot_result["tables"]):
            title = table_info["title"]
            df_table = table_info["data"]
            if df_table is None or len(df_table) == 0:
                continue

            writer.writerow([title] + [""] * (max(0, len(df_table.columns) - 1)))
            writer.writerow(list(df_table.columns))

            for _, row in df_table.iterrows():
                out = []
                for v in row:
                    if isinstance(v, (int, float)):
                        out.append(format_indian_number(v, 0))
                    else:
                        out.append(v)
                writer.writerow(out)

            if i < len(pivot_result["tables"]) - 1:
                writer.writerow([])
    else:
        df = pivot_result
        if df is None or len(df) == 0:
            writer.writerow(["No data"])
        else:
            writer.writerow(list(df.columns))
            for _, row in df.iterrows():
                out = []
                for v in row:
                    if isinstance(v, (int, float)):
                        out.append(format_indian_number(v, 0))
                    else:
                        out.append(v)
                writer.writerow(out)

    return buf.getvalue().encode("utf-8-sig")


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    if df is None or len(df) == 0:
        return "No data".encode("utf-8-sig")
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# =========================================================
# API: HEALTH / RELOAD / UPLOAD (RENDER READY)
# =========================================================
@app.get("/health")
async def health():
    df = get_data()
    dates = get_dates()
    return {
        "ok": True,
        "render": IS_RENDER,
        "data_dir": str(DATA_DIR),
        "input_dir": str(INPUT_DIR),
        "base_dir": str(BASE_DIR),
        "date_file": str(DATE_FILE),
        "candidate_input_dirs": [str(x) for x in CANDIDATE_INPUT_DIRS],
        "loaded": df is not None,
        "rows": int(len(df)) if df is not None else 0,
        "cols": int(len(df.columns)) if df is not None else 0,
        "dates": dates,
    }


@app.post("/reload")
async def reload_data():
    global _data_cache, _date_cache
    _data_cache = None
    _date_cache = None
    
    # Reload both data and dates
    df = load_files()
    dates = load_dates()
    
    if df is None:
        return {"status": "error", "message": "No data loaded. Check files exist in repo or upload to /upload."}
    return {
        "status": "success",
        "rows": int(len(df)),
        "cols": int(len(df.columns)),
        "dates": dates
    }


@app.post("/upload")
async def upload_files(files: list[UploadFile] = File(...)):
    """
    Upload your *_STOCK_Summary_report.xls files here.
    They will be saved into /mnt/data/input on Render Disk.
    Then call /reload.
    """
    INPUT_DIR.mkdir(parents=True, exist_ok=True)

    saved = []
    skipped = []
    for f in files:
        name = Path(f.filename).name

        # Allow only expected names (safe)
        if name not in FILE_MAPPINGS:
            skipped.append(name)
            continue

        content = await f.read()
        if not content:
            skipped.append(name)
            continue

        out = INPUT_DIR / name
        out.write_bytes(content)
        saved.append(name)

    if not saved:
        raise HTTPException(status_code=400, detail={"message": "No valid files uploaded", "skipped": skipped})

    return {"status": "success", "saved": saved, "skipped": skipped, "input_dir": str(INPUT_DIR)}


# =========================================================
# API: DATES (GET ONLY - NO SET)
# =========================================================
@app.get("/api/dates")
async def api_get_dates():
    """
    Get current Open Stock Date and Close Stock Date
    
    Dates are read from Date.xlsx file on app startup.
    To update: manually edit Date.xlsx and restart app.
    """
    dates = get_dates()
    return JSONResponse(content={
        "open_date": dates["open_date"],
        "close_date": dates["close_date"],
        "timestamp": datetime.now().isoformat(),
        "source": "Date.xlsx file (manual update required)"
    })


# =========================================================
# API: DIVISIONS
# =========================================================
@app.get("/api/product-divisions")
async def api_divisions(request: Request):
    sheet = request.query_params.get("sheet", None)
    divisions = get_divisions(sheet)
    return JSONResponse(content={"divisions": divisions})


# =========================================================
# DOWNLOAD ENDPOINTS (PIVOT + RAW)  (LOCATION ADDED)
# =========================================================
@app.get("/download")
async def download_csv(request: Request):
    sheet = request.query_params.get("sheet", SHEET_NAMES[0])
    divisions_param = request.query_params.get("divisions", "All")
    locations_param = request.query_params.get("locations", "All")

    divisions_list = [] if divisions_param == "All" else [d.strip() for d in divisions_param.split(",") if d.strip()]
    locations_list = [] if locations_param == "All" else [l.strip() for l in locations_param.split(",") if l.strip()]

    df = filter_data(
        divisions_list if divisions_list else None,
        locations_list if locations_list else None
    )

    pivot_result = create_pivot(df, sheet)
    if pivot_result is None:
        raise HTTPException(status_code=404, detail="No data to export")

    content = pivot_to_csv_bytes(pivot_result)
    now = datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
    filename = f"{sheet}_{now}.csv"

    return StreamingResponse(
        BytesIO(content),
        media_type="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/download-raw")
async def download_raw_csv(request: Request):
    divisions_param = request.query_params.get("divisions", "All")
    locations_param = request.query_params.get("locations", "All")

    divisions_list = [] if divisions_param == "All" else [d.strip() for d in divisions_param.split(",") if d.strip()]
    locations_list = [] if locations_param == "All" else [l.strip() for l in locations_param.split(",") if l.strip()]

    df = filter_data(
        divisions_list if divisions_list else None,
        locations_list if locations_list else None
    )
    if df is None or len(df) == 0:
        raise HTTPException(status_code=404, detail="No data to export")

    content = df_to_csv_bytes(df)
    now = datetime.now().strftime("%d_%m_%Y_%H_%M_%S")
    filename = f"RAW_DATA_{now}.csv"

    return StreamingResponse(
        BytesIO(content),
        media_type="text/csv",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# =========================================================
# DASHBOARD (WITH DATE DISPLAY - NO ADMIN BUTTON)
# =========================================================
@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    try:
        sheet = request.query_params.get("sheet", SHEET_NAMES[0])
        divisions_param = request.query_params.get("divisions", "All")
        locations_param = request.query_params.get("locations", "All")

        # Product Divisions selection
        if divisions_param == "All":
            divisions_list = []
            selected_divisions = []
        else:
            decoded = divisions_param
            divisions_list = [d.strip() for d in decoded.split(",") if d.strip()]
            selected_divisions = divisions_list[:]

        # Locations selection
        if locations_param == "All":
            locations_list = []
            selected_locations = []
        else:
            decoded_l = locations_param
            locations_list = [l.strip() for l in decoded_l.split(",") if l.strip()]
            selected_locations = locations_list[:]

        all_divisions = get_divisions(sheet)
        all_locations = get_locations()
        dates = get_dates()

        buttons = []
        for s in SHEET_NAMES:
            active = "active" if s == sheet else ""
            buttons.append(
                f'<a class="navbtn {active}" href="/?sheet={s}&divisions={divisions_param}&locations={locations_param}">{safe_title(s)}</a>'
            )
        buttons_html = "".join(buttons)

        df = filter_data(
            divisions_list if divisions_list else None,
            locations_list if locations_list else None
        )
        pivot_result = create_pivot(df, sheet)
        table_html = table_to_html(pivot_result)

        divisions_display = "All Divisions" if not selected_divisions else ", ".join(selected_divisions)
        locations_display = "All Locations" if not selected_locations else ", ".join(selected_locations)

        def esc_attr(s: str) -> str:
            return str(s).replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")

        # Product Division checkbox list
        checkbox_rows = []
        for div in all_divisions:
            checked = "checked" if div in selected_divisions else ""
            v = esc_attr(div)
            checkbox_rows.append(f'<label class="chk"><input type="checkbox" value="{v}" {checked}><span>{v}</span></label>')
        checkbox_html = "".join(checkbox_rows)

        # Location checkbox list
        loc_checkbox_rows = []
        for loc in all_locations:
            checked = "checked" if loc in selected_locations else ""
            v = esc_attr(loc)
            loc_checkbox_rows.append(f'<label class="chk"><input type="checkbox" value="{v}" {checked}><span>{v}</span></label>')
        loc_checkbox_html = "".join(loc_checkbox_rows)

        html = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Unnati Stock Kundali</title>
<style>
:root {
  --bg: #f4f6fb;
  --card: #ffffff;
  --text: #0f172a;
  --muted: #64748b;
  --line: #e5e7eb;
  --brand1: #0b1b3a;
  --brand2: #123a7a;
  --btn: #0f4c81;
  --btn2: #0a6a3a;
  --danger: #b42318;
  --shadow: 0 10px 25px rgba(15, 23, 42, 0.08);
}

* { box-sizing: border-box; }
body {
  margin: 0;
  font-family: "Segoe UI", Tahoma, Arial, sans-serif;
  background: var(--bg);
  color: var(--text);
  height: 100vh;
  overflow: hidden;
}

.app {
  display: grid;
  grid-template-rows: auto 1fr;
  height: 100vh;
}

header {
  background: linear-gradient(135deg, var(--brand1), var(--brand2));
  color: #fff;
  padding: 10px 14px;
  box-shadow: var(--shadow);
}

.hrow {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  flex-wrap: wrap;
}

.brand {
  display: flex;
  flex-direction: column;
  min-width: 260px;
}
.brand .title {
  font-size: 18px;
  font-weight: 800;
  letter-spacing: .2px;
}
.brand .sub {
  font-size: 12px;
  color: rgba(255,255,255,.75);
  margin-top: 2px;
}

.date-badge {
  background: rgba(255, 255, 255, 0.2);
  border: 2px solid rgba(255, 255, 255, 0.5);
  color: #fff;
  padding: 8px 12px;
  border-radius: 10px;
  font-size: 12px;
  font-weight: 900;
  display: flex;
  gap: 20px;
  white-space: nowrap;
  backdrop-filter: blur(10px);
}

.date-item {
  display: flex;
  flex-direction: column;
  gap: 2px;
  align-items: center;
}

.date-label {
  font-size: 10px;
  opacity: 0.85;
}

.date-value {
  font-size: 13px;
  font-weight: 900;
}

.controls {
  display: flex;
  align-items: center;
  gap: 10px;
  flex: 1;
  justify-content: center;
  min-width: 280px;
}

.ctrl-label {
  font-size: 13px;
  font-weight: 700;
  color: rgba(255,255,255,.9);
}

.multi {
  position: relative;
  width: min(520px, 70vw);
}

.multi-btn {
  width: 100%;
  background: #fff;
  border: 0;
  border-radius: 10px;
  padding: 8px 10px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
  cursor: pointer;
}

.multi-btn .left {
  display: flex;
  flex-direction: column;
  gap: 2px;
  align-items: flex-start;
}

.multi-btn .top {
  font-size: 13px;
  font-weight: 800;
  color: var(--text);
}

.multi-btn .hint {
  font-size: 12px;
  color: var(--muted);
  max-width: 430px;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.multi-btn .badge {
  background: #e2e8f0;
  color: #0f172a;
  font-weight: 800;
  font-size: 12px;
  padding: 4px 8px;
  border-radius: 999px;
}

.multi-menu {
  display: none;
  position: absolute;
  top: calc(100% + 10px);
  left: 0;
  width: 100%;
  min-width: 320px;
  max-height: 520px;
  background: #fff;
  border: 1px solid var(--line);
  border-radius: 12px;
  box-shadow: var(--shadow);
  z-index: 1000;
  overflow: hidden;
}

.multi-menu.show { display: grid; grid-template-rows: auto 1fr auto; }

.multi-top {
  padding: 10px 10px 8px;
  border-bottom: 1px solid var(--line);
  display: grid;
  grid-template-columns: 1fr auto auto;
  gap: 8px;
  align-items: center;
}

.multi-top input {
  width: 100%;
  border: 1px solid var(--line);
  border-radius: 10px;
  padding: 9px 10px;
  font-size: 13px;
  outline: none;
}

.mini-btn {
  border: 1px solid var(--line);
  background: #fff;
  border-radius: 10px;
  padding: 9px 10px;
  font-size: 13px;
  font-weight: 800;
  cursor: pointer;
}

.mini-btn:hover { background: #f8fafc; }

.multi-list {
  padding: 6px 8px;
  overflow: auto;
}

.chk {
  display: flex;
  gap: 10px;
  align-items: center;
  padding: 8px 8px;
  border-radius: 10px;
  cursor: pointer;
  user-select: none;
}

.chk:hover { background: #f1f5f9; }

.chk input {
  width: 16px;
  height: 16px;
  accent-color: var(--btn);
}

.chk span {
  font-size: 13px;
  color: #0f172a;
}

.multi-foot {
  padding: 10px;
  border-top: 1px solid var(--line);
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 10px;
}

.act {
  border: 0;
  border-radius: 10px;
  padding: 10px 12px;
  font-weight: 900;
  font-size: 13px;
  cursor: pointer;
}

.apply { background: #16a34a; color: #fff; }
.clear { background: #ef4444; color: #fff; }

.actions {
  display: flex;
  gap: 8px;
  align-items: center;
}

.btn {
  border: 0;
  border-radius: 10px;
  padding: 10px 12px;
  font-size: 13px;
  font-weight: 900;
  cursor: pointer;
  color: #fff;
}

.btn.export { background: #16a34a; }
.btn.raw { background: #0a6a3a; }
.btn.reset { background: #ef4444; }

main {
  display: grid;
  grid-template-columns: 230px 1fr;
  gap: 12px;
  padding: 12px;
  height: 100%;
  overflow: hidden;
}

aside {
  background: var(--card);
  border-radius: 14px;
  box-shadow: var(--shadow);
  border: 1px solid var(--line);
  overflow: auto;
  padding: 10px;
}

.navbtn {
  display: block;
  text-decoration: none;
  text-align: center;
  font-size: 12.5px;
  font-weight: 900;
  color: #fff;
  padding: 10px 10px;
  border-radius: 12px;
  margin-bottom: 8px;
  background: linear-gradient(135deg, #f59e0b, #d97706);
}

.navbtn.active {
  background: linear-gradient(135deg, #0b1b3a, #123a7a);
}

section.content {
  background: var(--card);
  border-radius: 14px;
  box-shadow: var(--shadow);
  border: 1px solid var(--line);
  overflow: hidden;
  display: grid;
  grid-template-rows: auto 1fr;
}

.topbar {
  padding: 12px 14px;
  border-bottom: 1px solid var(--line);
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 12px;
}

.topbar .h1 {
  font-size: 15px;
  font-weight: 900;
  color: #0b1b3a;
}

.topbar .meta {
  font-size: 12px;
  color: var(--muted);
  font-weight: 800;
  white-space: nowrap;
}

.view {
  overflow: auto;
  padding: 12px;
}

.tbl-title {
  background: linear-gradient(135deg, #4c1d95, #312e81);
  color: #fff;
  padding: 9px 12px;
  font-weight: 900;
  font-size: 13px;
  border-radius: 12px 12px 0 0;
}

.tbl-wrap {
  overflow: auto;
  border: 1px solid var(--line);
  border-top: 0;
  border-radius: 0 0 12px 12px;
}

table.pivot {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  font-size: 12.5px;
  min-width: 820px;
}

table.pivot thead th {
  position: sticky;
  top: 0;
  z-index: 10;
  background: linear-gradient(180deg, #1d4ed8, #1e40af);
  color: #fff;
  padding: 8px 10px;
  border-bottom: 1px solid #1e3a8a;
  text-align: center;
  font-weight: 900;
  white-space: nowrap;
}

table.pivot thead th:first-child {
  left: 0;
  z-index: 12;
  text-align: left;
}

table.pivot td {
  border-bottom: 1px solid var(--line);
  padding: 7px 10px;
  white-space: nowrap;
}

td.rowlabel {
  position: sticky;
  left: 0;
  z-index: 5;
  background: #eef2f7;
  font-weight: 900;
  color: #0b1b3a;
  text-align: left;
}

td.num { text-align: right; }

tr.total td {
  background: linear-gradient(180deg, #1d4ed8, #1e40af) !important;
  color: #fff !important;
  font-weight: 900;
}

.neg { color: #0b7a2d; font-weight: 900; }

@media (max-width: 980px) {
  main { grid-template-columns: 1fr; }
  aside { display: none; }
  table.pivot { min-width: 780px; }
}

@media (max-width: 768px) {
  body { overflow-x: hidden; }
  header { padding: 10px 10px; }

  .hrow {
    flex-direction: column;
    align-items: stretch;
    gap: 10px;
  }

  .brand { min-width: 0; width: 100%; text-align: left; }

  .date-badge {
    width: 100%;
    justify-content: space-around;
    padding: 10px;
  }

  .controls {
    width: 100%;
    min-width: 0;
    flex-direction: column;
    align-items: stretch;
    justify-content: flex-start;
    gap: 10px;
  }

  .ctrl-label { width: 100%; font-size: 12px; margin: 0; }

  .multi, .multi.loc { width: 100% !important; min-width: 0 !important; }

  .multi-btn { width: 100%; padding: 10px 10px; }

  .multi-menu {
    left: 0;
    width: 100% !important;
    min-width: 0 !important;
    max-height: 60vh;
  }

  .multi-top { grid-template-columns: 1fr; gap: 8px; }
  .mini-btn { width: 100%; }

  .actions {
    width: 100%;
    display: grid;
    grid-template-columns: 1fr;
    gap: 8px;
  }

  .btn { width: 100%; text-align: center; padding: 12px; font-size: 13px; }

  main { grid-template-columns: 1fr; padding: 8px; }
  aside { display: none; }

  .topbar { flex-direction: column; align-items: flex-start; gap: 6px; }
  .view { padding: 8px; }

  .tbl-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
  table.pivot { min-width: 760px; }
}

@media (max-width: 420px) {
  .brand .title { font-size: 16px; }
  .brand .sub { font-size: 11px; }
  .multi-btn .top { font-size: 12px; }
  .multi-btn .hint { font-size: 11px; }
  .date-badge { flex-direction: column; gap: 8px; }
}
</style>

<script>
function byId(id) { return document.getElementById(id); }

// PRODUCT DIVISION MULTI
function toggleMenu() {
  const m = byId("menu");
  m.classList.toggle("show");
  if (m.classList.contains("show")) {
    byId("search").focus();
    filterList();
  }
}
function closeMenu() {
  const m = byId("menu");
  m.classList.remove("show");
}
function getSelected() {
  const nodes = document.querySelectorAll('#list input[type="checkbox"]');
  const selected = [];
  nodes.forEach(cb => { if (cb.checked) selected.push(cb.value.trim()); });
  return selected;
}
function setAll(state) {
  const nodes = document.querySelectorAll('#list input[type="checkbox"]');
  nodes.forEach(cb => cb.checked = state);
  updateButtonText();
}
function clearSelection() { setAll(false); }

function updateButtonText() {
  const selected = getSelected();
  const badge = byId("badge");
  const hint = byId("hint");
  badge.textContent = selected.length.toString();

  if (selected.length === 0) {
    hint.textContent = "All Divisions";
  } else if (selected.length <= 3) {
    hint.textContent = selected.join(", ");
  } else {
    hint.textContent = selected.slice(0,3).join(", ") + " + " + (selected.length - 3).toString() + " more";
  }
}
function applyDivisions() {
  const selected = getSelected();
  const sheet = "__SHEET__";
  let divisionsParam = "All";
  if (selected.length > 0) divisionsParam = selected.join(",");
  closeMenu();

  const locationsParam = "__LOCATIONS_PARAM__";
  const newUrl =
    "/?sheet=" + encodeURIComponent(sheet) +
    "&divisions=" + encodeURIComponent(divisionsParam) +
    "&locations=" + encodeURIComponent(locationsParam);

  window.location.href = newUrl;
}
function filterList() {
  const q = byId("search").value.toLowerCase().trim();
  const items = document.querySelectorAll('#list .chk');
  let shown = 0;
  items.forEach(it => {
    const txt = it.innerText.toLowerCase();
    const ok = txt.indexOf(q) !== -1;
    it.style.display = ok ? "flex" : "none";
    if (ok) shown++;
  });
  byId("shown").textContent = shown.toString();
}

// LOCATION MULTI
function toggleLocMenu() {
  const m = byId("loc_menu");
  m.classList.toggle("show");
  if (m.classList.contains("show")) {
    byId("loc_search").focus();
    filterLocList();
  }
}
function closeLocMenu() {
  const m = byId("loc_menu");
  m.classList.remove("show");
}
function getLocSelected() {
  const nodes = document.querySelectorAll('#loc_list input[type="checkbox"]');
  const selected = [];
  nodes.forEach(cb => { if (cb.checked) selected.push(cb.value.trim()); });
  return selected;
}
function setLocAll(state) {
  const nodes = document.querySelectorAll('#loc_list input[type="checkbox"]');
  nodes.forEach(cb => cb.checked = state);
  updateLocButtonText();
}
function clearLocSelection() { setLocAll(false); }

function updateLocButtonText() {
  const selected = getLocSelected();
  const badge = byId("loc_badge");
  const hint = byId("loc_hint");
  badge.textContent = selected.length.toString();

  if (selected.length === 0) {
    hint.textContent = "All Locations";
  } else if (selected.length <= 6) {
    hint.textContent = selected.join(", ");
  } else {
    hint.textContent = selected.slice(0,6).join(", ") + " + " + (selected.length - 6).toString() + " more";
  }
}
function applyLocations() {
  const selected = getLocSelected();
  const sheet = "__SHEET__";
  const divisionsParam = "__DIVISIONS_PARAM__";
  let locationsParam = "All";
  if (selected.length > 0) locationsParam = selected.join(",");
  closeLocMenu();

  const newUrl =
    "/?sheet=" + encodeURIComponent(sheet) +
    "&divisions=" + encodeURIComponent(divisionsParam) +
    "&locations=" + encodeURIComponent(locationsParam);

  window.location.href = newUrl;
}
function filterLocList() {
  const q = byId("loc_search").value.toLowerCase().trim();
  const items = document.querySelectorAll('#loc_list .chk');
  let shown = 0;
  items.forEach(it => {
    const txt = it.innerText.toLowerCase();
    const ok = txt.indexOf(q) !== -1;
    it.style.display = ok ? "flex" : "none";
    if (ok) shown++;
  });
  byId("loc_shown").textContent = shown.toString();
}

// EXPORT / CLEAR
function exportData() {
  const sheet = "__SHEET__";
  const divisionsParam = "__DIVISIONS_PARAM__";
  const locationsParam = "__LOCATIONS_PARAM__";
  window.location.href =
    "/download?sheet=" + encodeURIComponent(sheet) +
    "&divisions=" + encodeURIComponent(divisionsParam) +
    "&locations=" + encodeURIComponent(locationsParam);
}
function exportRawData() {
  const divisionsParam = "__DIVISIONS_PARAM__";
  const locationsParam = "__LOCATIONS_PARAM__";
  window.location.href =
    "/download-raw?divisions=" + encodeURIComponent(divisionsParam) +
    "&locations=" + encodeURIComponent(locationsParam);
}
function clearAll() {
  window.location.href = "/?sheet=Overall_Division_Pivot&divisions=All&locations=All";
}

// Close menus when clicking outside
document.addEventListener("click", function(e) {
  const box = document.querySelector(".multi");
  if (box && !box.contains(e.target)) closeMenu();

  const lbox = document.querySelector(".multi.loc");
  if (lbox && !lbox.contains(e.target)) closeLocMenu();
});

// Update badges on checkbox change
document.addEventListener("change", function(e) {
  if (e.target && e.target.matches('#list input[type="checkbox"]')) updateButtonText();
  if (e.target && e.target.matches('#loc_list input[type="checkbox"]')) updateLocButtonText();
});
</script>
</head>

<body onload="updateButtonText(); updateLocButtonText();">
<div class="app">
  <header>
    <div class="hrow">
      <div class="brand">
        <div class="title">Unnati Stock Kundali</div>
      </div>

      <div class="date-badge">
        <div class="date-item">
          <div class="date-label">📍 Open Stock Date</div>
          <div class="date-value">__OPEN_DATE__</div>
        </div>
        <div class="date-item">
          <div class="date-label">📍 Close Stock Date</div>
          <div class="date-value">__CLOSE_DATE__</div>
        </div>
      </div>

      <div class="controls">
        <div class="ctrl-label">Product Division</div>

        <div class="multi">
          <button class="multi-btn" type="button" onclick="toggleMenu()">
            <div class="left">
              <div class="top">Select Divisions</div>
              <div class="hint" id="hint">All Divisions</div>
            </div>
            <div class="badge" id="badge">0</div>
          </button>

          <div class="multi-menu" id="menu">
            <div class="multi-top">
              <input id="search" type="text" placeholder="Search division..." oninput="filterList()" />
              <button class="mini-btn" type="button" onclick="setAll(true)">Select All</button>
              <button class="mini-btn" type="button" onclick="setAll(false)">Clear</button>
            </div>

            <div class="multi-list" id="list">
              __DIV_CHECKBOXES__
            </div>

            <div class="multi-foot">
              <button class="act apply" type="button" onclick="applyDivisions()">Apply</button>
              <button class="act clear" type="button" onclick="clearSelection()">Clear Selection</button>
            </div>
          </div>
        </div>

        <div style="font-size:12px;color:rgba(255,255,255,.8);font-weight:800;">
          Showing: <span id="shown">__DIV_COUNT__</span> / __DIV_COUNT__
        </div>

        <div class="ctrl-label">Location</div>

        <div class="multi loc">
          <button class="multi-btn" type="button" onclick="toggleLocMenu()">
            <div class="left">
              <div class="top">Select Locations</div>
              <div class="hint" id="loc_hint">All Locations</div>
            </div>
            <div class="badge" id="loc_badge">0</div>
          </button>

          <div class="multi-menu" id="loc_menu">
            <div class="multi-top">
              <input id="loc_search" type="text" placeholder="Search location..." oninput="filterLocList()" />
              <button class="mini-btn" type="button" onclick="setLocAll(true)">Select All</button>
              <button class="mini-btn" type="button" onclick="setLocAll(false)">Clear</button>
            </div>

            <div class="multi-list" id="loc_list">
              __LOC_CHECKBOXES__
            </div>

            <div class="multi-foot">
              <button class="act apply" type="button" onclick="applyLocations()">Apply</button>
              <button class="act clear" type="button" onclick="clearLocSelection()">Clear Selection</button>
            </div>
          </div>
        </div>

        <div style="font-size:12px;color:rgba(255,255,255,.8);font-weight:800;">
          Showing: <span id="loc_shown">__LOC_COUNT__</span> / __LOC_COUNT__
        </div>
      </div>

      <div class="actions">
        <button class="btn export" onclick="exportData()">Export CSV</button>
        <button class="btn raw" onclick="exportRawData()">Export Raw CSV</button>
        <button class="btn reset" onclick="clearAll()">Clear All</button>
      </div>
    </div>
  </header>

  <main>
    <aside>
      __LEFT_BUTTONS__
    </aside>

    <section class="content">
      <div class="topbar">
        <div class="h1">__TITLE_LINE__</div>
        <div class="meta">Rows depend on selected divisions / locations</div>
      </div>
      <div class="view">
        __TABLE_HTML__
      </div>
    </section>
  </main>
</div>
</body>
</html>
"""

        # Fill tokens safely
        title_line = f"{safe_title(sheet)} - {esc_attr(divisions_display)} | {esc_attr(locations_display)}"
        html = (html
                .replace("__SHEET__", esc_attr(sheet))
                .replace("__DIVISIONS_PARAM__", esc_attr(divisions_param))
                .replace("__LOCATIONS_PARAM__", esc_attr(locations_param))
                .replace("__DIV_CHECKBOXES__", checkbox_html)
                .replace("__LOC_CHECKBOXES__", loc_checkbox_html)
                .replace("__DIV_COUNT__", str(len(all_divisions)))
                .replace("__LOC_COUNT__", str(len(all_locations)))
                .replace("__LEFT_BUTTONS__", buttons_html)
                .replace("__TITLE_LINE__", title_line)
                .replace("__TABLE_HTML__", table_html)
                .replace("__OPEN_DATE__", dates["open_date"])
                .replace("__CLOSE_DATE__", dates["close_date"])
               )

        return HTMLResponse(content=html)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return HTMLResponse(f"<h1>Error: {e}</h1>")


# =========================================================
# MAIN (Render uses startCommand: uvicorn app:app --port $PORT)
# =========================================================
if __name__ == "__main__":
    port = int(os.getenv("PORT", "8002"))
    # Load dates on startup
    load_dates()
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
