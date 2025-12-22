# Map.py
import os
from flask import Flask, render_template_string, request
import pandas as pd
import folium
from folium.plugins import MarkerCluster, FeatureGroupSubGroup
import googlemaps

# =====================================================
# DIGITALOCEAN / LINUX-SAFE PATHS
# =====================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Excel lives next to Map.py by default
# Override with EXCEL_PATH env var on DigitalOcean if needed
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join(BASE_DIR, "All Sales Comps.xlsm"))
SHEET_NAME = os.getenv("SHEET_NAME", "Source Data")

# =====================================================
# GOOGLE MAPS (HARD-CODE KEY, BUT DO-SAFE)
# =====================================================
GMAPS_API_KEY = os.getenv("GMAPS_API_KEY", "AIzaSyBIcVzJwkW20rIbkqdi9Yfhpiog9fp8y4s").strip()
_gmaps_client = None

def get_gmaps_client():
    """Lazily create Google Maps client so bad keys never crash boot."""
    global _gmaps_client
    if _gmaps_client is not None:
        return _gmaps_client
    if not GMAPS_API_KEY:
        return None
    try:
        _gmaps_client = googlemaps.Client(key=GMAPS_API_KEY)
        return _gmaps_client
    except Exception as e:
        print(f"[WARN] Google Maps client disabled: {e}")
        _gmaps_client = None
        return None

# =====================================================
# DATA CONFIG (KEEP YOUR INDEXES)
# =====================================================
LAT_COL_INDEX = 5      # F
LON_COL_INDEX = 6      # G
SALE_DATE_INDEX = 11   # L

POPUP_FIELDS = [
    ("Property", 1),
    ("Full Address", 4),
    ("Units", 2),
    ("Vintage", 3),
    ("Sale Date", 11),
    ("Sale Price", 12),
    ("Price per Unit", 13),
    ("Cap Rate", 14),
    ("Buyer", 15),
    ("Seller", 16),
    ("Source", 17),
]

YEAR_COLOR_MAP = {
    2015: "darkblue",
    2016: "cadetblue",
    2017: "blue",
    2018: "lightblue",
    2019: "green",
    2020: "darkgreen",
    2021: "lightgreen",
    2022: "orange",
    2023: "red",
    2024: "purple",
    2025: "darkred",
}
DEFAULT_YEAR_COLOR = "gray"

# =====================================================
# CLUSTERING CONFIG (EDIT THESE)
# =====================================================
CLUSTER_MAX_RADIUS = int(os.getenv("CLUSTER_MAX_RADIUS", "60"))
CLUSTER_DISABLE_AT_ZOOM = int(os.getenv("CLUSTER_DISABLE_AT_ZOOM", "11"))
CLUSTER_SPIDERFY_ON_MAX = os.getenv("CLUSTER_SPIDERFY_ON_MAX", "true").lower() == "true"
CLUSTER_SHOW_COVERAGE = os.getenv("CLUSTER_SHOW_COVERAGE", "false").lower() == "true"
CLUSTER_ZOOM_TO_BOUNDS = os.getenv("CLUSTER_ZOOM_TO_BOUNDS", "true").lower() == "true"

# =====================================================
# FLASK APP
# =====================================================
app = Flask(__name__)

# =====================================================
# EXCEL HEADER ROW DETECTION
# (handles sheets where the real header isn't row 1)
# Keeps your fixed indexes working by reading with the correct header row.
# =====================================================
def detect_header_row(excel_path: str, sheet_name: str, scan_rows: int = 80) -> int:
    """
    Scan first scan_rows rows (header=None) to find the row containing lat/lon headers.
    Returns 0-based row index to use for header=...
    """
    try:
        preview = pd.read_excel(
            excel_path,
            sheet_name=sheet_name,
            engine="openpyxl",
            header=None,
            nrows=scan_rows,
        )
        for r in range(len(preview)):
            row = preview.iloc[r].astype(str).str.strip().str.lower().tolist()
            has_lat = any(("lat" in c) for c in row)
            has_lon = any(("lon" in c) or ("lng" in c) for c in row)
            if has_lat and has_lon:
                return r
    except Exception as e:
        print(f"[WARN] Header detection failed, defaulting to 0: {e}")
    return 0

# =====================================================
# DATAFRAME CACHE (per Gunicorn worker)
# Prevents re-reading big xlsm on every request (OOM/SIGKILL on small RAM)
# =====================================================
_DF_CACHE = None
_DF_CACHE_MTIME = None
_DF_CACHE_HEADER = None

def load_sales_comps_df():
    global _DF_CACHE, _DF_CACHE_MTIME, _DF_CACHE_HEADER

    if not os.path.isfile(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found at: {EXCEL_PATH}")

    mtime = os.path.getmtime(EXCEL_PATH)
    if _DF_CACHE is not None and _DF_CACHE_MTIME == mtime:
        return _DF_CACHE

    header_row = detect_header_row(EXCEL_PATH, SHEET_NAME)
    _DF_CACHE_HEADER = header_row

    df = pd.read_excel(
        EXCEL_PATH,
        sheet_name=SHEET_NAME,
        engine="openpyxl",
        header=header_row,
    )

    _DF_CACHE = df
    _DF_CACHE_MTIME = mtime
    return df

# =====================================================
# HELPERS (kept aligned with your original index-based logic)
# =====================================================
def geocode_address(address):
    gmaps = get_gmaps_client()
    if not gmaps or not address:
        return None
    try:
        result = gmaps.geocode(address)
        if result:
            loc = result[0]["geometry"]["location"]
            return loc["lat"], loc["lng"]
    except Exception as e:
        print("Geocode error:", e)
    return None

def safe_iloc(row, idx):
    try:
        return row.iloc[idx]
    except Exception:
        return None

def format_value(label, raw):
    if raw is None or pd.isna(raw):
        return ""

    if label == "Sale Date":
        dt = pd.to_datetime(raw, errors="coerce")
        return dt.strftime("%m/%Y") if pd.notna(dt) else ""

    if label in ("Sale Price", "Price per Unit"):
        num = pd.to_numeric(raw, errors="coerce")
        return f"${num:,.0f}" if pd.notna(num) else ""

    if label == "Cap Rate":
        num = pd.to_numeric(str(raw).replace("%", ""), errors="coerce")
        if pd.isna(num):
            return ""
        if num <= 1:
            num *= 100
        return f"{num:.2f}%"

    return str(raw)

def get_sale_year(raw):
    dt = pd.to_datetime(raw, errors="coerce")
    return int(dt.year) if pd.notna(dt) else None

def build_popup(row):
    title = format_value("Property", safe_iloc(row, 1))
    address = format_value("Full Address", safe_iloc(row, 4))

    rows = []
    for label, idx in POPUP_FIELDS:
        if idx in (1, 4):
            continue
        v = format_value(label, safe_iloc(row, idx))
        if v:
            rows.append(
                f"<tr><td style='padding:3px 8px; white-space:nowrap;'><b>{label}</b></td>"
                f"<td style='padding:3px 8px;'>{v}</td></tr>"
            )

    return f"""
    <div style="min-width:260px;max-width:420px;">
        <div style="font-size:16px;font-weight:700">{title}</div>
        <div style="font-size:12px;color:#555;margin-bottom:6px">{address}</div>
        <table style="font-size:13px;border-collapse:collapse;">{''.join(rows)}</table>
    </div>
    """

# =====================================================
# HEALTH CHECK (set DO health check path to /health)
# =====================================================
@app.get("/health")
def health():
    return "ok", 200

# =====================================================
# ROUTE
# =====================================================
@app.route("/", methods=["GET", "POST"])
def index():
    subject_address = request.form.get("address", "").strip()
    subject_location = geocode_address(subject_address) if subject_address else None

    try:
        df = load_sales_comps_df()
    except Exception as e:
        return f"Error loading Excel: {e}", 500

    # Keep your fixed index extraction
    df = df.copy()
    df["lat"] = pd.to_numeric(df.iloc[:, LAT_COL_INDEX], errors="coerce")
    df["lon"] = pd.to_numeric(df.iloc[:, LON_COL_INDEX], errors="coerce")
    df["year"] = df.iloc[:, SALE_DATE_INDEX].apply(get_sale_year)
    df = df.dropna(subset=["lat", "lon"])

    # Center map on subject if provided
    if subject_location:
        m = folium.Map(location=subject_location, zoom_start=12)
    else:
        m = folium.Map(location=[df["lat"].mean(), df["lon"].mean()], zoom_start=6)

    # Cluster
    cluster = MarkerCluster(
        name="Sales Comps",
        maxClusterRadius=CLUSTER_MAX_RADIUS,
        disableClusteringAtZoom=CLUSTER_DISABLE_AT_ZOOM,
        spiderfyOnMaxZoom=CLUSTER_SPIDERFY_ON_MAX,
        showCoverageOnHover=CLUSTER_SHOW_COVERAGE,
        zoomToBoundsOnClick=CLUSTER_ZOOM_TO_BOUNDS,
    ).add_to(m)

    # Year subgroups inside the SAME cluster
    year_layers = {}
    for y in YEAR_COLOR_MAP:
        fg = FeatureGroupSubGroup(cluster, name=str(y), show=True)
        fg.add_to(m)
        year_layers[y] = fg

    other_fg = FeatureGroupSubGroup(cluster, name="Other / Unknown", show=False)
    other_fg.add_to(m)

    for _, row in df.iterrows():
        year = row["year"]
        color = YEAR_COLOR_MAP.get(year, DEFAULT_YEAR_COLOR)
        target = year_layers.get(year, other_fg)

        folium.Marker(
            location=[row["lat"], row["lon"]],
            popup=folium.Popup(build_popup(row), max_width=450),
            tooltip=str(safe_iloc(row, 1)),
            icon=folium.Icon(color=color, icon="home", prefix="fa"),
        ).add_to(target)

    # Subject pin
    if subject_location:
        folium.Marker(
            location=subject_location,
            tooltip="Subject Property",
            popup=subject_address,
            icon=folium.Icon(color="black", icon="star", prefix="fa"),
        ).add_to(m)

    folium.LayerControl(collapsed=False).add_to(m)

    return render_template_string(
        """
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1" />
      <style>
        body { margin:0; font-family: Arial, sans-serif; overflow:hidden; }

        .topbar{
          position: fixed;
          top: 0; left: 0; right: 0;
          z-index: 999999;
          background:#f4f4f4;
          border-bottom:1px solid #ddd;
          padding:10px;
          display:flex;
          align-items:center;
          gap:10px;
        }
        .topbar input{
          width: 520px;
          max-width: calc(100vw - 220px);
          padding:8px;
        }
        .topbar button{
          padding:8px 14px;
          cursor:pointer;
        }

        .mapwrap{
          position: fixed;
          top: 56px;
          left: 0; right: 0; bottom: 0;
        }
        .mapwrap iframe{
          width: 100% !important;
          height: 100% !important;
          border: 0 !important;
        }
      </style>
    </head>
    <body>
      <form method="post" class="topbar">
        <input type="text" name="address" placeholder="Enter subject address"
               value="{{ subject_address }}">
        <button type="submit">Map</button>
      </form>

      <div class="mapwrap">
        {{ map_html|safe }}
      </div>
    </body>
    </html>
    """,
        map_html=m._repr_html_(),
        subject_address=subject_address,
    )

if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)
