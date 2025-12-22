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

# Excel file lives in the SAME folder as this Map.py by default.
# You can override on DigitalOcean by setting EXCEL_PATH env var.
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join(BASE_DIR, "All Sales Comps.xlsm"))
SHEET_NAME = os.getenv("SHEET_NAME", "Source Data")

# =====================================================
# PRACTICE ONLY – HARD-CODED GOOGLE MAPS KEY (but DO can use env var)
# =====================================================
GMAPS_API_KEY = os.getenv("GMAPS_API_KEY", "AIzaSyBIcVzJwkW20rIbkqdi9Yfhpiog9fp8y4s")
gmaps = googlemaps.Client(key=GMAPS_API_KEY)

# =====================================================
# DATA CONFIG
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
CLUSTER_MAX_RADIUS = 60        # pixels; bigger = more aggressive clustering
CLUSTER_DISABLE_AT_ZOOM = 11   # int; stops clustering when zoomed in this far
CLUSTER_SPIDERFY_ON_MAX = True # expands overlapping markers on click
CLUSTER_SHOW_COVERAGE = False  # shows polygon coverage on hover
CLUSTER_ZOOM_TO_BOUNDS = True  # zoom to cluster bounds when clicked

# =====================================================
# FLASK APP
# =====================================================
app = Flask(__name__)

# =====================================================
# HELPERS
# =====================================================
def geocode_address(address: str):
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
                f"<tr>"
                f"<td style='padding:3px 8px; white-space:nowrap;'><b>{label}</b></td>"
                f"<td style='padding:3px 8px;'>{v}</td>"
                f"</tr>"
            )

    return f"""
    <div style="min-width:260px;max-width:420px;">
        <div style="font-size:16px;font-weight:700">{title}</div>
        <div style="font-size:12px;color:#555;margin-bottom:6px">{address}</div>
        <table style="font-size:13px;border-collapse:collapse;">{''.join(rows)}</table>
    </div>
    """

def load_sales_comps_df():
    # Helpful error if the file isn’t present on DO
    if not os.path.exists(EXCEL_PATH):
        raise FileNotFoundError(
            f"Excel file not found at: {EXCEL_PATH}\n"
            f"Put 'All Sales Comps.xlsm' in the same folder as Map.py, "
            f"or set EXCEL_PATH as an environment variable."
        )

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
    df["lat"] = pd.to_numeric(df.iloc[:, LAT_COL_INDEX], errors="coerce")
    df["lon"] = pd.to_numeric(df.iloc[:, LON_COL_INDEX], errors="coerce")
    df["year"] = df.iloc[:, SALE_DATE_INDEX].apply(get_sale_year)
    df = df.dropna(subset=["lat", "lon"])
    return df

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
        # Render a simple readable error in the browser (super helpful on DO)
        return (
            f"<h2>App Error</h2><pre>{str(e)}</pre>"
            f"<p>EXCEL_PATH currently: <code>{EXCEL_PATH}</code></p>",
            500,
        )

    # Center map on subject if provided
    if subject_location:
        m = folium.Map(location=subject_location, zoom_start=12)
    else:
        m = folium.Map(location=[df["lat"].mean(), df["lon"].mean()], zoom_start=6)

    # ---- CLUSTER (tunable via config above) ----
    cluster = MarkerCluster(
        name="Sales Comps",
        maxClusterRadius=CLUSTER_MAX_RADIUS,
        disableClusteringAtZoom=CLUSTER_DISABLE_AT_ZOOM,
        spiderfyOnMaxZoom=CLUSTER_SPIDERFY_ON_MAX,
        showCoverageOnHover=CLUSTER_SHOW_COVERAGE,
        zoomToBoundsOnClick=CLUSTER_ZOOM_TO_BOUNDS,
    ).add_to(m)

    # Year subgroups inside SAME cluster (area-based)
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

    # SUBJECT PIN
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

        /* Keep the search bar always visible */
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

        /* Map area below the fixed bar */
        .mapwrap{
          position: fixed;
          top: 56px;  /* match bar height */
          left: 0; right: 0; bottom: 0;
        }

        /* Force folium iframe to fill the mapwrap */
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

# NOTE:
# In production on DigitalOcean, you run with gunicorn:
#   gunicorn -c gunicorn_config.py Map:app
# So this __main__ block is only for local testing.
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "5000")), debug=True)
