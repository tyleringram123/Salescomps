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

# Excel file lives in the SAME folder as this Map.py by default.
# You can override on DigitalOcean by setting EXCEL_PATH env var.
EXCEL_PATH = os.getenv("EXCEL_PATH", os.path.join(BASE_DIR, "All Sales Comps.xlsm"))
SHEET_NAME = os.getenv("SHEET_NAME", "Source Data")

# =====================================================
# PRACTICE ONLY – HARD-CODED GOOGLE MAPS KEY (DO can use env var)
# =====================================================
GMAPS_API_KEY = os.getenv("GMAPS_API_KEY", "AIzaSyBIcVzJwkW20rIbkqdi9Yfhpiog9fp8y4s")
gmaps = googlemaps.Client(key=GMAPS_API_KEY) if GMAPS_API_KEY else None

app = Flask(__name__)

# =====================================================
# Excel load cache (per Gunicorn worker)
# Prevents repeated Excel reads that cause OOM/SIGKILL on small RAM.
# =====================================================
_DF_CACHE = None
_DF_CACHE_MTIME = None


def load_sales_comps_df():
    global _DF_CACHE, _DF_CACHE_MTIME

    if not os.path.isfile(EXCEL_PATH):
        raise FileNotFoundError(f"Excel file not found at: {EXCEL_PATH}")

    mtime = os.path.getmtime(EXCEL_PATH)

    # Reuse cached DataFrame if the file hasn't changed
    if _DF_CACHE is not None and _DF_CACHE_MTIME == mtime:
        return _DF_CACHE

    # If you know the minimum columns needed, uncomment and set usecols for BIG memory savings.
    # Example:
    # usecols = ["Property Name", "Address", "City", "State", "Zip", "Latitude", "Longitude", "Type"]
    # df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl", usecols=usecols)

    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")

    _DF_CACHE = df
    _DF_CACHE_MTIME = mtime
    return df


def try_geocode(address: str):
    """Return (lat, lon, formatted_address) or (None, None, None)."""
    if not address or not gmaps:
        return None, None, None
    try:
        result = gmaps.geocode(address)
        if not result:
            return None, None, None
        loc = result[0]["geometry"]["location"]
        formatted = result[0].get("formatted_address", address)
        return loc["lat"], loc["lng"], formatted
    except Exception:
        return None, None, None


def first_existing_col(df, candidates):
    """Return first column name that exists from a list of candidate names."""
    cols = set(df.columns.astype(str))
    for c in candidates:
        if c in cols:
            return c
    return None


@app.get("/health")
def health():
    # IMPORTANT: do NOT load Excel here. Keep this endpoint light.
    return "ok", 200


@app.route("/", methods=["GET", "POST"])
def index():
    # ---- Read search address from form ----
    search_address = ""
    if request.method == "POST":
        search_address = (request.form.get("address") or "").strip()

    # ---- Load df (cached) ----
    try:
        df = load_sales_comps_df()
    except Exception as e:
        # Return a useful error instead of crashing worker / endless 500 loop
        return f"Error loading Excel: {e}", 500

    # ---- Try to find lat/lon columns (be flexible with naming) ----
    lat_col = first_existing_col(df, ["Latitude", "Lat", "latitude", "LAT", "Y"])
    lon_col = first_existing_col(df, ["Longitude", "Lng", "Lon", "longitude", "LON", "X"])

    if not lat_col or not lon_col:
        return (
            "Could not find Latitude/Longitude columns in your sheet.\n"
            "Looked for: Latitude/Lat/Y and Longitude/Lng/Lon/X.\n"
            f"Columns found: {list(df.columns)}",
            500,
        )

    # Coerce to numeric and drop bad rows
    df = df.copy()
    df[lat_col] = pd.to_numeric(df[lat_col], errors="coerce")
    df[lon_col] = pd.to_numeric(df[lon_col], errors="coerce")
    df = df.dropna(subset=[lat_col, lon_col])

    # ---- Determine map center ----
    center_lat = float(df[lat_col].median()) if len(df) else 39.5
    center_lon = float(df[lon_col].median()) if len(df) else -98.35

    subject_lat, subject_lon, subject_label = None, None, None
    if search_address:
        subject_lat, subject_lon, subject_label = try_geocode(search_address)
        if subject_lat is not None and subject_lon is not None:
            center_lat, center_lon = subject_lat, subject_lon

    # ---- Build map ----
    m = folium.Map(location=[center_lat, center_lon], zoom_start=11, tiles="OpenStreetMap")

    # Add the subject pin first (if any)
    if subject_lat is not None and subject_lon is not None:
        folium.Marker(
            location=[subject_lat, subject_lon],
            tooltip="Subject Pin",
            popup=folium.Popup(subject_label or search_address, max_width=400),
            icon=folium.Icon(color="red", icon="info-sign"),
        ).add_to(m)

    # ---- Clustering controls ----
    # You can adjust these two values to change clustering behavior:
    CLUSTER_MAX_ZOOM = int(os.getenv("CLUSTER_MAX_ZOOM", "14"))  # higher = clusters break apart later
    CLUSTER_DISABLE_AT_ZOOM = int(os.getenv("CLUSTER_DISABLE_AT_ZOOM", "18"))  # set lower to stop clustering sooner

    marker_cluster = MarkerCluster(
        name="Sales Comps",
        maxClusterRadius=50,
        disableClusteringAtZoom=CLUSTER_DISABLE_AT_ZOOM,
    ).add_to(m)

    # Optional subgrouping by a "Type" column if it exists
    type_col = first_existing_col(df, ["Type", "Comp Type", "Category", "Deal Type"])

    # Create subgroups if type_col exists
    subgroup_layers = {}
    if type_col:
        for tval in sorted(df[type_col].dropna().astype(str).unique()):
            subgroup = FeatureGroupSubGroup(marker_cluster, name=tval)
            subgroup.add_to(m)
            subgroup_layers[tval] = subgroup

    # Pick a label column (for popup)
    label_col = first_existing_col(
        df,
        ["Property Name", "Asset Name", "Name", "Community", "Property", "Address"],
    )

    # Add points
    for _, row in df.iterrows():
        lat = float(row[lat_col])
        lon = float(row[lon_col])

        label = ""
        if label_col and pd.notna(row.get(label_col)):
            label = str(row.get(label_col))
        else:
            label = "Comp"

        popup_html = f"<b>{label}</b><br>Lat: {lat}<br>Lon: {lon}"

        target_layer = marker_cluster
        if type_col:
            tval = row.get(type_col)
            if pd.notna(tval):
                tval = str(tval)
                target_layer = subgroup_layers.get(tval, marker_cluster)

        folium.CircleMarker(
            location=[lat, lon],
            radius=4,
            popup=folium.Popup(popup_html, max_width=350),
        ).add_to(target_layer)

    folium.LayerControl(collapsed=False).add_to(m)

    map_html = m.get_root().render()

    # ---- Sticky search bar template (does not disappear when zooming) ----
    template = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8"/>
        <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
        <title>Sales Comps Map</title>
        <style>
            body { margin:0; padding:0; font-family: Arial, sans-serif; }
            .topbar {
                position: sticky;
                top: 0;
                z-index: 9999;
                background: rgba(255,255,255,0.95);
                padding: 10px 12px;
                border-bottom: 1px solid #ddd;
            }
            .topbar form { display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
            .topbar input[type="text"] {
                width: min(720px, 80vw);
                padding: 10px;
                border: 1px solid #bbb;
                border-radius: 6px;
                font-size: 14px;
            }
            .topbar button {
                padding: 10px 14px;
                border: 1px solid #333;
                background: #fff;
                border-radius: 6px;
                cursor: pointer;
                font-size: 14px;
            }
            .hint { font-size: 12px; color: #555; margin-top: 6px; }
            /* Make map fill remaining space */
            .mapwrap { height: calc(100vh - 72px); }
            iframe, .folium-map { width: 100%; height: 100%; border: 0; }
        </style>
    </head>
    <body>
        <div class="topbar">
            <form method="POST">
                <input type="text" name="address" placeholder="Enter address to drop a subject pin..." value="{{addr}}">
                <button type="submit">Add Pin</button>
            </form>
            <div class="hint">
                Tip: On DigitalOcean, set Health Check Path to <b>/health</b> so it doesn't load Excel for checks.
                Also set Gunicorn workers to 1 on 512MB.
            </div>
        </div>

        <div class="mapwrap">
            {{ map_html | safe }}
        </div>
    </body>
    </html>
    """

    return render_template_string(template, map_html=map_html, addr=search_address)


if __name__ == "__main__":
    # Local dev only; DO uses gunicorn
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=True)

