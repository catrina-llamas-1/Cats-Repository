"""
Interactive folium dashboard: address pins filterable by submarket,
with a side table showing details for every pin currently on the map.

Reads addresses + submarket categories (+ any extra info columns) from an
Excel file, geocodes each address, and produces a single self-contained
HTML file with:
  - a Leaflet/folium map, one colored pin per address (color = submarket)
  - a checkbox filter panel (toggle any combination of submarkets)
  - a side table listing every currently-visible pin; clicking a row pans
    the map to that pin and opens its popup, clicking a pin highlights its row

Usage:
    pip install folium geopy pandas openpyxl
    python submarket_map_dashboard.py

Edit the CONFIG section below to point at your Excel file and columns.
"""

import json
import os
import pandas as pd
import folium
from branca.element import MacroElement
from jinja2 import Template
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# ── CONFIG ────────────────────────────────────────────────────────────────────
EXCEL_FILE    = "properties.xlsx"   # input spreadsheet (.xlsx/.xls/.csv)
ADDRESS_COL   = "Address"           # column holding the address to geocode
SUBMARKET_COL = "Submarket"         # column holding the filter category
INFO_COLS     = None                # list of extra columns to show in the table/popup;
                                     # None = use every other column in the file

OUTPUT_PATH  = "submarket_map.html"
MAP_TILE     = "cartodbpositron"
ZOOM_START   = 12
SIDEBAR_WIDTH_PX = 380

GEOCODER_USER_AGENT = "cats-repository-address-mapper"
CACHE_PATH           = "geocode_cache.csv"   # remembers past lookups so reruns don't re-geocode everything

# Categorical palette (validated for colorblind-safety, fixed hue order — do not reorder)
CATEGORY_COLORS = [
    "#2a78d6",  # blue
    "#1baf7a",  # aqua
    "#eda100",  # yellow
    "#008300",  # green
    "#4a3aa7",  # violet
    "#e34948",  # red
    "#e87ba4",  # magenta
    "#eb6834",  # orange
]
FALLBACK_COLOR = "#898781"  # used past the 8th distinct submarket
# ─────────────────────────────────────────────────────────────────────────────


def load_rows() -> pd.DataFrame:
    ext = os.path.splitext(EXCEL_FILE)[-1].lower()
    df = pd.read_csv(EXCEL_FILE, dtype=str) if ext == ".csv" else pd.read_excel(EXCEL_FILE, dtype=str)

    for col in (ADDRESS_COL, SUBMARKET_COL):
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in {EXCEL_FILE}. Available columns: {list(df.columns)}")

    df = df.dropna(subset=[ADDRESS_COL])
    df[SUBMARKET_COL] = df[SUBMARKET_COL].fillna("Uncategorized").str.strip().replace("", "Uncategorized")
    return df


def info_columns(df: pd.DataFrame) -> list[str]:
    if INFO_COLS is not None:
        return [c for c in INFO_COLS if c in df.columns]
    return [c for c in df.columns if c not in (ADDRESS_COL, SUBMARKET_COL)]


def load_cache() -> dict:
    if os.path.exists(CACHE_PATH):
        return pd.read_csv(CACHE_PATH, dtype=str).set_index("address").to_dict("index")
    return {}


def save_cache(cache: dict) -> None:
    pd.DataFrame.from_dict(cache, orient="index").rename_axis("address").reset_index().to_csv(CACHE_PATH, index=False)


def geocode_column(addresses: list[str]) -> dict:
    """Return {address: (lat, lon)} for every address that could be geocoded."""
    geolocator = Nominatim(user_agent=GEOCODER_USER_AGENT)
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

    cache = load_cache()
    resolved = {}
    failed = []

    for address in dict.fromkeys(addresses):  # de-duped, order-preserving
        cached = cache.get(address)
        if cached is not None:
            lat, lon = cached.get("lat"), cached.get("lon")
            if pd.isna(lat) or lat in ("", None):
                failed.append(address)
                continue
            resolved[address] = (float(lat), float(lon))
            continue

        location = geocode(address)
        if location:
            cache[address] = {"lat": location.latitude, "lon": location.longitude}
            resolved[address] = (location.latitude, location.longitude)
        else:
            cache[address] = {"lat": "", "lon": ""}
            failed.append(address)

    save_cache(cache)

    if failed:
        print(f"Could not geocode {len(failed)} address(es):")
        for a in failed:
            print(f"  - {a}")

    return resolved


def build_records(df: pd.DataFrame, coords: dict, info_cols: list[str]) -> list[dict]:
    records = []
    for _, row in df.iterrows():
        address = row[ADDRESS_COL]
        if address not in coords:
            continue
        lat, lon = coords[address]
        info = {c: row[c] for c in info_cols if c in row and not pd.isna(row[c])}
        records.append({
            "address": address,
            "submarket": row[SUBMARKET_COL],
            "lat": lat,
            "lon": lon,
            "info": info,
        })
    return records


def assign_colors(records: list[dict]) -> dict:
    categories = sorted({r["submarket"] for r in records})
    colors = {}
    for i, cat in enumerate(categories):
        colors[cat] = CATEGORY_COLORS[i] if i < len(CATEGORY_COLORS) else FALLBACK_COLOR
    if len(categories) > len(CATEGORY_COLORS):
        print(f"Note: {len(categories)} submarkets found but only {len(CATEGORY_COLORS)} distinct "
              f"colors are defined — some categories will share a color. Rely on the table/legend labels too.")
    return colors


class SubmarketDashboard(MacroElement):
    """Injects the sidebar (filters + table) and wires it up to the folium map."""

    def __init__(self, records: list[dict], info_cols: list[str], colors: dict, sidebar_width: int):
        super().__init__()
        self._name = "SubmarketDashboard"
        for i, r in enumerate(records):
            r["i"] = i
        self.records_json = json.dumps(records)
        self.info_cols_json = json.dumps(info_cols)
        self.colors_json = json.dumps(colors)
        self.sidebar_width = sidebar_width
        self._template = Template(u"""
{% macro header(this, kwargs) %}
<style>
  :root {
    --surface-1: #fcfcfb;
    --page: #f9f9f7;
    --ink-primary: #0b0b0b;
    --ink-secondary: #52514e;
    --ink-muted: #898781;
    --gridline: #e1e0d9;
    --border: rgba(11,11,11,0.10);
    --row-hover: #f0efec;
    --row-active: #e5eefb;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --surface-1: #1a1a19;
      --page: #0d0d0d;
      --ink-primary: #ffffff;
      --ink-secondary: #c3c2b7;
      --ink-muted: #898781;
      --gridline: #2c2c2a;
      --border: rgba(255,255,255,0.10);
      --row-hover: #24241f;
      --row-active: #16324f;
    }
  }
  html, body { height: 100%; margin: 0; padding: 0; }
  body {
    display: flex;
    flex-direction: row;
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
    background: var(--page);
    color: var(--ink-primary);
  }
  #{{this._parent.get_name()}} {
    flex: 1 1 auto !important;
    width: auto !important;
    height: 100% !important;
  }
  #sm-sidebar {
    flex: 0 0 {{this.sidebar_width}}px;
    height: 100vh;
    display: flex;
    flex-direction: column;
    background: var(--surface-1);
    border-left: 1px solid var(--border);
    box-sizing: border-box;
  }
  #sm-sidebar h2 {
    font-size: 0.95rem;
    margin: 0;
    padding: 14px 16px 4px;
    color: var(--ink-primary);
  }
  #sm-count {
    font-size: 0.8rem;
    color: var(--ink-secondary);
    padding: 0 16px 10px;
    border-bottom: 1px solid var(--gridline);
  }
  #sm-filters {
    padding: 10px 16px;
    border-bottom: 1px solid var(--gridline);
    max-height: 34vh;
    overflow-y: auto;
  }
  #sm-filters .sm-filter-actions {
    display: flex;
    gap: 10px;
    margin-bottom: 8px;
  }
  #sm-filters .sm-filter-actions a {
    font-size: 0.78rem;
    color: var(--ink-secondary);
    cursor: pointer;
    text-decoration: underline;
  }
  .sm-filter-row {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 3px 0;
    font-size: 0.85rem;
    color: var(--ink-primary);
  }
  .sm-swatch {
    width: 10px;
    height: 10px;
    border-radius: 50%;
    flex: 0 0 auto;
  }
  .sm-filter-row label {
    flex: 1 1 auto;
    cursor: pointer;
  }
  .sm-filter-row .sm-cat-count {
    color: var(--ink-muted);
    font-variant-numeric: tabular-nums;
  }
  #sm-table-wrap {
    flex: 1 1 auto;
    overflow-y: auto;
  }
  table#sm-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.8rem;
  }
  table#sm-table th {
    position: sticky;
    top: 0;
    background: var(--surface-1);
    text-align: left;
    color: var(--ink-muted);
    font-weight: 600;
    padding: 6px 8px;
    border-bottom: 1px solid var(--gridline);
    white-space: nowrap;
  }
  table#sm-table td {
    padding: 6px 8px;
    border-bottom: 1px solid var(--gridline);
    vertical-align: top;
    color: var(--ink-primary);
  }
  table#sm-table tr.sm-row {
    cursor: pointer;
  }
  table#sm-table tr.sm-row:hover {
    background: var(--row-hover);
  }
  table#sm-table tr.sm-row.sm-active {
    background: var(--row-active);
  }
  .sm-empty {
    padding: 16px;
    color: var(--ink-muted);
    font-size: 0.85rem;
  }
</style>
{% endmacro %}

{% macro html(this, kwargs) %}
<div id="sm-sidebar">
  <h2>Pins</h2>
  <div id="sm-count"></div>
  <div id="sm-filters"></div>
  <div id="sm-table-wrap">
    <table id="sm-table">
      <thead><tr id="sm-thead-row"></tr></thead>
      <tbody id="sm-tbody"></tbody>
    </table>
  </div>
</div>
{% endmacro %}

{% macro script(this, kwargs) %}
(function() {
  var map = {{ this._parent.get_name() }};
  var records = {{ this.records_json }};
  var infoCols = {{ this.info_cols_json }};
  var colors = {{ this.colors_json }};

  var categories = Object.keys(colors).sort();
  var activeCats = {};
  categories.forEach(function(c) { activeCats[c] = true; });

  // ── build markers ──────────────────────────────────────────────────────
  records.forEach(function(r) {
    var color = colors[r.submarket] || "#898781";
    var marker = L.circleMarker([r.lat, r.lon], {
      radius: 8,
      weight: 2,
      color: "#ffffff",
      fillColor: color,
      fillOpacity: 0.9
    });
    marker.bindPopup(popupContent(r));
    marker.on("click", function() { highlightRow(r.i); });
    r._marker = marker;
  });

  var allBounds = L.featureGroup(records.map(function(r) { return r._marker; }));
  if (records.length > 0) {
    map.fitBounds(allBounds.getBounds().pad(0.1));
  }

  function popupContent(r) {
    var div = document.createElement("div");
    var title = document.createElement("div");
    title.style.fontWeight = "600";
    title.style.marginBottom = "4px";
    title.textContent = r.address;
    div.appendChild(title);

    var sub = document.createElement("div");
    sub.style.marginBottom = "4px";
    sub.style.color = "#52514e";
    sub.textContent = r.submarket;
    div.appendChild(sub);

    infoCols.forEach(function(col) {
      if (r.info[col] === undefined) return;
      var line = document.createElement("div");
      line.style.fontSize = "0.85em";
      line.textContent = col + ": " + r.info[col];
      div.appendChild(line);
    });
    return div;
  }

  // ── filter panel ────────────────────────────────────────────────────────
  var filtersEl = document.getElementById("sm-filters");

  var actions = document.createElement("div");
  actions.className = "sm-filter-actions";
  var allLink = document.createElement("a");
  allLink.textContent = "All";
  var noneLink = document.createElement("a");
  noneLink.textContent = "None";
  actions.appendChild(allLink);
  actions.appendChild(noneLink);
  filtersEl.appendChild(actions);

  var checkboxes = {};
  categories.forEach(function(cat) {
    var count = records.filter(function(r) { return r.submarket === cat; }).length;

    var row = document.createElement("div");
    row.className = "sm-filter-row";

    var swatch = document.createElement("span");
    swatch.className = "sm-swatch";
    swatch.style.background = colors[cat];
    row.appendChild(swatch);

    var checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = true;
    checkbox.id = "sm-cat-" + cat.replace(/[^a-zA-Z0-9]/g, "_");
    checkbox.addEventListener("change", function() {
      activeCats[cat] = checkbox.checked;
      applyFilter();
    });
    row.appendChild(checkbox);

    var label = document.createElement("label");
    label.setAttribute("for", checkbox.id);
    label.textContent = cat;
    row.appendChild(label);

    var countEl = document.createElement("span");
    countEl.className = "sm-cat-count";
    countEl.textContent = count;
    row.appendChild(countEl);

    filtersEl.appendChild(row);
    checkboxes[cat] = checkbox;
  });

  allLink.addEventListener("click", function() {
    categories.forEach(function(cat) { checkboxes[cat].checked = true; activeCats[cat] = true; });
    applyFilter();
  });
  noneLink.addEventListener("click", function() {
    categories.forEach(function(cat) { checkboxes[cat].checked = false; activeCats[cat] = false; });
    applyFilter();
  });

  // ── table ──────────────────────────────────────────────────────────────
  var theadRow = document.getElementById("sm-thead-row");
  var columns = ["Address", "Submarket"].concat(infoCols);
  columns.forEach(function(col) {
    var th = document.createElement("th");
    th.textContent = col;
    theadRow.appendChild(th);
  });

  var tbody = document.getElementById("sm-tbody");
  var countEl2 = document.getElementById("sm-count");

  function buildTable(visibleRecords) {
    tbody.innerHTML = "";
    if (visibleRecords.length === 0) {
      var tr = document.createElement("tr");
      var td = document.createElement("td");
      td.colSpan = columns.length;
      td.className = "sm-empty";
      td.textContent = "No pins match the current filters.";
      tr.appendChild(td);
      tbody.appendChild(tr);
      return;
    }
    visibleRecords.forEach(function(r) {
      var tr = document.createElement("tr");
      tr.className = "sm-row";
      tr.id = "sm-row-" + r.i;
      tr.addEventListener("click", function() {
        map.flyTo([r.lat, r.lon], Math.max(map.getZoom(), 15));
        r._marker.openPopup();
        highlightRow(r.i);
      });

      var tdAddr = document.createElement("td");
      tdAddr.textContent = r.address;
      tr.appendChild(tdAddr);

      var tdSub = document.createElement("td");
      tdSub.textContent = r.submarket;
      tr.appendChild(tdSub);

      infoCols.forEach(function(col) {
        var td = document.createElement("td");
        td.textContent = r.info[col] !== undefined ? r.info[col] : "";
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });
  }

  function highlightRow(i) {
    var prev = tbody.querySelector("tr.sm-active");
    if (prev) prev.classList.remove("sm-active");
    var row = document.getElementById("sm-row-" + i);
    if (row) {
      row.classList.add("sm-active");
      row.scrollIntoView({ block: "nearest" });
    }
  }

  function applyFilter() {
    var visible = [];
    records.forEach(function(r) {
      var show = !!activeCats[r.submarket];
      if (show) {
        if (!map.hasLayer(r._marker)) r._marker.addTo(map);
        visible.push(r);
      } else {
        if (map.hasLayer(r._marker)) map.removeLayer(r._marker);
      }
    });
    buildTable(visible);
    countEl2.textContent = visible.length + " of " + records.length + " pins shown";
  }

  records.forEach(function(r) { r._marker.addTo(map); });
  applyFilter();
})();
{% endmacro %}
""")


def main():
    df = load_rows()
    info_cols = info_columns(df)

    print(f"Loaded {len(df)} row(s) from {EXCEL_FILE}")
    print(f"Geocoding {df[ADDRESS_COL].nunique()} unique address(es)...")
    coords = geocode_column(df[ADDRESS_COL].tolist())

    records = build_records(df, coords, info_cols)
    print(f"Plotting {len(records)}/{len(df)} row(s) (rows that failed to geocode are skipped).")

    if not records:
        print("Nothing to plot — no addresses were successfully geocoded.")
        return

    colors = assign_colors(records)

    m = folium.Map(location=[records[0]["lat"], records[0]["lon"]], zoom_start=ZOOM_START, tiles=MAP_TILE)
    m.add_child(SubmarketDashboard(records, info_cols, colors, SIDEBAR_WIDTH_PX))
    m.save(OUTPUT_PATH)
    print(f"Saved interactive dashboard to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
