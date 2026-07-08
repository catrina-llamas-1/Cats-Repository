"""
Plot address pins on an interactive folium map.

Give it a list of addresses (directly, or via a CSV/Excel file) and it
geocodes each one and drops a pin on an interactive HTML map you can open
in any browser and pan/zoom/click around.

Usage:
    pip install folium geopy pandas openpyxl
    python plot_address_map.py

Edit the CONFIG section below to supply your addresses.
"""

import os
import pandas as pd
import folium
from folium.plugins import MarkerCluster
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter

# ── CONFIG ────────────────────────────────────────────────────────────────────
ADDRESSES = [
    # "10230 Jasper Ave, Edmonton, AB",
    # "West Edmonton Mall, Edmonton, AB",
]

ADDRESS_FILE = None        # e.g. "addresses.xlsx" — set this to pull addresses from a file instead of the list above
ADDRESS_COL  = "address"   # column in ADDRESS_FILE holding the addresses
LABEL_COL    = None        # optional column in ADDRESS_FILE for popup labels (defaults to the address itself)

OUTPUT_PATH  = "address_map.html"
MAP_TILE     = "OpenStreetMap"
ZOOM_START   = 12
CLUSTER_PINS = True         # group nearby pins into clusters when zoomed out

GEOCODER_USER_AGENT = "cats-repository-address-mapper"
CACHE_PATH           = "geocode_cache.csv"   # remembers past lookups so reruns don't re-geocode everything
# ─────────────────────────────────────────────────────────────────────────────


def load_addresses() -> list[tuple[str, str]]:
    """Return a list of (address, popup_label) pairs from ADDRESS_FILE or ADDRESSES."""
    if ADDRESS_FILE:
        ext = os.path.splitext(ADDRESS_FILE)[-1].lower()
        df = pd.read_csv(ADDRESS_FILE, dtype=str) if ext == ".csv" else pd.read_excel(ADDRESS_FILE, dtype=str)
        df = df.dropna(subset=[ADDRESS_COL])
        addresses = df[ADDRESS_COL].tolist()
        labels = df[LABEL_COL].tolist() if LABEL_COL and LABEL_COL in df.columns else addresses
        return list(zip(addresses, labels))
    return [(a, a) for a in ADDRESSES if a and a.strip()]


def load_cache() -> dict:
    if os.path.exists(CACHE_PATH):
        return pd.read_csv(CACHE_PATH, dtype=str).set_index("address").to_dict("index")
    return {}


def save_cache(cache: dict) -> None:
    pd.DataFrame.from_dict(cache, orient="index").rename_axis("address").reset_index().to_csv(CACHE_PATH, index=False)


def geocode_addresses(pairs: list[tuple[str, str]]) -> list[tuple[str, str, float, float]]:
    """Geocode each address, using/populating a local cache to avoid repeat lookups."""
    geolocator = Nominatim(user_agent=GEOCODER_USER_AGENT)
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1)

    cache = load_cache()
    results = []
    failed = []

    for address, label in pairs:
        cached = cache.get(address)
        if cached is not None:
            lat, lon = cached.get("lat"), cached.get("lon")
            if pd.isna(lat) or lat in ("", None):
                failed.append(address)
                continue
            results.append((address, label, float(lat), float(lon)))
            continue

        location = geocode(address)
        if location:
            cache[address] = {"lat": location.latitude, "lon": location.longitude}
            results.append((address, label, location.latitude, location.longitude))
        else:
            cache[address] = {"lat": "", "lon": ""}
            failed.append(address)

    save_cache(cache)

    if failed:
        print(f"Could not geocode {len(failed)} address(es):")
        for a in failed:
            print(f"  - {a}")

    return results


def build_map(results: list[tuple[str, str, float, float]]) -> folium.Map:
    if not results:
        raise ValueError("No addresses were successfully geocoded.")

    avg_lat = sum(r[2] for r in results) / len(results)
    avg_lon = sum(r[3] for r in results) / len(results)

    m = folium.Map(location=[avg_lat, avg_lon], zoom_start=ZOOM_START, tiles=MAP_TILE)
    target = MarkerCluster(name="Pins").add_to(m) if CLUSTER_PINS else m

    for address, label, lat, lon in results:
        folium.Marker(
            location=[lat, lon],
            popup=folium.Popup(label, max_width=300),
            tooltip=label,
            icon=folium.Icon(color="red", icon="map-pin", prefix="fa"),
        ).add_to(target)

    if CLUSTER_PINS:
        folium.LayerControl().add_to(m)

    return m


def main():
    pairs = load_addresses()
    if not pairs:
        print("No addresses found. Add some to ADDRESSES or point ADDRESS_FILE at a file.")
        return

    print(f"Geocoding {len(pairs)} address(es)...")
    results = geocode_addresses(pairs)
    print(f"Successfully geocoded {len(results)}/{len(pairs)} address(es).")

    m = build_map(results)
    m.save(OUTPUT_PATH)
    print(f"Saved interactive map to: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
