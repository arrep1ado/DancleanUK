import io
import math
import random
import sqlite3
import time
from datetime import datetime, date
from functools import lru_cache
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# ============================================================
# DAN CLEAN UK - DAILY ROUTE OPTIMIZER
# Version 14.0
# ============================================================

APP_VERSION = "25.3"
DB_FILE = "dancleanuk.db"

st.set_page_config(
    page_title="DanCleanUK Route Optimizer",
    page_icon="🚗",
    layout="centered",
)

st.title("🚗 DanCleanUK Daily Route Optimizer")

st.markdown(
    """
    <style>
        html, body { overscroll-behavior-y: none; }
        .small-muted { color: #777; font-size: 0.9rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# DATABASE / PERSISTENCE
# ============================================================

def db_connect():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = db_connect()
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS jobs (
            job_id TEXT PRIMARY KEY,
            service_date TEXT NOT NULL,
            postcode TEXT NOT NULL,
            price REAL NOT NULL,
            phone TEXT NOT NULL,
            status TEXT NOT NULL DEFAULT 'pending',
            payment TEXT NOT NULL DEFAULT 'Waiting',
            payment_time TEXT,
            completed_time TEXT,
            route_order INTEGER,
            address_text TEXT,
            latitude REAL,
            longitude REAL,
            geo_query TEXT,
            created_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    conn.close()


def save_job(row):
    conn = db_connect()
    conn.execute(
        """
        INSERT INTO jobs (
            job_id, service_date, postcode, price, phone,
            status, payment, payment_time, completed_time,
            route_order, address_text, latitude, longitude,
            geo_query, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(job_id) DO UPDATE SET
            service_date=excluded.service_date,
            postcode=excluded.postcode,
            price=excluded.price,
            phone=excluded.phone,
            status=excluded.status,
            payment=excluded.payment,
            payment_time=excluded.payment_time,
            completed_time=excluded.completed_time,
            route_order=excluded.route_order,
            address_text=excluded.address_text,
            latitude=excluded.latitude,
            longitude=excluded.longitude,
            geo_query=excluded.geo_query
        """,
        (
            str(row["job_id"]),
            str(row["service_date"]),
            str(row["Postcode"]),
            float(row["Price"]),
            str(row["Phone"]),
            str(row["Status"]),
            str(row["Payment"]),
            clean_optional(row.get("PaymentTime")),
            clean_optional(row.get("CompletedTime")),
            clean_optional(row.get("route_order")),
            str(row.get("address_text", "")),
            safe_float(row.get("latitude")),
            safe_float(row.get("longitude")),
            str(row.get("geo_query", "")),
            str(row.get("created_at", datetime.now().isoformat())),
        ),
    )
    conn.commit()
    conn.close()


def save_dataframe(df):
    if df is None or df.empty:
        return
    for _, row in df.iterrows():
        save_job(row)


def load_day(service_date):
    conn = db_connect()
    rows = conn.execute(
        """
        SELECT *
        FROM jobs
        WHERE service_date = ?
        ORDER BY
            CASE WHEN route_order IS NULL THEN 1 ELSE 0 END,
            route_order,
            created_at
        """,
        (service_date,),
    ).fetchall()
    conn.close()

    if not rows:
        return pd.DataFrame()

    loaded = pd.DataFrame([dict(row) for row in rows])

    # SQLite column names are lowercase, while the rest of the app
    # consistently uses the display/import column names.  Normalise
    # persisted records here so restored days behave exactly like
    # freshly uploaded jobs.
    loaded = loaded.rename(
        columns={
            "postcode": "Postcode",
            "price": "Price",
            "phone": "Phone",
            "status": "Status",
            "payment": "Payment",
            "payment_time": "PaymentTime",
            "completed_time": "CompletedTime",
        }
    )

    # Keep all fields expected by the UI present even if an older
    # database/version did not contain a value.
    defaults = {
        "Status": "pending",
        "Payment": "Waiting",
        "PaymentTime": "",
        "CompletedTime": "",
        "route_order": None,
        "address_text": "",
        "latitude": None,
        "longitude": None,
        "geo_query": "",
    }
    for column, default in defaults.items():
        if column not in loaded.columns:
            loaded[column] = default

    loaded["Status"] = loaded["Status"].fillna("pending")
    loaded["Payment"] = loaded["Payment"].fillna("Waiting")
    loaded["PaymentTime"] = loaded["PaymentTime"].fillna("")
    loaded["CompletedTime"] = loaded["CompletedTime"].fillna("")
    loaded["address_text"] = loaded["address_text"].fillna("")
    loaded["geo_query"] = loaded["geo_query"].fillna("")

    return loaded


def delete_day(service_date):
    conn = db_connect()
    conn.execute(
        "DELETE FROM jobs WHERE service_date = ?",
        (service_date,),
    )
    conn.commit()
    conn.close()


init_db()


# ============================================================
# SESSION STATE
# ============================================================

if st.session_state.get("app_version") != APP_VERSION:
    old_cache = st.session_state.get("geocode_cache", {})
    st.session_state.clear()
    st.session_state.app_version = APP_VERSION
    st.session_state.geocode_cache = old_cache

if "geocode_cache" not in st.session_state:
    st.session_state.geocode_cache = {}

if "service_date" not in st.session_state:
    st.session_state.service_date = date.today().isoformat()


# ============================================================
# HELPERS
# ============================================================

def clean_val(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value).strip()

    if text.endswith(".0"):
        try:
            text = str(int(float(text)))
        except Exception:
            pass

    return text


def clean_optional(value):
    text = clean_val(value)
    return text if text else None


def safe_float(value):
    try:
        if value is None or pd.isna(value):
            return None
        return float(value)
    except Exception:
        return None


def normalise_postcode(value):
    return clean_val(value).upper().replace("  ", " ")


def normalise_phone(value):
    text = clean_val(value)
    if not text:
        return ""
    text = text.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    if text.startswith("07"):
        text = "44" + text[1:]
    elif text.startswith("+44"):
        text = text[1:]
    return text


def maps_url(destination):
    return (
        "https://www.google.com/maps/dir/?api=1"
        f"&destination={quote(str(destination))}"
        "&travelmode=driving"
    )


def format_duration(seconds):
    minutes = max(0, round(float(seconds) / 60))
    hours = minutes // 60
    mins = minutes % 60
    if hours:
        return f"{hours}h {mins}m"
    return f"{mins}m"


def make_job_id(service_date, row_number, postcode, phone):
    raw = f"{service_date}|{row_number}|{postcode}|{phone}"
    import hashlib
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ============================================================
# SETTINGS
# ============================================================

st.sidebar.title("⚙️ Settings")

service_date = st.sidebar.date_input(
    "Route date",
    value=date.fromisoformat(st.session_state.service_date),
)
service_date_str = service_date.isoformat()

if service_date_str != st.session_state.service_date:
    st.session_state.service_date = service_date_str
    st.session_state.pop("master_df", None)
    st.session_state.pop("route_data", None)
    st.rerun()

DEPOT_POSTCODE = st.sidebar.text_input(
    "Depot Postcode",
    value="NG31 9RA",
)

DEPOT_FULL_ADDRESS = st.sidebar.text_input(
    "Depot Address",
    value="192 Queensway, Grantham NG31 9RA",
)

FUEL_PRICE = st.sidebar.number_input(
    "Fuel Price (£/litre)",
    min_value=0.01,
    value=1.50,
    step=0.01,
)

MPG = st.sidebar.number_input(
    "Vehicle MPG",
    min_value=1.0,
    value=30.0,
    step=0.1,
)

TAX_RATE = (
    st.sidebar.slider(
        "Tax Deduction (%)",
        0,
        50,
        20,
    )
    / 100
)

st.sidebar.markdown("---")
st.sidebar.subheader("🏦 Payment Settings")

BUSINESS_NAME = st.sidebar.text_input(
    "Business name",
    value="DanCleanUK",
)

BANK_NAME = st.sidebar.text_input(
    "Bank name",
    value="Mettle",
)

SORT_CODE = st.sidebar.text_input(
    "Sort code",
    value="04-03-33",
)

ACCOUNT_NUMBER = st.sidebar.text_input(
    "Account number",
    value="72515806",
)

st.sidebar.markdown("---")
st.sidebar.subheader("🧠 Route Priorities")

TIME_PRIORITY = st.sidebar.slider(
    "Driving Time Priority",
    1,
    10,
    10,
)

DISTANCE_PRIORITY = st.sidebar.slider(
    "Distance/Fuel Priority",
    1,
    10,
    7,
)

CLUSTER_PRIORITY = st.sidebar.slider(
    "Stay Near Nearby Jobs",
    1,
    10,
    9,
)

st.sidebar.caption(
    "Higher values make that factor more important. "
    "The optimiser considers multiple possible starting customers "
    "instead of forcing the closest customer to be first."
)


# ============================================================
# API KEY
# ============================================================

if "API_KEY" not in st.secrets:
    st.error("API_KEY missing in Streamlit secrets.")
    st.info("Add API_KEY to your Streamlit secrets.")
    st.stop()

API_KEY = st.secrets["API_KEY"]


# ============================================================
# ADDRESS / GEOCODING
# ============================================================

def build_geo_query(row, default_postcode):
    parts = []

    address_columns = {
        "address",
        "street",
        "location",
        "house",
        "house number",
        "house_number",
        "property",
        "property address",
        "address line 1",
        "address1",
        "address line",
        "address2",
        "town",
        "city",
    }

    for column in row.index:
        if str(column).strip().lower() in address_columns:
            value = clean_val(row[column])
            if value and value not in parts:
                parts.append(value)

    postcode = normalise_postcode(row.get("Postcode", ""))

    if postcode:
        parts.append(postcode)
    else:
        parts.append(default_postcode)

    parts.append("United Kingdom")
    return ", ".join(parts)


def get_address_text(row):
    parts = []

    for column in row.index:
        name = str(column).strip().lower()
        if name in {
            "address",
            "street",
            "location",
            "house",
            "house number",
            "house_number",
            "property",
            "property address",
            "address line 1",
            "address1",
            "address line",
            "address2",
            "town",
            "city",
        }:
            value = clean_val(row.get(column))
            if value and value not in parts:
                parts.append(value)

    postcode = normalise_postcode(row.get("Postcode", ""))
    if postcode:
        parts.append(postcode)

    return ", ".join(parts)


def cache_key_for(query, postcode):
    return (str(query).strip() + "|" + str(postcode).strip()).lower()


def geocode_candidates(query, postcode):
    """Build several sensible geocoding queries.

    Older/terminated postcodes are a particular problem in Grantham.  A
    postcode can be perfectly valid historical customer data but no longer
    be returned by postcodes.io.  In that situation we deliberately try the
    street + town before giving up.
    """
    query = str(query or "").strip()
    postcode = normalise_postcode(postcode)

    candidates = []

    def add(value):
        value = str(value or "").strip()
        if value and value not in candidates:
            candidates.append(value)

    add(query)

    # If the imported address contains a postcode, remove it and explicitly
    # add Grantham. This is much more reliable for old Grantham postcodes.
    street_part = query
    if postcode:
        street_part = street_part.replace(postcode, "").strip(" ,")
        street_part = street_part.replace(postcode.replace(" ", ""), "").strip(" ,")

    if street_part:
        add(f"{street_part}, Grantham, Lincolnshire, United Kingdom")
        add(f"{street_part}, Grantham, United Kingdom")

    if postcode:
        add(f"{postcode}, Grantham, Lincolnshire, United Kingdom")
        add(f"{postcode}, Grantham, United Kingdom")
        add(f"{postcode}, United Kingdom")

    return candidates


def nominatim_search(query, headers):
    """Query Nominatim with a couple of retries and basic validation."""
    for attempt in range(2):
        try:
            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": query,
                    "format": "json",
                    "limit": 3,
                    "countrycodes": "gb",
                    "addressdetails": 1,
                },
                headers=headers,
                timeout=15,
            )

            if response.status_code == 200:
                data = response.json()
                if data:
                    # Prefer results that look like a UK/Grantham address.
                    for item in data:
                        try:
                            lat = float(item["lat"])
                            lon = float(item["lon"])
                        except Exception:
                            continue

                        if -90 <= lat <= 90 and -180 <= lon <= 180:
                            return (lat, lon)

            if response.status_code in (429, 500, 502, 503, 504):
                time.sleep(1.5 * (attempt + 1))
                continue

        except Exception:
            if attempt == 0:
                time.sleep(1.0)

    return None


# These are legacy Grantham postcodes that are no longer in use but may still
# appear on genuine customer records.  The coordinates are postcode-area
# coordinates, used only after the live geocoders fail.  This prevents a real
# customer being silently removed from the route.
LEGACY_POSTCODE_COORDS = {
    "NG31 7AN": (52.909806, -0.640572),
    "NG31 9EH": (52.909052, -0.630469),
}


def get_coords(query_string, postcode):
    """Locate a customer using address-first geocoding with fallbacks.

    V24 changes only the geocoding priority used by V23:
      1. cached successful result
      2. Nominatim full address / street candidates
      3. postcodes.io postcode centroid
      4. known legacy-postcode coordinate fallback

    The important difference is that customers sharing a postcode are now
    given a chance to receive different coordinates when their full street
    addresses are known.  This is especially important for multiple houses
    on the same postcode, such as 180/190/194/196 Queensway.
    """
    query = str(query_string or "").strip()
    postcode = normalise_postcode(postcode)
    key = cache_key_for(query, postcode)

    if key in st.session_state.geocode_cache:
        return st.session_state.geocode_cache[key]

    # --------------------------------------------------------
    # 1-3. NOMINATIM FIRST - FULL ADDRESS / STREET CANDIDATES
    # --------------------------------------------------------
    # Try the complete customer address before falling back to a postcode
    # centroid.  This prevents several different houses in the same postcode
    # from automatically sharing one coordinate.
    headers = {
        "User-Agent": "DanCleanUKRouteOptimizer/24.0"
    }

    for candidate in geocode_candidates(query, postcode):
        coords = nominatim_search(candidate, headers)
        if coords is not None:
            st.session_state.geocode_cache[key] = coords
            return coords

        # Public Nominatim service asks clients to be considerate.  Keep the
        # existing short spacing between uncached requests.
        time.sleep(0.35)

    # --------------------------------------------------------
    # 4. POSTCODES.IO POSTCODE FALLBACK
    # --------------------------------------------------------
    # If the exact address cannot be located, use the normal postcode
    # coordinate rather than losing the customer from the route.
    if postcode:
        for pc in dict.fromkeys([postcode, postcode.replace(" ", "")]):
            try:
                response = requests.get(
                    f"https://api.postcodes.io/postcodes/{quote(pc)}",
                    timeout=10,
                )
                if response.status_code == 200:
                    result = response.json().get("result")
                    if result:
                        lat = result.get("latitude")
                        lon = result.get("longitude")
                        if lat is not None and lon is not None:
                            coords = (float(lat), float(lon))
                            st.session_state.geocode_cache[key] = coords
                            return coords
            except Exception:
                pass

    # --------------------------------------------------------
    # 5. LEGACY POSTCODE FALLBACK
    # --------------------------------------------------------
    if postcode in LEGACY_POSTCODE_COORDS:
        coords = LEGACY_POSTCODE_COORDS[postcode]
        st.session_state.geocode_cache[key] = coords
        return coords

    return None


# ============================================================
# ROUTING
# ============================================================

def haversine_km(lat1, lon1, lat2, lon2):
    radius = 6371.0
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)

    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(math.radians(lat1))
        * math.cos(math.radians(lat2))
        * math.sin(dlon / 2) ** 2
    )

    return radius * 2 * math.asin(math.sqrt(a))


def haversine_points(point_a, point_b):
    """Calculate distance between two [longitude, latitude] points."""
    return haversine_km(
        point_a[1], point_a[0],
        point_b[1], point_b[0],
    )


def location_cache_key(locations):
    """Convert coordinate lists into a hashable cache key."""
    return tuple(
        (float(point[0]), float(point[1]))
        for point in locations
    )


def offline_matrix(locations):
    n = len(locations)
    distances = [[0.0] * n for _ in range(n)]
    durations = [[0.0] * n for _ in range(n)]

    road_factor = 1.30
    average_speed = 35.0

    for i in range(n):
        lon1, lat1 = locations[i]

        for j in range(n):
            if i == j:
                continue

            lon2, lat2 = locations[j]
            km = haversine_km(lat1, lon1, lat2, lon2)
            road_km = km * road_factor

            distances[i][j] = road_km * 1000
            durations[i][j] = road_km / average_speed * 3600

    return distances, durations


def get_ors_matrix(locations):
    try:
        response = requests.post(
            "https://api.openrouteservice.org/v2/matrix/driving-car",
            json={
                "locations": locations,
                "metrics": ["distance", "duration"],
                "units": "m",
            },
            headers={
                "Authorization": API_KEY,
                "Content-Type": "application/json",
            },
            timeout=90,
        )

        if response.status_code != 200:
            return None, None

        data = response.json()
        distances = data.get("distances")
        durations = data.get("durations")

        if not distances or not durations:
            return None, None

        return distances, durations

    except Exception:
        return None, None


def route_metrics(route, distances, durations, fuel_price, mpg):
    total_distance = 0.0
    total_time = 0.0

    for i in range(len(route) - 1):
        a = route[i]
        b = route[i + 1]
        total_distance += distances[a][b]
        total_time += durations[a][b]

    miles = total_distance / 1000 * 0.621371
    litres = miles / mpg * 4.54609
    fuel_cost = litres * fuel_price

    return {
        "distance_m": total_distance,
        "time_s": total_time,
        "miles": miles,
        "litres": litres,
        "fuel_cost": fuel_cost,
    }


def route_score(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    locations=None,
):
    """
    Score a complete route.

    The important change in v14 is that route shape is now treated as a
    first-class objective.  The previous optimiser could find a route with
    reasonable total mileage while still jumping out of one dense area and
    then coming back later.  That is exactly the behaviour we want to avoid.

    We therefore score:
      1. real road driving time
      2. real road driving distance
      3. strong continuity / backtracking penalties
      4. a small geographic-shape penalty when coordinates are available
    """
    metrics = route_metrics(
        route,
        distances,
        durations,
        fuel_price,
        mpg,
    )

    driving_minutes = metrics["time_s"] / 60.0
    driving_miles = metrics["miles"]

    continuity = calculate_continuity_penalty(route, distances)

    zone_penalty = 0.0
    if locations is not None:
        zone_penalty = calculate_zone_penalty(route, locations)

    shape_penalty = 0.0
    backtrack_penalty = 0.0
    if locations is not None:
        shape_penalty = calculate_shape_penalty(route, locations)
        backtrack_penalty = calculate_geographic_backtracking_penalty(route, locations)

    return (
        driving_minutes * TIME_PRIORITY
        + driving_miles * DISTANCE_PRIORITY
        + continuity * CLUSTER_PRIORITY * 10.0
        + zone_penalty * CLUSTER_PRIORITY * 1.80
        + shape_penalty * CLUSTER_PRIORITY * 0.90
        + backtrack_penalty * CLUSTER_PRIORITY * 1.80
    )


def calculate_continuity_penalty(route, distances):
    """
    Strongly discourage leaving a dense local group while nearby work remains.

    The value returned is expressed in miles so the final route score remains
    easy to reason about.
    """
    if len(route) <= 3:
        return 0.0

    customers = set(route[1:-1])
    penalty = 0.0

    for pos in range(1, len(route) - 1):
        current = route[pos]
        next_stop = route[pos + 1]
        remaining = customers.difference(route[:pos + 1])

        if not remaining:
            continue

        next_miles = distances[current][next_stop] / 1609.344
        remaining_miles = {
            job: distances[current][job] / 1609.344
            for job in remaining
        }

        nearby_2 = [m for m in remaining_miles.values() if m <= 2.0]
        nearby_4 = [m for m in remaining_miles.values() if m <= 4.0]
        nearby_6 = [m for m in remaining_miles.values() if m <= 6.0]

        # If there are several genuinely nearby jobs, jumping away from them
        # is a strong sign of a poor route shape.
        if len(nearby_2) >= 1 and next_miles > 3.0:
            penalty += (next_miles - 3.0) * (1.0 + 0.65 * len(nearby_2))

        if len(nearby_4) >= 2 and next_miles > 5.0:
            penalty += (next_miles - 5.0) * (1.5 + 0.45 * len(nearby_4))

        if len(nearby_6) >= 3 and next_miles > 7.0:
            penalty += (next_miles - 7.0) * (1.8 + 0.30 * len(nearby_6))

        # More generally, compare the chosen next stop with the nearest
        # remaining stop.  A very large ratio means we are skipping local work.
        nearest = min(remaining_miles.values())
        if next_miles > max(nearest * 1.75, nearest + 1.5):
            excess = next_miles - max(nearest * 1.75, nearest + 1.5)
            penalty += excess * 2.5

    return penalty


def calculate_shape_penalty(route, locations):
    """
    Small geometric penalty for routes that repeatedly reverse direction.

    This is deliberately secondary to live road time.  It is only there to
    prefer a natural sweep through the work area when two routes are otherwise
    similar.
    """
    if len(route) < 5:
        return 0.0

    def bearing(a, b):
        lon1, lat1 = locations[a]
        lon2, lat2 = locations[b]
        y = math.sin(math.radians(lon2 - lon1)) * math.cos(math.radians(lat2))
        x = (
            math.cos(math.radians(lat1)) * math.sin(math.radians(lat2))
            - math.sin(math.radians(lat1))
            * math.cos(math.radians(lat2))
            * math.cos(math.radians(lon2 - lon1))
        )
        angle = math.degrees(math.atan2(y, x))
        return (angle + 360.0) % 360.0

    bearings = []
    for i in range(1, len(route) - 1):
        bearings.append(bearing(route[i - 1], route[i]))

    penalty = 0.0
    for i in range(1, len(bearings)):
        delta = abs(bearings[i] - bearings[i - 1])
        delta = min(delta, 360.0 - delta)
        if delta > 115:
            penalty += (delta - 115) / 45.0

    return penalty


def calculate_geographic_backtracking_penalty(route, locations):
    """Apply a moderate penalty when the route moves back toward the depot
    while useful work remains farther out in the same general direction.

    This is intentionally softer than the old zone-transition logic. It uses
    actual coordinates rather than postcode groups, so a road layout can still
    justify a turn without the optimiser being forced into a rigid zone order.
    """
    if len(route) < 5:
        return 0.0

    depot = locations[0]

    def bearing_from_depot(index):
        lon, lat = locations[index]
        dlon = (lon - depot[0]) * math.cos(math.radians(depot[1]))
        dlat = lat - depot[1]
        return (math.degrees(math.atan2(dlon, dlat)) + 360.0) % 360.0

    def radius_from_depot(index):
        return haversine_points(depot, locations[index])

    penalty = 0.0

    for pos in range(1, len(route) - 1):
        current = route[pos]
        nxt = route[pos + 1]
        current_radius = radius_from_depot(current)
        next_radius = radius_from_depot(nxt)

        # Only consider a meaningful move back toward the depot.
        radial_backtrack = current_radius - next_radius
        if radial_backtrack < 1.5:
            continue

        current_bearing = bearing_from_depot(current)
        next_bearing = bearing_from_depot(nxt)

        # Is there still unvisited work farther out in roughly the same
        # direction? If so, returning inward is more likely to be genuine
        # route backtracking rather than a necessary local road turn.
        remaining = route[pos + 1:-1]
        for other in remaining:
            if other == nxt:
                continue
            other_radius = radius_from_depot(other)
            if other_radius <= current_radius + 2.0:
                continue

            other_bearing = bearing_from_depot(other)
            delta = abs(other_bearing - next_bearing)
            delta = min(delta, 360.0 - delta)

            if delta <= 55.0:
                # Scale gently: the optimiser should prefer progress, but
                # real road time/distance still dominate.
                penalty += min(radial_backtrack, 8.0) * 0.55
                break

        # Penalise a sharp reversal between consecutive legs, but only when
        # it is also accompanied by radial backtracking.
        if pos >= 2:
            previous = route[pos - 1]
            a = locations[previous]
            b = locations[current]
            c = locations[nxt]

            def leg_bearing(p1, p2):
                lon1, lat1 = p1
                lon2, lat2 = p2
                y = math.sin(math.radians(lon2 - lon1)) * math.cos(math.radians(lat2))
                x = (
                    math.cos(math.radians(lat1)) * math.sin(math.radians(lat2))
                    - math.sin(math.radians(lat1)) * math.cos(math.radians(lat2))
                    * math.cos(math.radians(lon2 - lon1))
                )
                return (math.degrees(math.atan2(y, x)) + 360.0) % 360.0

            first = leg_bearing(a, b)
            second = leg_bearing(b, c)
            turn = abs(second - first)
            turn = min(turn, 360.0 - turn)
            if turn > 120.0 and radial_backtrack > 2.0:
                penalty += (turn - 120.0) / 35.0

    return penalty


@lru_cache(maxsize=32)
def geographic_zone_labels(location_tuple):
    """Create stable geographic zones from customer coordinates.

    Zones are calculated from the actual customer spread rather than from
    postcode text, so the optimiser still works when the user imports a
    completely different day's jobs.
    """
    locations = list(location_tuple)
    customer_count = len(locations) - 1
    if customer_count <= 0:
        return tuple()

    if customer_count <= 12:
        k = 3
    elif customer_count <= 24:
        k = 4
    elif customer_count <= 40:
        k = 5
    else:
        k = 6
    k = min(k, customer_count)

    # Deterministic farthest-point seeds. This avoids depending on sklearn.
    seeds = [1]
    while len(seeds) < k:
        best_idx = None
        best_dist = -1.0
        for idx in range(1, customer_count + 1):
            if idx in seeds:
                continue
            nearest = min(
                haversine_points(locations[idx], locations[s])
                for s in seeds
            )
            if nearest > best_dist:
                best_dist = nearest
                best_idx = idx
        if best_idx is None:
            break
        seeds.append(best_idx)

    labels = [0] * (customer_count + 1)
    centroids = [locations[i] for i in seeds]

    for _ in range(12):
        changed = False
        for idx in range(1, customer_count + 1):
            distances_to_centroids = [
                haversine_points(locations[idx], c) for c in centroids
            ]
            label = min(range(len(centroids)), key=lambda x: distances_to_centroids[x])
            if labels[idx] != label + 1:
                labels[idx] = label + 1
                changed = True

        new_centroids = []
        for zone in range(1, len(centroids) + 1):
            members = [
                locations[i] for i in range(1, customer_count + 1)
                if labels[i] == zone
            ]
            if members:
                lon = sum(p[0] for p in members) / len(members)
                lat = sum(p[1] for p in members) / len(members)
                new_centroids.append((lon, lat))
            else:
                new_centroids.append(centroids[zone - 1])
        centroids = new_centroids
        if not changed:
            break

    return tuple(labels[1:])


def calculate_zone_penalty(route, locations):
    """Penalise leaving a geographic work zone before clearing it.

    This is deliberately softer than live road time/distance. It prevents
    the optimiser from doing things like Grantham -> NG33 -> NG31 -> NG32
    when a clean geographical sweep is available, without forcing an
    unrealistic postcode-based route.
    """
    if len(route) < 4:
        return 0.0

    labels = geographic_zone_labels(location_cache_key(locations))
    if not labels:
        return 0.0

    def zone(customer):
        return labels[customer - 1]

    penalty = 0.0
    visited_zones = []

    for pos in range(1, len(route) - 1):
        current = route[pos]
        nxt = route[pos + 1]
        current_zone = zone(current)
        next_zone = zone(nxt)

        if current_zone != next_zone:
            remaining_same_zone = any(
                zone(x) == current_zone for x in route[pos + 1:-1]
            )
            if remaining_same_zone:
                penalty += 2.5

            if next_zone in visited_zones:
                penalty += 3.5
            visited_zones.append(next_zone)
        elif current_zone not in visited_zones:
            visited_zones.append(current_zone)

    # A second visit to an already-cleared zone is particularly undesirable.
    for zone_id in set(visited_zones):
        occurrences = visited_zones.count(zone_id)
        if occurrences > 1:
            penalty += (occurrences - 1) * 2.0

    return penalty


def geographic_zone_routes(locations):
    """Build candidate routes which clear dynamically detected areas."""
    customer_count = len(locations) - 1
    if customer_count <= 0:
        return []

    labels = geographic_zone_labels(location_cache_key(locations))
    zones = {}
    for customer in range(1, customer_count + 1):
        zones.setdefault(labels[customer - 1], []).append(customer)

    if len(zones) <= 1:
        return []

    depot = locations[0]

    centroids = {}
    for zone_id, members in zones.items():
        lon = sum(locations[i][0] for i in members) / len(members)
        lat = sum(locations[i][1] for i in members) / len(members)
        centroids[zone_id] = (lon, lat)

    # Create a few sensible zone orders. Starting with the zone nearest the
    # depot is usually good, but testing each possible first zone matters
    # because the depot can sit between two natural work areas.
    zone_ids = list(zones)
    routes = []

    for start_zone in sorted(
        zone_ids,
        key=lambda z: haversine_points(depot, centroids[z])
    ):
        remaining = set(zone_ids)
        remaining.remove(start_zone)
        order = [start_zone]
        current = start_zone
        while remaining:
            next_zone = min(
                remaining,
                key=lambda z: haversine_points(centroids[current], centroids[z])
            )
            order.append(next_zone)
            remaining.remove(next_zone)
            current = next_zone

        for zone_order in (order, list(reversed(order))):
            sequence = []
            for zone_id in zone_order:
                members = zones[zone_id][:]
                # Within each zone, sort by angle around the zone centroid.
                c_lon, c_lat = centroids[zone_id]
                members.sort(
                    key=lambda i: (
                        math.degrees(
                            math.atan2(
                                (locations[i][0] - c_lon) * math.cos(math.radians(c_lat)),
                                locations[i][1] - c_lat,
                            )
                        ) + 360.0
                    ) % 360.0
                )
                sequence.extend(members)
            routes.append([0] + sequence + [0])
            routes.append([0] + list(reversed(sequence)) + [0])

    return routes


def build_greedy_route(
    first_customer,
    distances,
    durations,
    mode="time",
):
    customer_count = len(distances) - 1

    route = [0, first_customer]
    remaining = set(range(1, customer_count + 1))
    remaining.discard(first_customer)
    current = first_customer

    while remaining:
        candidates = []

        for candidate in remaining:
            direct_time = durations[current][candidate]
            direct_distance = distances[current][candidate]

            future = [x for x in remaining if x != candidate]
            if future:
                # Look one step ahead, but also reward candidates which sit
                # inside a dense local group.
                nearest_future = min(
                    future,
                    key=lambda x: durations[candidate][x]
                )
                future_time = durations[candidate][nearest_future]
                future_distance = distances[candidate][nearest_future]

                local_count = sum(
                    1
                    for x in future
                    if distances[candidate][x] <= 8000
                )
            else:
                future_time = durations[candidate][0]
                future_distance = distances[candidate][0]
                local_count = 0

            if mode == "distance":
                score = (
                    direct_distance
                    + future_distance * 0.35
                    - local_count * 1800.0
                )
            elif mode == "balanced":
                score = (
                    direct_time * 0.60
                    + (direct_distance / 10.0) * 0.25
                    + future_time * 0.20
                    - local_count * 900.0
                )
            else:
                score = (
                    direct_time * 0.65
                    + future_time * 0.25
                    - local_count * 950.0
                )

            candidates.append((score, candidate))

        candidates.sort(key=lambda x: x[0])
        next_customer = candidates[0][1]
        route.append(next_customer)
        remaining.remove(next_customer)
        current = next_customer

    route.append(0)
    return route


def angular_sweep_routes(locations):
    """
    Generate natural geographical sweeps around the depot.

    We test several angular offsets and both clockwise and anticlockwise
    directions.  This creates routes which stay in one geographical area
    instead of bouncing between NG31/NG32/NG13 repeatedly.
    """
    customer_count = len(locations) - 1
    if customer_count <= 0:
        return []

    depot_lon, depot_lat = locations[0]

    def angle_and_radius(index):
        lon, lat = locations[index]
        dlon = (lon - depot_lon) * math.cos(math.radians(depot_lat))
        dlat = lat - depot_lat
        angle = (math.degrees(math.atan2(dlon, dlat)) + 360.0) % 360.0
        radius = math.sqrt(dlon * dlon + dlat * dlat)
        return angle, radius

    ordered = list(range(1, customer_count + 1))
    ordered.sort(key=angle_and_radius)

    routes = []
    n = len(ordered)

    # Rotating the sweep is important because the depot is in the middle of
    # the working area.  Test every possible cut, not just the first angle.
    for direction in (1, -1):
        base = ordered if direction == 1 else ordered[::-1]
        for start in range(n):
            sequence = base[start:] + base[:start]
            routes.append([0] + sequence + [0])

    return routes


def directional_sector_routes(locations, distances=None, durations=None):
    """
    Build routes as a real geographical sweep around the depot.

    Instead of asking k-means to decide what a "zone" is, this uses the
    actual bearing of every customer from the depot.  Customers are divided
    into contiguous angular sectors, then those sectors are cleared in one
    direction without deliberately jumping back to an earlier sector.

    Road distance/time is used inside each sector when matrices are available,
    so this is a geographical sweep guided by the road network rather than a
    simple postcode sort.
    """
    customer_count = len(locations) - 1
    if customer_count < 2:
        return []

    depot_lon, depot_lat = locations[0]
    customers = list(range(1, customer_count + 1))

    def bearing_radius(index):
        lon, lat = locations[index]
        dlon = (lon - depot_lon) * math.cos(math.radians(depot_lat))
        dlat = lat - depot_lat
        bearing = (math.degrees(math.atan2(dlon, dlat)) + 360.0) % 360.0
        radius = math.hypot(dlon, dlat)
        return bearing, radius

    info = {i: bearing_radius(i) for i in customers}
    ordered = sorted(customers, key=lambda i: (info[i][0], info[i][1]))
    routes = []

    # Four to six sectors works well for normal daily lists.  The sectors are
    # balanced by number of jobs, which avoids one huge sector and many tiny
    # ones when the jobs are unevenly distributed.
    for sector_count in (4, 5, 6):
        if customer_count < sector_count:
            continue

        base_size = customer_count // sector_count
        remainder = customer_count % sector_count
        sectors = []
        pos = 0
        for sector_id in range(sector_count):
            size = base_size + (1 if sector_id < remainder else 0)
            sectors.append(ordered[pos:pos + size])
            pos += size

        for direction in (1, -1):
            sector_order = list(range(sector_count))
            if direction == -1:
                sector_order.reverse()

            # Try every angular cut. This matters because the depot is not
            # necessarily at the edge of the working area.
            for cut in range(sector_count):
                rotated = sector_order[cut:] + sector_order[:cut]
                sequence = []
                previous = 0

                for sector_id in rotated:
                    members = sectors[sector_id][:]
                    if not members:
                        continue

                    # Primary order is angular.  For the second direction we
                    # reverse it. This keeps the route moving through the
                    # sector instead of zig-zagging across it.
                    members.sort(key=lambda i: (info[i][0], info[i][1]),
                                 reverse=(direction == -1))

                    # Road-aware orientation: compare the cost of entering the
                    # sector at either end and keep the cheaper end first.
                    if distances is not None and len(members) > 1:
                        forward_cost = distances[previous][members[0]]
                        reverse_cost = distances[previous][members[-1]]
                        if reverse_cost < forward_cost:
                            members.reverse()

                    sequence.extend(members)
                    previous = members[-1]

                if sequence:
                    routes.append([0] + sequence + [0])
                    routes.append([0] + list(reversed(sequence)) + [0])

    # Also create a finer sweep by assigning jobs to angular bins from the
    # actual bearing range. This catches cases where one balanced sector cuts
    # through a natural road/settlement boundary.
    for sector_count in (5, 6):
        width = 360.0 / sector_count
        for offset in (0.0, width / 2.0):
            bins = [[] for _ in range(sector_count)]
            for customer in customers:
                angle = (info[customer][0] - offset) % 360.0
                bucket = min(sector_count - 1, int(angle / width))
                bins[bucket].append(customer)

            for direction in (1, -1):
                ids = list(range(sector_count))
                if direction == -1:
                    ids.reverse()
                sequence = []
                previous = 0
                for bucket in ids:
                    members = bins[bucket][:]
                    members.sort(key=lambda i: (info[i][0], info[i][1]),
                                 reverse=(direction == -1))
                    if distances is not None and len(members) > 1:
                        if distances[previous][members[-1]] < distances[previous][members[0]]:
                            members.reverse()
                    sequence.extend(members)
                    if members:
                        previous = members[-1]
                if sequence:
                    routes.append([0] + sequence + [0])

    return routes


def cheapest_insertion_route(
    distances,
    durations,
):
    customer_count = len(distances) - 1

    if customer_count <= 0:
        return [0, 0]

    # Try a few geographically extreme seeds rather than only the furthest.
    seeds = sorted(
        range(1, customer_count + 1),
        key=lambda x: durations[0][x],
        reverse=True,
    )[: min(6, customer_count)]

    best_route = None
    best_value = float("inf")

    for first in seeds:
        route = [0, first, 0]
        remaining = set(range(1, customer_count + 1))
        remaining.remove(first)

        while remaining:
            best_choice = None

            for customer in remaining:
                for position in range(1, len(route)):
                    before = route[position - 1]
                    after = route[position]

                    old_time = durations[before][after]
                    new_time = durations[before][customer] + durations[customer][after]
                    old_distance = distances[before][after]
                    new_distance = distances[before][customer] + distances[customer][after]

                    # Reward inserting beside nearby unvisited work.
                    neighbour_bonus = 0.0
                    for other in remaining:
                        if other == customer:
                            continue
                        if distances[customer][other] <= 8000:
                            neighbour_bonus += 120.0

                    increase = (
                        (new_time - old_time)
                        + (new_distance - old_distance) / 8.0
                        - neighbour_bonus
                    )

                    if best_choice is None or increase < best_choice[0]:
                        best_choice = (increase, customer, position)

            _, customer, position = best_choice
            route.insert(position, customer)
            remaining.remove(customer)

        value = sum(durations[route[i]][route[i + 1]] for i in range(len(route) - 1))
        if value < best_value:
            best_value = value
            best_route = route

    return best_route


def two_opt(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    locations=None,
):
    best = route[:]
    best_score = route_score(best, distances, durations, fuel_price, mpg, locations)

    improved = True
    while improved:
        improved = False

        for i in range(1, len(best) - 2):
            for j in range(i + 1, len(best) - 1):
                candidate = best[:i] + best[i:j + 1][::-1] + best[j + 1:]
                candidate_score = route_score(
                    candidate, distances, durations, fuel_price, mpg, locations
                )

                if candidate_score < best_score - 0.01:
                    best = candidate
                    best_score = candidate_score
                    improved = True
                    break
            if improved:
                break

    return best


def relocate_improvement(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    locations=None,
):
    best = route[:]
    best_score = route_score(best, distances, durations, fuel_price, mpg, locations)

    improved = True
    while improved:
        improved = False

        for i in range(1, len(best) - 1):
            customer = best[i]
            shortened = best[:i] + best[i + 1:]

            for j in range(1, len(shortened)):
                candidate = shortened[:j] + [customer] + shortened[j:]
                candidate_score = route_score(
                    candidate, distances, durations, fuel_price, mpg, locations
                )

                if candidate_score < best_score - 0.01:
                    best = candidate
                    best_score = candidate_score
                    improved = True
                    break

            if improved:
                break

    return best


def swap_improvement(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    locations=None,
):
    best = route[:]
    best_score = route_score(best, distances, durations, fuel_price, mpg, locations)

    improved = True
    while improved:
        improved = False

        for i in range(1, len(best) - 2):
            for j in range(i + 1, len(best) - 1):
                candidate = best[:]
                candidate[i], candidate[j] = candidate[j], candidate[i]
                candidate_score = route_score(
                    candidate, distances, durations, fuel_price, mpg, locations
                )

                if candidate_score < best_score - 0.01:
                    best = candidate
                    best_score = candidate_score
                    improved = True
                    break

            if improved:
                break

    return best


def generate_candidate_routes(distances, durations, locations=None):
    customer_count = len(distances) - 1
    if customer_count <= 0:
        return []

    candidates = []

    # 1. Dynamic geographic-zone routes. These are the backbone of v14:
    # clear one natural area before moving to the next.
    if locations is not None:
        # The main geographical candidates are now true directional sweeps.
        # Keep the older zone/angular candidates as fallbacks so a sweep is
        # never forced when the road network makes another shape genuinely
        # shorter.
        candidates.extend(directional_sector_routes(locations, distances, durations))
        candidates.extend(geographic_zone_routes(locations))
        candidates.extend(angular_sweep_routes(locations))

    # 2. Multi-start greedy routes.  Test all starts for small lists and a
    # useful spread of starts for larger lists.
    starts = list(range(1, customer_count + 1))
    starts.sort(key=lambda x: durations[0][x])

    if customer_count > 40:
        selected = starts[:10]
        selected += starts[-10:]
        selected += starts[:: max(1, customer_count // 10)]
        starts = list(dict.fromkeys(selected))

    for first_customer in starts:
        candidates.append(build_greedy_route(first_customer, distances, durations, "time"))
        candidates.append(build_greedy_route(first_customer, distances, durations, "balanced"))
        candidates.append(build_greedy_route(first_customer, distances, durations, "distance"))

    # 3. Insertion construction.
    insertion = cheapest_insertion_route(distances, durations)
    if insertion:
        candidates.append(insertion)
        if len(insertion) > 3:
            candidates.append([0] + insertion[1:-1][::-1] + [0])

    # Remove exact duplicates while keeping deterministic order.
    unique = []
    seen = set()
    for candidate in candidates:
        key = tuple(candidate)
        if key not in seen:
            seen.add(key)
            unique.append(candidate)

    return unique


def improve_route(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    locations=None,
    preserve_structure=False,
):
    """
    Improve a route without destroying its geographical structure.

    V19 deliberately treats a geographical sweep as a route structure, not
    merely another score penalty.  Ordinary routes may use the full local
    search. Sweep routes only receive safe local improvements, because an
    unrestricted relocate/swap can undo the whole sweep and send the van back
    into an area that was already cleared.
    """
    if preserve_structure:
        return improve_sweep_route(
            route,
            distances,
            durations,
            fuel_price,
            mpg,
            locations,
        )

    improved = two_opt(
        route, distances, durations, fuel_price, mpg, locations
    )
    improved = relocate_improvement(
        improved, distances, durations, fuel_price, mpg, locations
    )
    improved = swap_improvement(
        improved, distances, durations, fuel_price, mpg, locations
    )
    improved = two_opt(
        improved, distances, durations, fuel_price, mpg, locations
    )
    return improved


def improve_sweep_route(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    locations,
):
    """
    Safely improve a sweep route while preserving its geographical direction.

    We do not perform unrestricted relocate/swap operations here.  Instead we
    use adjacent swaps and short reversals only when they improve the route.
    This keeps the broad order of geographical areas intact while allowing
    the live road matrix to tidy up the order inside those areas.
    """
    if not route or locations is None or len(route) < 5:
        return route[:]

    best = route[:]
    best_score = route_score(
        best, distances, durations, fuel_price, mpg, locations
    )

    # A sweep should remain a sweep.  Only make local changes which involve
    # neighbouring stops.  This can remove a bad local zig-zag without moving
    # a customer across the entire route.
    for _ in range(3):
        changed = False

        # Adjacent swaps.
        for i in range(1, len(best) - 2):
            candidate = best[:]
            candidate[i], candidate[i + 1] = candidate[i + 1], candidate[i]
            candidate_score = route_score(
                candidate, distances, durations, fuel_price, mpg, locations
            )
            if candidate_score < best_score - 0.01:
                best = candidate
                best_score = candidate_score
                changed = True

        # Very short 2-stop reversals only.  Never reverse a large section.
        for i in range(1, len(best) - 3):
            candidate = best[:]
            candidate[i:i + 2] = reversed(candidate[i:i + 2])
            candidate_score = route_score(
                candidate, distances, durations, fuel_price, mpg, locations
            )
            if candidate_score < best_score - 0.01:
                best = candidate
                best_score = candidate_score
                changed = True

        if not changed:
            break

    return best


def build_driver_sweep_routes(locations, distances=None, durations=None):
    """Build general-purpose geographic sweep routes.

    The route must work for arbitrary customer locations, so this function
    deliberately does not know anything about Grantham, villages or postcodes.

    Instead it treats the depot as the centre of the working area, creates
    several possible circular geographic sweeps, and chooses the road-aware
    ordering later.  Nearby customers are kept together by angle/radius and
    the route is never forced to return to the depot between territories.
    """
    customer_count = len(locations) - 1
    if customer_count <= 0:
        return []

    depot = locations[0]
    customers = list(range(1, customer_count + 1))

    def bearing(index):
        lon, lat = locations[index]
        dlon = (lon - depot[0]) * math.cos(math.radians(depot[1]))
        dlat = lat - depot[1]
        return (math.degrees(math.atan2(dlon, dlat)) + 360.0) % 360.0

    def radius(index):
        return haversine_points(depot, locations[index])

    radii = {i: radius(i) for i in customers}
    bearings = {i: bearing(i) for i in customers}

    def road_cost(a, b):
        if durations is not None:
            return durations[a][b]
        if distances is not None:
            return distances[a][b]
        return haversine_points(locations[a], locations[b])

    def nearest_first(members, start):
        """Order a small geographic territory without crossing the whole day."""
        remaining = set(members)
        result = []
        current = start
        while remaining:
            nxt = min(
                remaining,
                key=lambda x: (
                    road_cost(current, x),
                    radii[x],
                ),
            )
            result.append(nxt)
            remaining.remove(nxt)
            current = nxt
        return result

    routes = []

    # Generate many possible angular cuts.  This makes the method independent
    # of where the customer's geography happens to sit relative to north.
    cut_count = 18 if customer_count >= 12 else 12
    cuts = [360.0 * i / cut_count for i in range(cut_count)]

    # Test 2, 3 and 4 broad territories.  The number is derived from the job
    # count rather than hard-coded to the current test addresses.
    territory_counts = [2]
    if customer_count >= 8:
        territory_counts.append(3)
    if customer_count >= 16:
        territory_counts.append(4)

    for territory_count in territory_counts:
        width = 360.0 / territory_count

        for cut in cuts:
            bands = [[] for _ in range(territory_count)]
            for job in customers:
                bucket = int(((bearings[job] - cut) % 360.0) / width)
                bucket = min(territory_count - 1, bucket)
                bands[bucket].append(job)

            if any(not band for band in bands):
                continue

            # Try both directions around the depot.  For each direction, try
            # every possible starting territory so the best sweep is not tied
            # to an arbitrary compass direction.
            for direction in (1, -1):
                base = list(range(territory_count))
                if direction == -1:
                    base.reverse()

                for start_pos in range(territory_count):
                    order = base[start_pos:] + base[:start_pos]
                    sequence = []
                    current = 0

                    for band_id in order:
                        members = bands[band_id]
                        forward = nearest_first(members, current)
                        reverse = list(reversed(forward))

                        # Choose the orientation which connects most naturally
                        # from the previous territory.
                        if road_cost(current, reverse[0]) < road_cost(current, forward[0]):
                            chosen = reverse
                        else:
                            chosen = forward

                        sequence.extend(chosen)
                        current = chosen[-1]

                    if len(sequence) == customer_count:
                        routes.append([0] + sequence + [0])

    # Pure polar sweeps are useful when the customers form one broad corridor.
    # They are especially valuable for completely new areas not resembling
    # the current test data.
    for reverse_angle in (False, True):
        angular = sorted(
            customers,
            key=lambda i: (bearings[i], radii[i]),
            reverse=reverse_angle,
        )
        routes.append([0] + angular + [0])

        # Also test the reverse radial order inside each small angular group.
        radial = sorted(
            customers,
            key=lambda i: (bearings[i], -radii[i]),
            reverse=reverse_angle,
        )
        routes.append([0] + radial + [0])

    # Deduplicate deterministic candidates.
    unique = []
    seen = set()
    for route in routes:
        key = tuple(route)
        if key not in seen and len(route) == customer_count + 2:
            seen.add(key)
            unique.append(route)

    return unique


def improve_driver_sweep_route(route, distances, durations, fuel_price, mpg, locations):
    """Make only small road-aware changes without breaking territory order."""
    if not route or len(route) < 5:
        return route[:]

    best = route[:]
    best_score = route_score(best, distances, durations, fuel_price, mpg, locations)

    # Adjacent swaps only.  V25.3 adds one small protection: do not swap
    # jobs that belong to different dynamically detected geographic zones.
    # This keeps a road-efficient improvement from accidentally tearing apart
    # an area that the sweep has already grouped together.  The zones are
    # calculated from the actual coordinates, so nothing is hard-coded to
    # Grantham or to the current test postcodes.
    zone_labels = geographic_zone_labels(location_cache_key(locations))

    for _ in range(3):
        changed = False
        for i in range(1, len(best) - 2):
            left = best[i]
            right = best[i + 1]
            if left <= 0 or right <= 0:
                continue
            if zone_labels[left - 1] != zone_labels[right - 1]:
                continue

            candidate = best[:]
            candidate[i], candidate[i + 1] = candidate[i + 1], candidate[i]
            score = route_score(candidate, distances, durations, fuel_price, mpg, locations)
            if score < best_score - 0.01:
                best = candidate
                best_score = score
                changed = True
        if not changed:
            break
    return best


def optimise_route(
    distances,
    durations,
    fuel_price,
    mpg,
    locations=None,
):
    """General-purpose driver-style route optimiser.

    V23 makes the geographical sweep the primary structure.  It is not tuned
    to the current 30-address example: all territories are generated from the
    actual customer coordinates for whatever jobs are supplied.
    """
    customer_count = len(distances) - 1
    if customer_count <= 0:
        return [0, 0]

    structured_results = []
    if locations is not None:
        structured_candidates = build_driver_sweep_routes(
            locations, distances, durations
        )

        # Keep the older geographic candidates as additional general-purpose
        # options, but they are evaluated alongside the new sweep rather than
        # replacing it with a pure road-distance answer.
        structured_candidates.extend(
            directional_sector_routes(locations, distances, durations)
        )
        structured_candidates.extend(
            geographic_zone_routes(locations)
        )

        seen = set()
        for candidate in structured_candidates:
            key = tuple(candidate)
            if key in seen or len(candidate) != customer_count + 2:
                continue
            seen.add(key)

            improved = improve_driver_sweep_route(
                candidate,
                distances,
                durations,
                fuel_price,
                mpg,
                locations,
            )
            metrics = route_metrics(
                improved, distances, durations, fuel_price, mpg
            )
            score = route_score(
                improved, distances, durations, fuel_price, mpg, locations
            )
            structured_results.append(
                (score, metrics["time_s"], metrics["distance_m"], improved)
            )

    structured_results.sort(key=lambda x: x[0])

    # Road-efficient benchmark.  This is retained so the app can reject a
    # genuinely absurd sweep caused by an unusual road network.
    fallback_candidates = []
    starts = list(range(1, customer_count + 1))
    starts.sort(key=lambda x: durations[0][x])

    if customer_count > 40:
        selected = starts[:12] + starts[-12:] + starts[::max(1, customer_count // 12)]
        starts = list(dict.fromkeys(selected))

    for first_customer in starts:
        fallback_candidates.append(
            build_greedy_route(first_customer, distances, durations, "time")
        )
        fallback_candidates.append(
            build_greedy_route(first_customer, distances, durations, "balanced")
        )
        fallback_candidates.append(
            build_greedy_route(first_customer, distances, durations, "distance")
        )

    insertion = cheapest_insertion_route(distances, durations)
    if insertion:
        fallback_candidates.append(insertion)
        if len(insertion) > 3:
            fallback_candidates.append([0] + insertion[1:-1][::-1] + [0])

    unique_fallbacks = []
    seen = set()
    for candidate in fallback_candidates:
        key = tuple(candidate)
        if key not in seen:
            seen.add(key)
            unique_fallbacks.append(candidate)

    fallback_results = []
    for candidate in unique_fallbacks[:80]:
        improved = improve_route(
            candidate,
            distances,
            durations,
            fuel_price,
            mpg,
            locations,
            preserve_structure=False,
        )
        metrics = route_metrics(
            improved, distances, durations, fuel_price, mpg
        )
        score = route_score(
            improved, distances, durations, fuel_price, mpg, locations
        )
        fallback_results.append(
            (score, metrics["time_s"], metrics["distance_m"], improved)
        )

    fallback_results.sort(key=lambda x: x[0])

    if not structured_results:
        return fallback_results[0][3] if fallback_results else None
    if not fallback_results:
        return structured_results[0][3]

    best_structured = structured_results[0]
    best_fallback = fallback_results[0]

    structured_time = best_structured[1]
    structured_distance = best_structured[2]
    fallback_time = best_fallback[1]
    fallback_distance = best_fallback[2]

    time_ratio = structured_time / max(fallback_time, 1.0)
    distance_ratio = structured_distance / max(fallback_distance, 1.0)

    # V25 gives the geographical sweep a little more authority.  A clean
    # driver-style territory route is allowed to cost a modest amount more
    # than the pure road-time benchmark because repeatedly returning to an
    # area that has already been cleared is expensive in real working time.
    if time_ratio <= 1.20 and distance_ratio <= 1.20:
        return best_structured[3]

    combined_ratio = time_ratio * 0.60 + distance_ratio * 0.40
    if combined_ratio <= 1.16:
        return best_structured[3]

    return best_fallback[3]


# ============================================================
# DESTINATION / WHATSAPP
# ============================================================

def get_destination(row):
    address = clean_val(row.get("address_text"))
    if address:
        return address

    geo_query = clean_val(row.get("geo_query"))
    if geo_query:
        return geo_query

    postcode = normalise_postcode(row.get("Postcode"))
    return postcode


def whatsapp_url(phone, price):
    message = (
        f"Hi from {BUSINESS_NAME}! "
        f"Your service is complete today. "
        f"Total: £{price:.2f}. "
        f"Please pay via bank transfer to {BANK_NAME} - "
        f"Sort Code: {SORT_CODE} "
        f"Account: {ACCOUNT_NUMBER}. "
        f"Thank you!"
    )

    return (
        "https://wa.me/"
        f"{quote(normalise_phone(phone))}"
        f"?text={quote(message)}"
    )


# ============================================================
# LOAD EXISTING DAY
# ============================================================

existing_day = load_day(service_date_str)

if (
    not existing_day.empty
    and "master_df" not in st.session_state
):
    st.session_state.master_df = existing_day.copy()

    # Existing persisted routes already have coordinates.
    if "route_data" not in st.session_state:
        customer_rows = existing_day[
            existing_day["Status"].astype(str).str.lower() != "depot"
        ].copy()

        if not customer_rows.empty:
            completed = (
                customer_rows["Status"]
                .astype(str)
                .str.lower()
                .eq("completed")
                .sum()
            )
            st.session_state.route_data = {
                "revenue": float(customer_rows["Price"].sum()),
                "fuel_cost": 0.0,
                "take_home": float(
                    customer_rows["Price"].sum()
                    * (1 - TAX_RATE)
                ),
                "miles": 0.0,
                "litres": 0.0,
                "time": 0.0,
                "offline": False,
                "jobs": len(customer_rows),
                "completed": int(completed),
                "persisted_only": True,
            }


# ============================================================
# START NEW DAY / RESET
# ============================================================

if st.sidebar.button(
    "🔄 Start New Day / Reset",
    use_container_width=True,
):
    delete_day(service_date_str)
    for key in [
        "master_df",
        "route_data",
        "failed_jobs",
    ]:
        st.session_state.pop(key, None)
    st.rerun()


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📁 Upload your day's CSV or Excel file",
    type=["csv", "xlsx"],
)

if uploaded_file is not None:
    if st.session_state.get("uploaded_filename") != uploaded_file.name:
        st.session_state.pop("master_df", None)
        st.session_state.pop("route_data", None)
        st.session_state.pop("failed_jobs", None)

    st.session_state.uploaded_filename = uploaded_file.name

    if "master_df" not in st.session_state:
        try:
            if uploaded_file.name.lower().endswith(".xlsx"):
                df = pd.read_excel(uploaded_file)
            else:
                df = pd.read_csv(uploaded_file)

            df.columns = (
                df.columns.astype(str)
                .str.strip()
            )

            required = ["Postcode", "Price", "Phone"]
            missing = [
                column
                for column in required
                if column not in df.columns
            ]

            if missing:
                st.error(
                    "Missing required columns: "
                    + ", ".join(missing)
                )
                st.stop()

            df = df.dropna(how="all").copy()

            df["Postcode"] = (
                df["Postcode"]
                .fillna("")
                .astype(str)
                .map(normalise_postcode)
            )

            df = df[
                ~df["Postcode"].isin(["", "NAN", "NAT"])
            ].copy()

            df["Price"] = pd.to_numeric(
                df["Price"],
                errors="coerce",
            )

            df = df[
                df["Price"].notna()
                & (df["Price"] > 0)
            ].copy()

            df["Phone"] = (
                df["Phone"]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            df = df[
                ~df["Phone"]
                .str.lower()
                .isin(["", "nan", "nat"])
            ].copy()

            df["Phone"] = df["Phone"].map(normalise_phone)

            # Add stable identifiers and persistence fields.
            job_ids = []
            address_texts = []

            for row_number, (_, row) in enumerate(
                df.iterrows(),
                start=1,
            ):
                job_ids.append(
                    make_job_id(
                        service_date_str,
                        row_number,
                        row["Postcode"],
                        row["Phone"],
                    )
                )
                address_texts.append(
                    get_address_text(row)
                )

            df["job_id"] = job_ids
            df["service_date"] = service_date_str
            df["Status"] = "pending"
            df["Payment"] = "Waiting"
            df["PaymentTime"] = ""
            df["CompletedTime"] = ""
            df["route_order"] = None
            df["address_text"] = address_texts
            df["latitude"] = None
            df["longitude"] = None
            df["geo_query"] = ""
            df["created_at"] = now_text()

            # If this day already exists, do not overwrite completed
            # or payment information.
            existing = load_day(service_date_str)

            if not existing.empty:
                existing_small = existing[
                    [
                        "job_id",
                        "Status",
                        "Payment",
                        "PaymentTime",
                        "CompletedTime",
                        "route_order",
                        "latitude",
                        "longitude",
                        "geo_query",
                        "address_text",
                    ]
                ].copy()

                df = df.drop(
                    columns=[
                        "Status",
                        "Payment",
                        "PaymentTime",
                        "CompletedTime",
                        "route_order",
                        "latitude",
                        "longitude",
                        "geo_query",
                        "address_text",
                    ],
                    errors="ignore",
                ).merge(
                    existing_small,
                    on="job_id",
                    how="left",
                )

                df["Status"] = df["Status"].fillna("pending")
                df["Payment"] = df["Payment"].fillna("Waiting")
                df["PaymentTime"] = df["PaymentTime"].fillna("")
                df["CompletedTime"] = df["CompletedTime"].fillna("")
                df["address_text"] = df["address_text"].fillna("")
                df["geo_query"] = df["geo_query"].fillna("")

            save_dataframe(df)
            st.session_state.master_df = df.reset_index(drop=True)
            st.session_state.pop("route_data", None)
            st.session_state.pop("failed_jobs", None)
            st.rerun()

        except Exception as exc:
            st.error(f"Error loading file: {exc}")


# ============================================================
# WAIT FOR DATA
# ============================================================

if "master_df" not in st.session_state:
    st.info(
        "Upload your day's file to begin. "
        "If you have already planned this date, the saved day "
        "will load automatically."
    )
    st.stop()


df = st.session_state.master_df


# ============================================================
# PLAN ROUTE
# ============================================================

if st.button(
    "🚀 PLAN / RE-PLAN BEST DAILY ROUTE",
    type="primary",
    use_container_width=True,
):
    if df.empty:
        st.error("No valid customer jobs found.")
        st.stop()

    with st.spinner("📍 Locating your depot..."):
        depot_coords = get_coords(
            DEPOT_FULL_ADDRESS,
            DEPOT_POSTCODE,
        )

    if depot_coords is None:
        st.error("Could not locate the depot.")
        st.stop()

    # Plan only jobs that have not already been completed.
    work_df = df[
        df["Status"].astype(str).str.lower() != "completed"
    ].copy()

    if work_df.empty:
        st.success("🎉 All jobs for this day are already completed.")
        st.stop()

    progress = st.progress(
        0,
        text="Locating customer addresses...",
    )

    valid_rows = []
    failed_rows = []

    total = len(work_df)

    for number, (idx, row) in enumerate(
        work_df.iterrows(),
        start=1,
    ):
        query = build_geo_query(
            row,
            DEPOT_POSTCODE,
        )

        coords = get_coords(
            query,
            row["Postcode"],
        )

        if coords is None:
            failed_rows.append(idx)
        else:
            valid_rows.append(idx)

            # Store geocoding results in the master dataframe.
            df.at[idx, "latitude"] = coords[0]
            df.at[idx, "longitude"] = coords[1]
            df.at[idx, "geo_query"] = query
            df.at[idx, "address_text"] = get_address_text(row)

        progress.progress(
            number / total,
            text=f"Locating customer {number}/{total}",
        )

        # Respect Nominatim's public-service usage.
        # Cached addresses do not incur the delay.
        if cache_key_for(query, row["Postcode"]) not in st.session_state.geocode_cache:
            time.sleep(1)

    progress.empty()

    if failed_rows:
        st.session_state.failed_jobs = (
            df.loc[failed_rows].copy()
        )
    else:
        st.session_state.failed_jobs = pd.DataFrame()

    if not valid_rows:
        st.error("No customer addresses could be located.")
        st.stop()

    rows = [
        {
            "Postcode": DEPOT_POSTCODE,
            "Price": 0.0,
            "Phone": "",
            "Status": "depot",
            "Payment": "Waiting",
            "geo_query": DEPOT_FULL_ADDRESS,
            "latitude": depot_coords[0],
            "longitude": depot_coords[1],
        }
    ]

    locations = [[depot_coords[1], depot_coords[0]]]

    for idx in valid_rows:
        row = df.loc[idx].to_dict()
        rows.append(row)
        locations.append(
            [
                float(row["longitude"]),
                float(row["latitude"]),
            ]
        )

    routing_df = pd.DataFrame(rows)

    with st.spinner(
        "🛣️ Getting actual road distances and driving times..."
    ):
        distances, durations = get_ors_matrix(locations)

    using_offline = False

    if distances is None or durations is None:
        using_offline = True
        st.warning(
            "OpenRouteService could not provide the live road matrix. "
            "Using an offline estimate instead."
        )
        distances, durations = offline_matrix(locations)

    with st.spinner(
        f"🧠 Optimising {len(valid_rows)} customer stops..."
    ):
        route = optimise_route(
            distances,
            durations,
            FUEL_PRICE,
            MPG,
            locations,
        )

    if not route:
        st.error("The route optimiser could not create a route.")
        st.stop()

    metrics = route_metrics(
        route,
        distances,
        durations,
        FUEL_PRICE,
        MPG,
    )

    revenue = float(
        routing_df["Price"].sum()
    )

    fuel_cost = metrics["fuel_cost"]
    pre_tax_profit = revenue - fuel_cost
    take_home = pre_tax_profit * (1 - TAX_RATE)

    # Save route order back against stable job IDs.
    final_rows = routing_df.iloc[route].reset_index(drop=True)

    customer_order = 0

    for _, route_row in final_rows.iterrows():
        if str(route_row.get("Status", "")).lower() == "depot":
            continue

        job_id = route_row.get("job_id")
        if job_id in set(df["job_id"].astype(str)):
            master_idx = df.index[
                df["job_id"].astype(str) == str(job_id)
            ][0]

            df.at[master_idx, "route_order"] = customer_order
            customer_order += 1

    # Persist every changed row.
    save_dataframe(df)

    st.session_state.master_df = df

    st.session_state.route_data = {
        "revenue": revenue,
        "fuel_cost": fuel_cost,
        "take_home": take_home,
        "miles": metrics["miles"],
        "litres": metrics["litres"],
        "time": metrics["time_s"],
        "offline": using_offline,
        "jobs": len(valid_rows),
        "completed": int(
            df["Status"]
            .astype(str)
            .str.lower()
            .eq("completed")
            .sum()
        ),
        "persisted_only": False,
    }

    st.rerun()


# ============================================================
# DASHBOARD
# ============================================================

route_data = st.session_state.get("route_data")

if route_data:
    st.markdown("---")
    st.subheader("💰 Daily Route Summary")

    customer_df = df.copy()

    total_jobs = len(customer_df)
    completed_jobs = int(
        customer_df["Status"]
        .astype(str)
        .str.lower()
        .eq("completed")
        .sum()
    )

    paid_jobs = int(
        customer_df["Payment"]
        .astype(str)
        .str.lower()
        .isin(["cash", "bank transfer", "card", "paid"])
        .sum()
    )

    unpaid_jobs = int(
        customer_df["Payment"]
        .astype(str)
        .str.lower()
        .isin(["not paid", "waiting"])
        .sum()
    )

    c1, c2 = st.columns(2)
    with c1:
        st.metric(
            "Take-Home",
            f"£{route_data['take_home']:.2f}",
        )
    with c2:
        st.metric(
            "Revenue",
            f"£{route_data['revenue']:.2f}",
        )

    c3, c4 = st.columns(2)
    with c3:
        st.metric(
            "Driving Distance",
            f"{route_data['miles']:.1f} miles",
        )
    with c4:
        st.metric(
            "Driving Time",
            format_duration(route_data["time"]),
        )

    c5, c6 = st.columns(2)
    with c5:
        st.metric(
            "Fuel Cost",
            f"£{route_data['fuel_cost']:.2f}",
        )
    with c6:
        st.metric(
            "Fuel Used",
            f"{route_data['litres']:.1f} litres",
        )

    c7, c8, c9 = st.columns(3)
    with c7:
        st.metric("Jobs", total_jobs)
    with c8:
        st.metric("Completed", completed_jobs)
    with c9:
        st.metric("Paid", paid_jobs)

    if unpaid_jobs:
        st.warning(
            f"{unpaid_jobs} job(s) are still showing as Waiting / Not Paid."
        )

    if route_data["offline"]:
        st.warning(
            "This route used estimated distances because "
            "the live road-routing matrix was unavailable."
        )
    elif route_data.get("persisted_only"):
        st.info(
            "This day's jobs were restored from saved records. "
            "Press PLAN / RE-PLAN to refresh live route mileage and time."
        )
    else:
        st.success(
            "✅ Route calculated using live road distance and driving time."
        )


# ============================================================
# FAILED ADDRESSES
# ============================================================

failed_jobs = st.session_state.get(
    "failed_jobs",
    pd.DataFrame(),
)

if not failed_jobs.empty:
    st.markdown("---")
    st.error(
        f"{len(failed_jobs)} customer(s) could not be located. "
        "They were NOT included in the route."
    )

    with st.expander("View unlocated customers"):
        st.dataframe(
            failed_jobs,
            use_container_width=True,
        )


# ============================================================
# SIDEBAR NEXT STOP
# ============================================================

customers = df[
    df["Status"].astype(str).str.lower() != "depot"
].copy()

pending = customers[
    customers["Status"].astype(str).str.lower() != "completed"
].copy()

if "route_order" in pending.columns:
    pending["_sort"] = pd.to_numeric(
        pending["route_order"],
        errors="coerce",
    )
    pending = pending.sort_values(
        ["_sort"],
        na_position="last",
    )

if not pending.empty:
    next_row = pending.iloc[0]
    destination = get_destination(next_row)

    st.sidebar.markdown("---")
    st.sidebar.subheader("🧭 Route Navigation")

    st.sidebar.link_button(
        "🚗 Navigate to Next Stop",
        maps_url(destination),
        use_container_width=True,
    )

    st.sidebar.caption(
        f"Next: {destination}"
    )

    st.sidebar.caption(
        f"{len(pending)} stops remaining"
    )
else:
    st.sidebar.markdown("---")
    st.sidebar.success(
        "🎉 All customer stops completed!"
    )


# ============================================================
# ROUTE LIST
# ============================================================

st.markdown("---")
st.subheader("📍 Planned Route")

# Completed jobs are kept in records but displayed separately.
route_df = df[
    df["Status"].astype(str).str.lower() != "completed"
].copy()

if "route_order" in route_df.columns:
    route_df["_route_sort"] = pd.to_numeric(
        route_df["route_order"],
        errors="coerce",
    )
    route_df = route_df.sort_values(
        "_route_sort",
        na_position="last",
    )

with st.container(border=True):
    st.write("### 🏠 START — GRANTHAM DEPOT")
    st.write(DEPOT_FULL_ADDRESS)

if route_df.empty:
    st.success("No pending stops remain.")
else:
    for display_number, (idx, row) in enumerate(
        route_df.iterrows(),
        start=1,
    ):
        postcode = clean_val(row.get("Postcode"))
        price = float(row.get("Price", 0))
        phone = clean_val(row.get("Phone"))
        status = clean_val(row.get("Status")).lower()
        payment = clean_val(row.get("Payment")) or "Waiting"
        destination = get_destination(row)

        if status == "completed":
            icon = "✅"
            text = "Completed"
        else:
            icon = "⏳"
            text = "Pending"

        extra = []

        ignored = {
            "postcode",
            "price",
            "phone",
            "status",
            "payment",
            "paymenttime",
            "completedtime",
            "job_id",
            "service_date",
            "route_order",
            "address_text",
            "latitude",
            "longitude",
            "geo_query",
            "created_at",
            "_route_sort",
        }

        for column in row.index:
            if str(column).lower() in ignored:
                continue

            value = clean_val(row.get(column))
            if value:
                extra.append(
                    f"**{column}:** {value}"
                )

        with st.container(border=True):
            st.write(
                f"### {icon} STOP {display_number} — {postcode}"
            )

            if extra:
                st.write(" | ".join(extra))

            st.write(
                f"**Price:** £{price:.2f} "
                f"| **Status:** {text} "
                f"| **Payment:** {payment}"
            )

            col1, col2 = st.columns(2)

            with col1:
                if status != "completed":
                    if st.button(
                        "✅ Mark Complete",
                        key=f"complete_{row['job_id']}",
                        use_container_width=True,
                    ):
                        master_idx = df.index[
                            df["job_id"].astype(str)
                            == str(row["job_id"])
                        ][0]

                        df.at[
                            master_idx,
                            "Status",
                        ] = "completed"

                        df.at[
                            master_idx,
                            "CompletedTime",
                        ] = now_text()

                        save_job(df.loc[master_idx])
                        st.session_state.master_df = df
                        st.rerun()
                else:
                    st.success("Completed")

            with col2:
                st.link_button(
                    "🚗 Navigate Here",
                    maps_url(destination),
                    key=f"nav_{row['job_id']}",
                    use_container_width=True,
                )

            pay1, pay2, pay3 = st.columns(3)

            with pay1:
                if st.button(
                    "💵 Cash",
                    key=f"cash_{row['job_id']}",
                    use_container_width=True,
                ):
                    master_idx = df.index[
                        df["job_id"].astype(str)
                        == str(row["job_id"])
                    ][0]

                    df.at[
                        master_idx,
                        "Payment",
                    ] = "Cash"

                    df.at[
                        master_idx,
                        "PaymentTime",
                    ] = now_text()

                    save_job(df.loc[master_idx])
                    st.session_state.master_df = df
                    st.rerun()

            with pay2:
                if st.button(
                    "🏦 Bank Transfer",
                    key=f"bank_{row['job_id']}",
                    use_container_width=True,
                ):
                    master_idx = df.index[
                        df["job_id"].astype(str)
                        == str(row["job_id"])
                    ][0]

                    df.at[
                        master_idx,
                        "Payment",
                    ] = "Bank Transfer"

                    df.at[
                        master_idx,
                        "PaymentTime",
                    ] = now_text()

                    save_job(df.loc[master_idx])
                    st.session_state.master_df = df
                    st.rerun()

            with pay3:
                if st.button(
                    "❌ Not Paid",
                    key=f"notpaid_{row['job_id']}",
                    use_container_width=True,
                ):
                    master_idx = df.index[
                        df["job_id"].astype(str)
                        == str(row["job_id"])
                    ][0]

                    df.at[
                        master_idx,
                        "Payment",
                    ] = "Not Paid"

                    df.at[
                        master_idx,
                        "PaymentTime",
                    ] = now_text()

                    save_job(df.loc[master_idx])
                    st.session_state.master_df = df
                    st.rerun()

            if status == "completed" and phone:
                st.link_button(
                    "💬 Send WhatsApp Payment Message",
                    whatsapp_url(phone, price),
                    use_container_width=True,
                )

with st.container(border=True):
    st.write("### 🏁 FINISH — GRANTHAM DEPOT")
    st.write(DEPOT_FULL_ADDRESS)

    st.link_button(
        "🚗 Navigate Back to Depot",
        maps_url(DEPOT_FULL_ADDRESS),
        use_container_width=True,
    )


# ============================================================
# COMPLETED / PAYMENT SUMMARY
# ============================================================

completed_df = df[
    df["Status"].astype(str).str.lower() == "completed"
].copy()

if not completed_df.empty:
    st.markdown("---")
    st.subheader("✅ Completed Jobs")

    show_completed = completed_df[
        [
            column
            for column in [
                "Postcode",
                "Price",
                "Phone",
                "Payment",
                "PaymentTime",
                "CompletedTime",
            ]
            if column in completed_df.columns
        ]
    ].copy()

    st.dataframe(
        show_completed,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

st.sidebar.markdown("---")
st.sidebar.subheader("📊 Export Records")

export_df = df.copy()

if not export_df.empty:
    export_df = export_df[
        export_df["Status"].astype(str).str.lower() != "depot"
    ].copy()

    export_columns_to_remove = [
        "latitude",
        "longitude",
        "geo_query",
        "created_at",
        "_route_sort",
    ]

    export_df = export_df.drop(
        columns=[
            c for c in export_columns_to_remove
            if c in export_df.columns
        ],
        errors="ignore",
    )

    # Put useful columns first.
    preferred = [
        "Postcode",
        "Price",
        "Phone",
        "Status",
        "Payment",
        "PaymentTime",
        "CompletedTime",
        "route_order",
        "address_text",
        "job_id",
    ]

    ordered = [
        c for c in preferred if c in export_df.columns
    ]

    remaining = [
        c for c in export_df.columns if c not in ordered
    ]

    export_df = export_df[ordered + remaining]

    output = io.BytesIO()
    workbook = Workbook()

    summary_ws = workbook.active
    summary_ws.title = "Daily Summary"

    total_revenue = float(export_df["Price"].sum()) if not export_df.empty else 0
    completed_count = int(
        export_df["Status"]
        .astype(str)
        .str.lower()
        .eq("completed")
        .sum()
    ) if not export_df.empty else 0

    cash_total = float(
        export_df.loc[
            export_df["Payment"].astype(str).str.lower() == "cash",
            "Price",
        ].sum()
    ) if not export_df.empty else 0

    bank_total = float(
        export_df.loc[
            export_df["Payment"].astype(str).str.lower() == "bank transfer",
            "Price",
        ].sum()
    ) if not export_df.empty else 0

    unpaid_total = float(
        export_df.loc[
            export_df["Payment"].astype(str).str.lower().isin(
                ["waiting", "not paid"]
            ),
            "Price",
        ].sum()
    ) if not export_df.empty else 0

    summary_rows = [
        ["DanCleanUK Daily Report", ""],
        ["Route Date", service_date_str],
        ["Depot", DEPOT_FULL_ADDRESS],
        ["Total Jobs", len(export_df)],
        ["Completed Jobs", completed_count],
        ["Revenue", total_revenue],
        ["Cash Received", cash_total],
        ["Bank Transfer Received", bank_total],
        ["Outstanding / Unpaid", unpaid_total],
        ["Fuel Cost", route_data["fuel_cost"] if route_data else 0],
        ["Fuel Used (litres)", route_data["litres"] if route_data else 0],
        ["Driving Miles", route_data["miles"] if route_data else 0],
        ["Driving Time", format_duration(route_data["time"]) if route_data else "0m"],
        ["Tax Rate", f"{TAX_RATE * 100:.0f}%"],
        ["Estimated Take-Home", route_data["take_home"] if route_data else 0],
    ]

    for row in summary_rows:
        summary_ws.append(row)

    header_fill = PatternFill(
        start_color="2F4F4F",
        end_color="2F4F4F",
        fill_type="solid",
    )

    header_font = Font(
        color="FFFFFF",
        bold=True,
    )

    summary_ws["A1"].fill = header_fill
    summary_ws["A1"].font = header_font

    summary_ws.column_dimensions["A"].width = 28
    summary_ws.column_dimensions["B"].width = 45

    records_ws = workbook.create_sheet("Customer Records")

    for row in dataframe_to_rows(
        export_df,
        index=False,
        header=True,
    ):
        records_ws.append(row)

    for cell in records_ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center"
        )

    for column_cells in records_ws.columns:
        max_length = 0
        letter = column_cells[0].column_letter

        for cell in column_cells:
            try:
                max_length = max(
                    max_length,
                    len(str(cell.value)),
                )
            except Exception:
                pass

        records_ws.column_dimensions[letter].width = min(
            max_length + 2,
            50,
        )

    workbook.save(output)

    st.sidebar.download_button(
        "⬇️ Download Excel Report",
        output.getvalue(),
        f"DanCleanUK_{service_date_str}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.sidebar.caption(
    f"DanCleanUK Route Optimizer v{APP_VERSION}"
)
