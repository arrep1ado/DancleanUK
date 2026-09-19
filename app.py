import io
import math
import re
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
# Version 25.49
# ============================================================

APP_VERSION = "25.51"
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
    # Geocoding is part of route correctness. Never carry coordinates from an
    # older app version into a new geocoding/route engine.
    st.session_state.clear()
    st.session_state.app_version = APP_VERSION
    st.session_state.geocode_cache = {}

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
st.sidebar.subheader("🧠 Route Optimisation")

# V25.49 uses one consistent whole-day economic objective.
# There are deliberately no separate time/distance/nearby sliders.
# Those competing filters could override the route search itself.
DRIVING_TIME_VALUE_PER_HOUR = 6.0

st.sidebar.caption(
    "The optimiser compares complete routes using actual road fuel cost "
    "plus a modest value for driving time. It does not impose a driving-time "
    "limit or force nearby customers to be consecutive."
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



def ors_exact_geocode(query, postcode, expected_house_number="", expected_street=""):
    """Find an address-level coordinate without silently using the postcode centroid.

    ORS/Pelias supports both structured and unstructured forward geocoding.
    We try both forms and accept a result only when the returned label/address
    is consistent with the requested house number and street.
    """
    expected_house_number = clean_val(expected_house_number)
    expected_street = clean_val(expected_street).lower()
    postcode_clean = normalise_postcode(postcode)

    headers = {
        "Authorization": API_KEY,
        "Accept": "application/json",
    }

    def valid_feature(feature):
        geometry = feature.get("geometry") or {}
        coords = geometry.get("coordinates") or []
        if len(coords) < 2:
            return None

        try:
            lon = float(coords[0])
            lat = float(coords[1])
        except Exception:
            return None

        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            return None

        props = feature.get("properties") or {}
        label = clean_val(
            props.get("label")
            or props.get("name")
            or feature.get("label")
        ).lower()

        returned_house = clean_val(
            props.get("housenumber")
            or props.get("house_number")
            or props.get("number")
        )

        returned_street = clean_val(
            props.get("street")
            or props.get("streetname")
            or props.get("name")
        ).lower()

        house_match = (
            not expected_house_number
            or returned_house == expected_house_number
            or bool(
                re.search(
                    rf"(?<!\d){re.escape(expected_house_number)}(?!\d)",
                    label,
                )
            )
        )

        street_match = (
            not expected_street
            or expected_street in returned_street
            or expected_street in label
        )

        if house_match and street_match:
            return (lat, lon)

        return None

    endpoints = [
        (
            "https://api.heigit.org/openrouteservice/geocode/search/structured",
            {
                "api_key": API_KEY,
                "address": query,
                "postalcode": postcode_clean,
                "country": "GB",
                "size": 20,
                "layers": "address",
            },
        ),
        (
            "https://api.heigit.org/openrouteservice/geocode/search",
            {
                "api_key": API_KEY,
                "text": query,
                "size": 20,
                "layers": "address",
            },
        ),
    ]

    for endpoint, params in endpoints:
        try:
            response = requests.get(
                endpoint,
                params=params,
                headers=headers,
                timeout=20,
            )

            if response.status_code != 200:
                continue

            data = response.json() or {}

            for feature in data.get("features") or []:
                coords = valid_feature(feature)
                if coords is not None:
                    return coords

        except Exception:
            continue

    return None


def geocode_candidates(query, postcode):
    """Build address-first geocoding queries from today's imported record."""
    query = str(query or "").strip()
    postcode = normalise_postcode(postcode)

    candidates = []

    def add(value):
        value = str(value or "").strip()
        if value and value not in candidates:
            candidates.append(value)

    add(query)

    if postcode:
        street_part = query.replace(postcode, "").strip(" ,")
        street_part = street_part.replace(
            postcode.replace(" ", ""), ""
        ).strip(" ,")
    else:
        street_part = query

    if street_part:
        add(f"{street_part}, United Kingdom")

    if postcode:
        add(f"{postcode}, United Kingdom")

    return candidates


def nominatim_search(
    query,
    headers,
    expected_house_number=None,
    expected_street=None,
    postcode=None,
):
    """Address-aware Nominatim lookup with structured + free-text searches."""
    expected_house_number = clean_val(expected_house_number)
    expected_street = clean_val(expected_street).lower()
    postcode = normalise_postcode(postcode or "")

    searches = [
        {
            "q": query,
            "format": "jsonv2",
            "addressdetails": 1,
            "limit": 20,
            "countrycodes": "gb",
        },
    ]

    if expected_street and postcode:
        searches.append(
            {
                "street": (
                    f"{expected_house_number} {expected_street}"
                    if expected_house_number
                    else expected_street
                ),
                "postalcode": postcode,
                "country": "United Kingdom",
                "format": "jsonv2",
                "addressdetails": 1,
                "limit": 20,
            }
        )

    for params in searches:
        try:
            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params=params,
                headers=headers,
                timeout=20,
            )

            if response.status_code != 200:
                continue

            for item in response.json() or []:
                address = item.get("address") or {}
                display = clean_val(item.get("display_name")).lower()

                returned_house = clean_val(
                    address.get("house_number")
                )
                returned_road = clean_val(
                    address.get("road")
                    or address.get("pedestrian")
                    or address.get("residential")
                ).lower()

                house_match = (
                    not expected_house_number
                    or returned_house == expected_house_number
                    or bool(
                        re.search(
                            rf"(?<!\d){re.escape(expected_house_number)}(?!\d)",
                            display,
                        )
                    )
                )

                street_match = (
                    not expected_street
                    or expected_street in returned_road
                    or expected_street in display
                )

                if not (house_match and street_match):
                    continue

                try:
                    return (
                        float(item["lat"]),
                        float(item["lon"]),
                    )
                except Exception:
                    continue

        except Exception:
            continue

    return None


# Safe coordinates for two older Grantham postcodes that can appear in
# genuine customer records but may no longer be returned by postcodes.io.
# These are only a last-resort fallback after exact address geocoding fails.
# They keep the job in the correct Grantham area rather than dropping it.
LEGACY_POSTCODE_COORDS = {
    "NG31 7AN": (52.909806, -0.640572),
    "NG31 9EH": (52.909052, -0.630469),
}


def get_postcode_coords(postcode):
    """Return the official postcode centroid used as a safety anchor.

    The postcode is not used as the normal house-level location. It is only
    used to validate an exact geocoder result and as a safe fallback when the
    house cannot be resolved. This prevents a bad geocoder match in a distant
    part of the UK from creating a multi-thousand-mile route.
    """
    postcode = normalise_postcode(postcode)
    if not postcode:
        return None

    cache_key = f"__POSTCODE__|{postcode}".lower()
    cached = st.session_state.geocode_cache.get(cache_key)
    if cached is not None:
        return cached

    for pc in dict.fromkeys([postcode, postcode.replace(" ", "")]):
        try:
            response = requests.get(
                f"https://api.postcodes.io/postcodes/{quote(pc)}",
                timeout=10,
            )
            if response.status_code != 200:
                continue

            result = response.json().get("result") or {}
            lat = result.get("latitude")
            lon = result.get("longitude")
            if lat is None or lon is None:
                continue

            coords = (float(lat), float(lon))
            st.session_state.geocode_cache[cache_key] = coords
            return coords
        except Exception:
            continue

    return None


def get_coords(query_string, postcode):
    """Locate an address safely, while guaranteeing valid postcode jobs stay in the route.

    Exact house/street geocoding is preferred. Every exact result is checked
    against the official postcode centroid before it can enter the route.
    If exact geocoding fails, the postcode centroid is used instead of
    dropping the customer. This is deliberately conservative: a slightly
    imprecise point inside the correct postcode is far safer than a wrong
    address hundreds of miles away.
    """
    query = str(query_string or "").strip()
    postcode = normalise_postcode(postcode)
    key = cache_key_for(query, postcode)

    cached = st.session_state.geocode_cache.get(key)
    if cached is not None:
        return cached

    headers = {
        "User-Agent": "DanCleanUKRouteOptimizer/25.49"
    }

    # Extract a likely house number and street from the imported address.
    house_match = re.search(r"(?<!\d)(\d+[A-Za-z]?)\b", query)
    expected_house = house_match.group(1) if house_match else ""

    expected_street = ""
    if house_match:
        tail = query[house_match.end():]
        expected_street = tail.split(",")[0].strip()

    # Get the postcode anchor BEFORE accepting an exact result. This is the
    # critical protection against V25.44's 2,500-mile failure.
    postcode_anchor = get_postcode_coords(postcode)

    def safe_exact(coords):
        if coords is None:
            return None
        try:
            lat, lon = float(coords[0]), float(coords[1])
        except Exception:
            return None

        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            return None

        if postcode_anchor is not None:
            # A UK postcode normally covers a small local area.  Five km is a
            # deliberately generous safety boundary for rural postcodes while
            # still rejecting obviously wrong geocoder hits.
            if haversine_km(
                postcode_anchor[0],
                postcode_anchor[1],
                lat,
                lon,
            ) > 5.0:
                return None

        return (lat, lon)

    # 1. Exact ORS/Pelias lookup using the V25.43 endpoint. Do NOT use the
    # V25.44 api.heigit.org/pelias endpoint which produced the catastrophic
    # false locations in the live test.
    exact_ors = ors_exact_geocode(
        query,
        postcode,
        expected_house_number=expected_house,
        expected_street=expected_street,
    )
    exact_ors = safe_exact(exact_ors)
    if exact_ors is not None:
        st.session_state.geocode_cache[key] = exact_ors
        return exact_ors

    # 2. Address-aware Nominatim lookup, also protected by the postcode anchor.
    for candidate in geocode_candidates(query, postcode):
        coords = nominatim_search(
            candidate,
            headers,
            expected_house_number=expected_house,
            expected_street=expected_street,
            postcode=postcode,
        )
        coords = safe_exact(coords)
        if coords is not None:
            st.session_state.geocode_cache[key] = coords
            return coords

        time.sleep(0.35)

    # 3. Safe fallback: keep the customer in the route at the correct postcode
    # area. This is much better than silently removing a legitimate job.
    if postcode_anchor is not None:
        st.session_state.geocode_cache[key] = postcode_anchor
        return postcode_anchor

    # 4. Historical Grantham postcode fallback. These coordinates are known
    # local anchors and are used only when the live postcode service has no
    # record. This prevents legitimate jobs such as NG31 7AN / NG31 9EH from
    # becoming "unlocated".
    legacy = LEGACY_POSTCODE_COORDS.get(postcode)
    if legacy is not None:
        st.session_state.geocode_cache[key] = legacy
        return legacy

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
            "https://api.heigit.org/openrouteservice/v2/matrix/driving-car",
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
    """Single economic objective for the complete route.

    V25.49 deliberately removes the old time/distance/cluster/shape scoring
    system. The optimiser is no longer asked to satisfy several competing
    filters. It simply compares the complete driving day using: 

      * actual road fuel cost
      * a modest value assigned to driving time

    Geographic proximity is therefore allowed to help naturally through the
    real road mileage/time, but it cannot veto a route that is economically
    better overall.
    """
    metrics = route_metrics(
        route,
        distances,
        durations,
        fuel_price,
        mpg,
    )

    driving_hours = metrics["time_s"] / 3600.0
    return metrics["fuel_cost"] + (driving_hours * DRIVING_TIME_VALUE_PER_HOUR)














def build_greedy_route(
    first_customer,
    distances,
    durations,
    mode="economic",
    fuel_price=None,
    mpg=None,
):
    """Build a neutral route using the same economic edge objective.

    V25.50 removes the old hidden nearby-job/local-density rewards.  Greedy
    construction is now only a starting point; the final route is judged by
    the same whole-route fuel + driving-time objective as every other route.
    """
    if fuel_price is None:
        fuel_price = FUEL_PRICE
    if mpg is None:
        mpg = MPG

    customer_count = len(distances) - 1
    route = [0, first_customer]
    remaining = list(range(1, customer_count + 1))
    remaining.remove(first_customer)
    current = first_customer

    def edge_cost(a, b):
        miles = float(distances[a][b]) / 1000.0 * 0.621371
        litres = miles / float(mpg) * 4.54609
        fuel_cost = litres * float(fuel_price)
        time_cost = (float(durations[a][b]) / 3600.0) * DRIVING_TIME_VALUE_PER_HOUR
        return fuel_cost + time_cost

    while remaining:
        best = None

        for candidate in remaining:
            cost = edge_cost(current, candidate)

            # Small look-ahead only uses the same economic edge cost.
            # It does not reward postcode proximity, density, zones or shape.
            future = [x for x in remaining if x != candidate]
            if future:
                continuation = min(
                    edge_cost(candidate, x) for x in future
                )
                score = cost + continuation * 0.35
            else:
                score = cost + edge_cost(candidate, 0)

            item = (score, cost, candidate)
            if best is None or item < best:
                best = item

        next_customer = best[2]
        route.append(next_customer)
        remaining.remove(next_customer)
        current = next_customer

    route.append(0)
    return route





def cheapest_insertion_route(
    distances,
    durations,
):
    customer_count = len(distances) - 1

    if customer_count <= 0:
        return [0, 0]

    # Use several deterministic economic starting seeds. The seed is only a
    # construction choice; the complete-route objective decides the winner.
    def edge_economic_cost(a, b):
        miles = float(distances[a][b]) / 1000.0 * 0.621371
        litres = miles / float(MPG) * 4.54609
        fuel_cost = litres * float(FUEL_PRICE)
        time_cost = (float(durations[a][b]) / 3600.0) * DRIVING_TIME_VALUE_PER_HOUR
        return fuel_cost + time_cost

    ordered_seeds = sorted(
        range(1, customer_count + 1),
        key=lambda x: (edge_economic_cost(0, x), x),
    )
    seed_count = min(8, customer_count)
    seeds = ordered_seeds[:seed_count]
    if customer_count > seed_count:
        seeds.extend(ordered_seeds[-min(4, customer_count - seed_count):])
    seeds = list(dict.fromkeys(seeds))

    best_route = None
    best_value = float("inf")

    for first in seeds:
        route = [0, first, 0]
        remaining = list(range(1, customer_count + 1))
        remaining.remove(first)

        while remaining:
            best_choice = None

            for customer in sorted(remaining):
                for position in range(1, len(route)):
                    before = route[position - 1]
                    after = route[position]

                    old_time = durations[before][after]
                    new_time = durations[before][customer] + durations[customer][after]
                    old_distance = distances[before][after]
                    new_distance = distances[before][customer] + distances[customer][after]

                    # Score the insertion only by the same economic edge
                    # objective used everywhere else. No nearby-job bonus,
                    # locality reward or postcode grouping is applied.
                    old_miles = float(old_distance) / 1000.0 * 0.621371
                    new_miles = float(new_distance) / 1000.0 * 0.621371
                    old_fuel = (old_miles / float(MPG) * 4.54609) * float(FUEL_PRICE)
                    new_fuel = (new_miles / float(MPG) * 4.54609) * float(FUEL_PRICE)
                    increase = (new_fuel - old_fuel) + (
                        (new_time - old_time) / 3600.0
                    ) * DRIVING_TIME_VALUE_PER_HOUR

                    if best_choice is None or (increase, customer, position) < best_choice:
                        best_choice = (increase, customer, position)

            _, customer, position = best_choice
            route.insert(position, customer)
            remaining.remove(customer)

        value = sum(durations[route[i]][route[i + 1]] for i in range(len(route) - 1))
        if value < best_value:
            best_value = value
            best_route = route

    return best_route


















def build_nearest_pocket_route(
    distances,
    durations,
    start_customer,
):
    """Build a road-aware candidate without locality gates.

    V25.49 keeps this as one candidate generator only. It does not force a
    nearby customer, impose an escape distance, or penalise leaving a pocket.
    The complete-route objective decides whether this candidate is useful.
    """
    customer_count = len(distances) - 1
    if customer_count <= 0:
        return [0, 0]

    route = [0, start_customer]
    remaining = list(range(1, customer_count + 1))
    remaining.remove(start_customer)
    current = start_customer

    while remaining:
        def candidate_key(j):
            direct_time = float(durations[current][j])
            direct_distance = float(distances[current][j])
            future = [x for x in remaining if x != j]

            if future:
                next_job = min(
                    future,
                    key=lambda x: (
                        float(durations[j][x]),
                        float(distances[j][x]),
                        x,
                    ),
                )
                lookahead_time = float(durations[j][next_job])
                lookahead_distance = float(distances[j][next_job])
            else:
                lookahead_time = float(durations[j][0])
                lookahead_distance = float(distances[j][0])

            # Small look-ahead uses the same fuel + time economics.
            direct_miles = direct_distance / 1000.0 * 0.621371
            lookahead_miles = lookahead_distance / 1000.0 * 0.621371
            direct_fuel = (direct_miles / float(MPG) * 4.54609) * float(FUEL_PRICE)
            lookahead_fuel = (lookahead_miles / float(MPG) * 4.54609) * float(FUEL_PRICE)
            direct_cost = direct_fuel + (direct_time / 3600.0) * DRIVING_TIME_VALUE_PER_HOUR
            lookahead_cost = lookahead_fuel + (lookahead_time / 3600.0) * DRIVING_TIME_VALUE_PER_HOUR
            return (
                direct_cost + 0.20 * lookahead_cost,
                direct_distance + 0.20 * lookahead_distance,
                j,
            )

        chosen = min(remaining, key=candidate_key)
        route.append(chosen)
        remaining.remove(chosen)
        current = chosen

    route.append(0)
    return route





def _practical_route_key(route, distances, durations, fuel_price=None, mpg=None):
    """Return the one objective used everywhere in V25.49.

    No locality, zone, shape, backtracking or time-limit penalty is included.
    If a fuel price/MPG pair is not supplied, use the app's current settings.
    """
    if fuel_price is None:
        fuel_price = FUEL_PRICE
    if mpg is None:
        mpg = MPG

    metrics = route_metrics(
        route,
        distances,
        durations,
        fuel_price,
        mpg,
    )
    time_cost = (metrics["time_s"] / 3600.0) * DRIVING_TIME_VALUE_PER_HOUR
    economic_score = metrics["fuel_cost"] + time_cost

    return (
        economic_score,
        metrics["fuel_cost"],
        metrics["time_s"],
        metrics["distance_m"],
        tuple(route),
    )

def _route_search(route, distances, durations, max_rounds=2, fuel_price=None, mpg=None):
    """Deep deterministic whole-route improvement.

    Unlike the old neighbouring-stop polish, this searches the complete route
    for 2-opt reversals and Or-opt relocations.  A move may therefore move a
    customer across a large part of the day when that genuinely improves the
    complete road journey.  Depot remains fixed at both ends.
    """
    if not route or len(route) < 5:
        return route[:]

    best = route[:]
    best_key = _practical_route_key(best, distances, durations, fuel_price, mpg)
    n = len(best)

    for _ in range(max_rounds):
        changed = False

        # 2-opt: reverse every possible internal section.  This is the main
        # anti-zigzag move because it changes the order of an entire pocket.
        for i in range(1, n - 2):
            for j in range(i + 1, n - 1):
                candidate = best[:]
                candidate[i:j + 1] = reversed(candidate[i:j + 1])
                key = _practical_route_key(candidate, distances, durations, fuel_price, mpg)
                if key < best_key:
                    best = candidate
                    best_key = key
                    changed = True

        # Or-opt: move one, two, or three consecutive jobs to another place.
        # This is particularly useful when a sweep has one address stranded
        # on the wrong side of a territory.
        for block_size in (1, 2, 3):
            if block_size >= n - 2:
                continue
            i = 1
            while i + block_size < n - 1:
                block = best[i:i + block_size]
                remainder = best[:i] + best[i + block_size:]
                moved = False

                for j in range(1, len(remainder)):
                    if j == i:
                        continue
                    candidate = remainder[:j] + block + remainder[j:]
                    key = _practical_route_key(candidate, distances, durations, fuel_price, mpg)
                    if key < best_key:
                        best = candidate
                        best_key = key
                        n = len(best)
                        changed = True
                        moved = True
                        break

                if moved:
                    # Restart the scan from the beginning after every accepted
                    # relocation so improvements are never missed.
                    i = 1
                    continue
                i += 1

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
    """V25.51 complete-day route optimiser.

    Every candidate is built from the same economic road-cost logic and every
    final decision is made by the complete-route objective: fuel cost plus the
    value assigned to driving time. There are no locality, zone, shape,
    backtracking, postcode or driving-time filters.
    """
    customer_count = len(distances) - 1
    if customer_count <= 0:
        return [0, 0]

    candidates = []

    def add_candidate(candidate):
        if not candidate:
            return
        if len(candidate) != customer_count + 2:
            return
        if candidate[0] != 0 or candidate[-1] != 0:
            return
        if set(candidate[1:-1]) != set(range(1, customer_count + 1)):
            return
        candidates.append(candidate[:])

    def depot_economic_cost(customer):
        miles = float(distances[0][customer]) / 1000.0 * 0.621371
        litres = miles / float(mpg) * 4.54609
        return (litres * float(fuel_price)) + (
            float(durations[0][customer]) / 3600.0
        ) * DRIVING_TIME_VALUE_PER_HOUR

    starts = list(range(1, customer_count + 1))
    starts.sort(key=lambda x: (depot_economic_cost(x), x))

    # For normal daily lists (such as 33 jobs) every starting customer is
    # tested. Larger lists use a deterministic spread to keep Streamlit fast.
    if customer_count > 60:
        picks = starts[:20] + starts[-20:]
        step = max(1, customer_count // 20)
        picks.extend(starts[::step])
        starts = list(dict.fromkeys(picks))

    for first_customer in starts:
        add_candidate(
            build_greedy_route(
                first_customer,
                distances,
                durations,
                "economic",
                fuel_price=fuel_price,
                mpg=mpg,
            )
        )

    insertion = cheapest_insertion_route(distances, durations)
    if insertion:
        add_candidate(insertion)
        if len(insertion) > 3:
            add_candidate([0] + insertion[1:-1][::-1] + [0])

    # Unconstrained driver-style candidates.
    for first_customer in starts[:min(24, len(starts))]:
        add_candidate(
            build_nearest_pocket_route(
                distances,
                durations,
                first_customer,
            )
        )

    if not candidates:
        return None

    unique = []
    seen = set()
    for route in candidates:
        key = tuple(route)
        if key not in seen:
            seen.add(key)
            unique.append(route)

    # Keep enough diversity to prevent one construction family from dominating,
    # while avoiding an excessive deep-search runtime.
    unique.sort(
        key=lambda r: _practical_route_key(
            r, distances, durations, fuel_price, mpg
        )
    )
    construction_pool = unique[:16]

    if len(unique) > 16:
        tail_step = max(1, len(unique) // 6)
        for route in unique[::tail_step][:6]:
            if tuple(route) not in {tuple(x) for x in construction_pool}:
                construction_pool.append(route)

    refined = []
    seen_refined = set()
    for route in construction_pool:
        improved = _route_search(
            route,
            distances,
            durations,
            max_rounds=2,
            fuel_price=fuel_price,
            mpg=mpg,
        )
        key = tuple(improved)
        if key not in seen_refined:
            seen_refined.add(key)
            refined.append(improved)

    if not refined:
        refined = construction_pool

    refined.sort(
        key=lambda r: _practical_route_key(
            r, distances, durations, fuel_price, mpg
        )
    )
    return refined[0]


def surgical_route_polish(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    max_passes=2,
):
    """Final whole-route polish using the same V25.50 economic objective.

    A change does not have to improve both time and distance. If it costs a
    little more time but saves enough fuel/mileage to reduce the total route
    cost, it is allowed. This keeps the final polish consistent with the main
    optimiser instead of reintroducing an old hidden filter.
    """
    if not route or len(route) < 5:
        return route[:] if route else route

    best = route[:]
    best_key = _practical_route_key(
        best, distances, durations, fuel_price, mpg
    )

    for _ in range(max_passes):
        found = False

        for i in range(1, len(best) - 1):
            customer = best[i]
            shortened = best[:i] + best[i + 1:]

            for j in range(1, len(shortened)):
                candidate = shortened[:j] + [customer] + shortened[j:]
                candidate_key = _practical_route_key(
                    candidate, distances, durations, fuel_price, mpg
                )

                if candidate_key < best_key:
                    best = candidate
                    best_key = candidate_key
                    found = True
                    break

            if found:
                break

        if found:
            continue

        # Also test complete 2-opt reversals during the final polish.
        for i in range(1, len(best) - 2):
            for j in range(i + 1, len(best) - 1):
                candidate = best[:]
                candidate[i:j + 1] = reversed(candidate[i:j + 1])
                candidate_key = _practical_route_key(
                    candidate, distances, durations, fuel_price, mpg
                )
                if candidate_key < best_key:
                    best = candidate
                    best_key = candidate_key
                    found = True
                    break
            if found:
                break

        if not found:
            break

    return best


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
        # Failed geocodes must never retain an old route position.
        # Keep the customer in the database, but remove it from the route
        # until its address can be located successfully.
        for failed_idx in failed_rows:
            df.at[failed_idx, "route_order"] = None
            df.at[failed_idx, "latitude"] = None
            df.at[failed_idx, "longitude"] = None
            df.at[failed_idx, "geo_query"] = ""

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

    # Detect multiple different customer records receiving the same coordinate.
    # This is a data-quality warning, not a hard-coded route rule.
    coordinate_groups = {}
    for idx in range(1, len(locations)):
        coord_key = (
            round(float(locations[idx][0]), 6),
            round(float(locations[idx][1]), 6),
        )
        coordinate_groups.setdefault(coord_key, []).append(idx)

    duplicate_groups = [
        group for group in coordinate_groups.values()
        if len(group) > 1
    ]

    if duplicate_groups:
        duplicate_labels = []
        for group in duplicate_groups:
            labels = []
            for idx in group:
                labels.append(
                    str(routing_df.iloc[idx].get("address_text", "customer"))
                )
            duplicate_labels.append(" / ".join(labels))

        st.warning(
            "⚠️ Multiple customer addresses resolved to the same map "
            "coordinate. The route will use that coordinate, but exact "
            "house-level routing could not be confirmed for: "
            + "; ".join(duplicate_labels)
        )

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

    # v25.37: final whole-route polish after complete-day optimisation has been
    # selected.  It can only replace the route when the COMPLETE route is
    # strictly better in both live road time and live road distance.
    route = surgical_route_polish(
        route,
        distances,
        durations,
        FUEL_PRICE,
        MPG,
    )

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
        "unlocated": len(failed_rows),
        "completed": int(
            df["Status"]
            .astype(str)
            .str.lower()
            .eq("completed")
            .sum()
        ),
        "persisted_only": False,
    }

    # Render the Daily Route Summary immediately in this run.  Do not rerun
    # after planning: the route data above is already in session state and the
    # dashboard below can display it without risking the summary being hidden.


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
        st.metric("Jobs Routed", route_data.get("jobs", 0))
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
# Only located customers with a current route position belong in the
# Planned Route. Unlocated customers stay in the database and appear only
# in the unlocated-customer section.
route_df = df[
    (df["Status"].astype(str).str.lower() != "completed")
    & df["route_order"].notna()
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
