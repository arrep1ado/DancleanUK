import io
import math
import random
import sqlite3
import time
from datetime import datetime, date
from urllib.parse import quote

import pandas as pd
import requests
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# ============================================================
# APP CONFIG
# ============================================================

APP_VERSION = "12.1"
DB_FILE = "dancleanuk.db"

st.set_page_config(
    page_title="DanCleanUK Daily Route Optimizer",
    page_icon="🚐",
    layout="wide",
)


# ============================================================
# DATABASE
# ============================================================

def db_connect():
    return sqlite3.connect(DB_FILE)


def init_db():
    conn = db_connect()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS jobs (
            job_id TEXT PRIMARY KEY,
            service_date TEXT,
            postcode TEXT,
            price REAL,
            phone TEXT,
            status TEXT,
            payment TEXT,
            payment_time TEXT,
            completed_time TEXT,
            route_order INTEGER,
            address_text TEXT,
            latitude REAL,
            longitude REAL,
            geo_query TEXT,
            created_at TEXT
        )
    """)

    conn.commit()
    conn.close()


init_db()


# ============================================================
# SESSION STATE
# ============================================================

if "app_version" not in st.session_state:
    st.session_state.app_version = APP_VERSION

if st.session_state.app_version != APP_VERSION:
    for key in list(st.session_state.keys()):
        if key != "geocode_cache":
            del st.session_state[key]

    st.session_state.app_version = APP_VERSION


if "geocode_cache" not in st.session_state:
    st.session_state.geocode_cache = {}


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_val(value):
    if pd.isna(value):
        return ""

    return str(value).strip()


def clean_optional(value):
    value = clean_val(value)

    if value.lower() in {
        "",
        "nan",
        "none",
        "null",
    }:
        return ""

    return value


def safe_float(value, default=0.0):
    try:
        if pd.isna(value):
            return default

        return float(value)

    except Exception:
        return default


def normalise_postcode(value):
    value = clean_val(value)

    return " ".join(
        value.upper().split()
    )


def postcode_without_space(value):
    return normalise_postcode(
        value
    ).replace(" ", "")


def normalise_phone(value):
    value = clean_val(value)

    return value


def maps_url(address):
    return (
        "https://www.google.com/maps/search/?api=1&query="
        + quote(address)
    )


def format_duration(seconds):
    seconds = max(0, int(seconds))

    hours = seconds // 3600
    minutes = (seconds % 3600) // 60

    if hours:
        return f"{hours}h {minutes:02d}m"

    return f"{minutes}m"


def make_job_id(postcode, address="", service_date=""):
    raw = (
        f"{service_date}|"
        f"{postcode}|"
        f"{address}"
    )

    import hashlib

    return hashlib.sha1(
        raw.encode("utf-8")
    ).hexdigest()[:16]


def now_text():
    return datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


# ============================================================
# DATABASE FUNCTIONS
# ============================================================

def save_job(row):
    conn = db_connect()
    cur = conn.cursor()

    cur.execute("""
        INSERT OR REPLACE INTO jobs (
            job_id,
            service_date,
            postcode,
            price,
            phone,
            status,
            payment,
            payment_time,
            completed_time,
            route_order,
            address_text,
            latitude,
            longitude,
            geo_query,
            created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row.get("job_id"),
        row.get("service_date"),
        row.get("postcode"),
        safe_float(row.get("price")),
        row.get("phone", ""),
        row.get("status", "Pending"),
        row.get("payment", "Unpaid"),
        row.get("payment_time", ""),
        row.get("completed_time", ""),
        row.get("route_order"),
        row.get("address_text", ""),
        row.get("latitude"),
        row.get("longitude"),
        row.get("geo_query", ""),
        row.get("created_at", now_text()),
    ))

    conn.commit()
    conn.close()


def save_dataframe(df):
    if df is None or df.empty:
        return

    for _, row in df.iterrows():
        save_job(row.to_dict())


def load_day(service_date):
    conn = db_connect()

    df = pd.read_sql_query(
        """
        SELECT *
        FROM jobs
        WHERE service_date = ?
        ORDER BY
            CASE
                WHEN route_order IS NULL THEN 999999
                ELSE route_order
            END,
            postcode
        """,
        conn,
        params=(service_date,),
    )

    conn.close()

    return df


def delete_day(service_date):
    conn = db_connect()
    cur = conn.cursor()

    cur.execute(
        """
        DELETE FROM jobs
        WHERE service_date = ?
        """,
        (service_date,),
    )

    conn.commit()
    conn.close()


# ============================================================
# SIDEBAR SETTINGS
# ============================================================

st.sidebar.title("⚙️ Route Settings")

route_date = st.sidebar.date_input(
    "Route date",
    value=date.today(),
)

route_date_text = route_date.isoformat()

depot_postcode = st.sidebar.text_input(
    "Depot postcode",
    value="NG31 9RA",
)

depot_address = st.sidebar.text_input(
    "Depot address",
    value="192 Queensway, Grantham NG31 9RA",
)

fuel_price = st.sidebar.number_input(
    "Fuel price £/litre",
    min_value=0.01,
    value=1.50,
    step=0.01,
)

mpg = st.sidebar.number_input(
    "Vehicle MPG",
    min_value=1.0,
    value=30.0,
    step=1.0,
)

tax_rate = st.sidebar.number_input(
    "Tax rate %",
    min_value=0.0,
    max_value=100.0,
    value=20.0,
    step=1.0,
)

st.sidebar.markdown("---")

st.sidebar.caption(
    f"Route Optimizer version {APP_VERSION}"
)


# ============================================================
# API KEY
# ============================================================

try:
    API_KEY = st.secrets["API_KEY"]
except Exception:
    API_KEY = ""


# ============================================================
# GEOCODING
# ============================================================

def build_geo_query(row):
    parts = []

    possible_columns = [
        "address",
        "Address",
        "street",
        "Street",
        "address_text",
        "Address Text",
        "house",
        "House",
        "house_number",
        "House Number",
        "town",
        "Town",
        "city",
        "City",
        "location",
        "Location",
    ]

    for column in possible_columns:

        if column in row.index:

            value = clean_optional(
                row[column]
            )

            if value and value not in parts:
                parts.append(value)

    postcode = normalise_postcode(
        row.get("postcode", "")
    )

    if postcode:
        parts.append(postcode)

    parts.append("UK")

    return ", ".join(parts)


def get_address_text(row):
    parts = []

    possible_columns = [
        "address",
        "Address",
        "street",
        "Street",
        "address_text",
        "Address Text",
        "house",
        "House",
        "house_number",
        "House Number",
        "town",
        "Town",
        "city",
        "City",
    ]

    for column in possible_columns:

        if column in row.index:

            value = clean_optional(
                row[column]
            )

            if value and value not in parts:
                parts.append(value)

    postcode = normalise_postcode(
        row.get("postcode", "")
    )

    if postcode:
        parts.append(postcode)

    return ", ".join(parts)


def cache_key_for(query):
    return "geo:" + query.lower().strip()


def get_coords(row):
    """
    ROBUST GEOCODING

    The route only needs reliable geographical coordinates.

    Priority:

    1. Postcodes.io exact postcode
    2. Postcodes.io postcode without spaces
    3. Nominatim postcode
    4. Nominatim full address

    This means a bad/odd street address in the spreadsheet
    will NOT prevent a valid postcode from being routed.

    The full address is still retained separately for
    Google Maps navigation.
    """

    postcode = normalise_postcode(
        row.get("postcode", "")
    )

    query = build_geo_query(row)

    # --------------------------------------------------------
    # Cache by postcode first
    # --------------------------------------------------------

    if postcode:

        postcode_cache_key = (
            "postcode:"
            + postcode_without_space(
                postcode
            )
        )

        if (
            postcode_cache_key
            in st.session_state.geocode_cache
        ):

            return st.session_state.geocode_cache[
                postcode_cache_key
            ]

    # Full query cache
    cache_key = cache_key_for(query)

    if cache_key in st.session_state.geocode_cache:

        return st.session_state.geocode_cache[
            cache_key
        ]

    headers = {
        "User-Agent":
            "DanCleanUK-Daily-Route-Optimizer/12.1"
    }

    # ========================================================
    # 1. POSTCODES.IO - NORMAL POSTCODE
    # ========================================================

    if postcode:

        try:

            clean_postcode = (
                postcode_without_space(
                    postcode
                )
            )

            response = requests.get(
                "https://api.postcodes.io/postcodes/"
                + quote(clean_postcode),
                timeout=15,
            )

            if response.status_code == 200:

                data = response.json()

                result_data = data.get(
                    "result"
                )

                if result_data:

                    lat = float(
                        result_data[
                            "latitude"
                        ]
                    )

                    lon = float(
                        result_data[
                            "longitude"
                        ]
                    )

                    result = {
                        "latitude": lat,
                        "longitude": lon,
                        "geo_query": postcode,
                    }

                    st.session_state.geocode_cache[
                        postcode_cache_key
                    ] = result

                    st.session_state.geocode_cache[
                        cache_key
                    ] = result

                    return result

        except Exception:
            pass

    # ========================================================
    # 2. POSTCODES.IO - FORMATTED POSTCODE
    # ========================================================

    if postcode:

        try:

            response = requests.get(
                "https://api.postcodes.io/postcodes/"
                + quote(postcode),
                timeout=15,
            )

            if response.status_code == 200:

                data = response.json()

                result_data = data.get(
                    "result"
                )

                if result_data:

                    lat = float(
                        result_data[
                            "latitude"
                        ]
                    )

                    lon = float(
                        result_data[
                            "longitude"
                        ]
                    )

                    result = {
                        "latitude": lat,
                        "longitude": lon,
                        "geo_query": postcode,
                    }

                    st.session_state.geocode_cache[
                        postcode_cache_key
                    ] = result

                    st.session_state.geocode_cache[
                        cache_key
                    ] = result

                    return result

        except Exception:
            pass

    # ========================================================
    # 3. NOMINATIM POSTCODE ONLY
    # ========================================================

    if postcode:

        try:

            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": postcode + ", UK",
                    "format": "json",
                    "limit": 1,
                    "countrycodes": "gb",
                },
                headers=headers,
                timeout=20,
            )

            if response.status_code == 200:

                data = response.json()

                if data:

                    lat = float(
                        data[0]["lat"]
                    )

                    lon = float(
                        data[0]["lon"]
                    )

                    result = {
                        "latitude": lat,
                        "longitude": lon,
                        "geo_query": postcode,
                    }

                    st.session_state.geocode_cache[
                        postcode_cache_key
                    ] = result

                    st.session_state.geocode_cache[
                        cache_key
                    ] = result

                    time.sleep(1)

                    return result

        except Exception:
            pass

    # ========================================================
    # 4. NOMINATIM FULL ADDRESS
    # ========================================================

    if query:

        try:

            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": query,
                    "format": "json",
                    "limit": 1,
                    "countrycodes": "gb",
                },
                headers=headers,
                timeout=20,
            )

            if response.status_code == 200:

                data = response.json()

                if data:

                    lat = float(
                        data[0]["lat"]
                    )

                    lon = float(
                        data[0]["lon"]
                    )

                    result = {
                        "latitude": lat,
                        "longitude": lon,
                        "geo_query": query,
                    }

                    if postcode:
                        st.session_state.geocode_cache[
                            postcode_cache_key
                        ] = result

                    st.session_state.geocode_cache[
                        cache_key
                    ] = result

                    time.sleep(1)

                    return result

        except Exception:
            pass

    return None


# ============================================================
# ROUTING
# ============================================================

def haversine_km(
    lat1,
    lon1,
    lat2,
    lon2,
):
    R = 6371.0

    p1 = math.radians(lat1)
    p2 = math.radians(lat2)

    dlat = math.radians(
        lat2 - lat1
    )

    dlon = math.radians(
        lon2 - lon1
    )

    a = (
        math.sin(dlat / 2) ** 2
        +
        math.cos(p1)
        * math.cos(p2)
        * math.sin(dlon / 2) ** 2
    )

    return (
        2
        * R
        * math.asin(
            math.sqrt(a)
        )
    )


def offline_matrix(coords):
    """
    Offline fallback.

    coords:
        [(longitude, latitude), ...]
    """

    n = len(coords)

    distance_matrix = [
        [0.0] * n
        for _ in range(n)
    ]

    duration_matrix = [
        [0.0] * n
        for _ in range(n)
    ]

    road_factor = 1.30
    average_speed_kmh = 35.0

    for i in range(n):

        for j in range(n):

            if i == j:
                continue

            km = haversine_km(
                coords[i][1],
                coords[i][0],
                coords[j][1],
                coords[j][0],
            )

            road_km = (
                km * road_factor
            )

            distance_matrix[i][j] = (
                road_km * 1000
            )

            duration_matrix[i][j] = (
                road_km
                / average_speed_kmh
                * 3600
            )

    return (
        distance_matrix,
        duration_matrix,
    )


def get_ors_matrix(
    coords,
    api_key,
):
    """
    Real road distance/time matrix
    from OpenRouteService.
    """

    url = (
        "https://api.openrouteservice.org/"
        "v2/matrix/driving-car"
    )

    payload = {
        "locations": coords,
        "metrics": [
            "distance",
            "duration",
        ],
        "units": "m",
    }

    headers = {
        "Authorization": api_key,
        "Content-Type": "application/json",
    }

    response = requests.post(
        url,
        json=payload,
        headers=headers,
        timeout=60,
    )

    response.raise_for_status()

    data = response.json()

    distances = data.get(
        "distances"
    )

    durations = data.get(
        "durations"
    )

    if not distances or not durations:
        raise ValueError(
            "OpenRouteService returned no matrix data."
        )

    return (
        distances,
        durations,
    )


# ============================================================
# ROUTE METRICS
# ============================================================

def route_metrics(
    route,
    distance_matrix,
    duration_matrix,
    mpg,
    fuel_price,
):
    total_distance_m = 0.0
    total_time_s = 0.0

    for i in range(
        len(route) - 1
    ):

        a = route[i]
        b = route[i + 1]

        distance = (
            distance_matrix[a][b]
        )

        duration = (
            duration_matrix[a][b]
        )

        if distance is None:
            distance = float("inf")

        if duration is None:
            duration = float("inf")

        total_distance_m += distance
        total_time_s += duration

    total_distance_miles = (
        total_distance_m / 1609.344
    )

    total_time_hours = (
        total_time_s / 3600
    )

    litres_used = (
        total_distance_miles
        / max(mpg, 1)
        * 4.54609
    )

    fuel_cost = (
        litres_used * fuel_price
    )

    return {
        "distance_m": total_distance_m,
        "distance_miles":
            total_distance_miles,
        "time_s": total_time_s,
        "time_hours":
            total_time_hours,
        "litres": litres_used,
        "fuel_cost": fuel_cost,
    }


def route_time(
    route,
    duration_matrix,
):
    total = 0.0

    for i in range(
        len(route) - 1
    ):

        total += duration_matrix[
            route[i]
        ][
            route[i + 1]
        ]

    return total


def route_distance(
    route,
    distance_matrix,
):
    total = 0.0

    for i in range(
        len(route) - 1
    ):

        total += distance_matrix[
            route[i]
        ][
            route[i + 1]
        ]

    return total


def geographical_cluster_penalty(
    route,
    distance_matrix,
):
    """
    Mild penalty for very long jumps between otherwise
    geographically close jobs.

    Driving time remains the main priority.
    """

    if len(route) <= 4:
        return 0.0

    penalty = 0.0

    leg_lengths = []

    for i in range(
        len(route) - 1
    ):

        leg_lengths.append(
            distance_matrix[
                route[i]
            ][
                route[i + 1]
            ]
        )

    if not leg_lengths:
        return 0.0

    average_leg = (
        sum(leg_lengths)
        / len(leg_lengths)
    )

    if average_leg <= 0:
        return 0.0

    for leg in leg_lengths:

        if leg > (
            average_leg * 2.0
        ):

            penalty += (
                leg
                - average_leg * 2.0
            )

    return penalty


def route_score(
    route,
    distance_matrix,
    duration_matrix,
):
    """
    TIME-FIRST route score.

    1. Driving time
    2. Distance
    3. Very small geographical penalty

    The uploaded file order is NOT used.
    """

    time_s = route_time(
        route,
        duration_matrix,
    )

    distance_m = route_distance(
        route,
        distance_matrix,
    )

    cluster_penalty = (
        geographical_cluster_penalty(
            route,
            distance_matrix,
        )
    )

    score = (
        time_s
        + distance_m * 0.03
        + cluster_penalty * 0.01
    )

    return score


# ============================================================
# INITIAL ROUTES
# ============================================================

def nearest_neighbour_route(
    duration_matrix,
    distance_matrix,
    start_mode="time",
):
    """
    Creates a route from the depot.

    Customer order in the uploaded file is ignored.
    """

    n = len(
        duration_matrix
    )

    customers = set(
        range(1, n)
    )

    route = [0]

    current = 0

    while customers:

        best_customer = None
        best_value = float("inf")

        for candidate in customers:

            direct_time = (
                duration_matrix[
                    current
                ][
                    candidate
                ]
            )

            direct_distance = (
                distance_matrix[
                    current
                ][
                    candidate
                ]
            )

            if start_mode == "distance":

                value = direct_distance

            elif start_mode == "combined":

                value = (
                    direct_time
                    + direct_distance * 0.025
                )

            else:

                value = direct_time

            if value < best_value:

                best_value = value
                best_customer = candidate

        route.append(
            best_customer
        )

        customers.remove(
            best_customer
        )

        current = best_customer

    route.append(0)

    return route


def lookahead_route(
    duration_matrix,
    distance_matrix,
):
    """
    Looks at both the next job and the job after that.
    """

    n = len(
        duration_matrix
    )

    remaining = set(
        range(1, n)
    )

    route = [0]

    current = 0

    while remaining:

        best_job = None
        best_score = float("inf")

        for candidate in remaining:

            first_time = (
                duration_matrix[
                    current
                ][
                    candidate
                ]
            )

            first_distance = (
                distance_matrix[
                    current
                ][
                    candidate
                ]
            )

            others = (
                remaining
                - {candidate}
            )

            if others:

                next_best_time = min(
                    duration_matrix[
                        candidate
                    ][
                        x
                    ]
                    for x in others
                )

                next_best_distance = min(
                    distance_matrix[
                        candidate
                    ][
                        x
                    ]
                    for x in others
                )

            else:

                next_best_time = (
                    duration_matrix[
                        candidate
                    ][0]
                )

                next_best_distance = (
                    distance_matrix[
                        candidate
                    ][0]
                )

            value = (
                first_time
                + next_best_time * 0.45
                + first_distance * 0.02
                + next_best_distance * 0.01
            )

            if value < best_score:

                best_score = value
                best_job = candidate

        route.append(
            best_job
        )

        remaining.remove(
            best_job
        )

        current = best_job

    route.append(0)

    return route


def cheapest_insertion_route(
    duration_matrix,
    distance_matrix,
):
    """
    Builds a route by inserting each job where it creates
    the smallest increase in total driving time.
    """

    n = len(
        duration_matrix
    )

    customers = list(
        range(1, n)
    )

    if not customers:
        return [0, 0]

    first = max(
        customers,
        key=lambda x:
        (
            duration_matrix[0][x]
            +
            duration_matrix[x][0]
        ),
    )

    route = [
        0,
        first,
        0,
    ]

    customers.remove(
        first
    )

    while customers:

        best_job = None
        best_position = None
        best_increase = float(
            "inf"
        )

        for job in customers:

            for position in range(
                1,
                len(route),
            ):

                before = route[
                    position - 1
                ]

                after = route[
                    position
                ]

                old_cost = (
                    duration_matrix[
                        before
                    ][
                        after
                    ]
                    +
                    distance_matrix[
                        before
                    ][
                        after
                    ] * 0.03
                )

                new_cost = (
                    duration_matrix[
                        before
                    ][
                        job
                    ]
                    +
                    duration_matrix[
                        job
                    ][
                        after
                    ]
                    +
                    (
                        distance_matrix[
                            before
                        ][
                            job
                        ]
                        +
                        distance_matrix[
                            job
                        ][
                            after
                        ]
                    ) * 0.03
                )

                increase = (
                    new_cost
                    - old_cost
                )

                if increase < best_increase:

                    best_increase = increase
                    best_job = job
                    best_position = position

        route.insert(
            best_position,
            best_job,
        )

        customers.remove(
            best_job
        )

    return route


def sweep_route(
    duration_matrix,
    distance_matrix,
    coords,
):
    """
    Geographic sweep around the depot.
    """

    depot_lat = coords[0][1]
    depot_lon = coords[0][0]

    jobs = []

    for index in range(
        1,
        len(coords),
    ):

        lon = coords[index][0]
        lat = coords[index][1]

        angle = math.atan2(
            lat - depot_lat,
            lon - depot_lon,
        )

        radius = haversine_km(
            depot_lat,
            depot_lon,
            lat,
            lon,
        )

        jobs.append(
            (
                angle,
                radius,
                index,
            )
        )

    forward = sorted(
        jobs,
        key=lambda x: (
            x[0],
            x[1],
        ),
    )

    reverse = list(
        reversed(forward)
    )

    candidates = []

    for job_list in [
        forward,
        reverse,
    ]:

        route = [0]

        for _, _, index in job_list:
            route.append(index)

        route.append(0)

        candidates.append(
            route
        )

    return min(
        candidates,
        key=lambda r:
        route_score(
            r,
            distance_matrix,
            duration_matrix,
        ),
    )


# ============================================================
# LOCAL SEARCH
# ============================================================

def two_opt(
    route,
    distance_matrix,
    duration_matrix,
):
    """
    2-opt improvement.
    """

    if len(route) <= 4:
        return route

    improved = True

    while improved:

        improved = False

        current_score = route_score(
            route,
            distance_matrix,
            duration_matrix,
        )

        n = len(route)

        for i in range(
            1,
            n - 3,
        ):

            for j in range(
                i + 1,
                n - 1,
            ):

                if j - i <= 1:
                    continue

                candidate = (
                    route[:i]
                    +
                    list(
                        reversed(
                            route[
                                i:j + 1
                            ]
                        )
                    )
                    +
                    route[
                        j + 1:
                    ]
                )

                candidate_score = (
                    route_score(
                        candidate,
                        distance_matrix,
                        duration_matrix,
                    )
                )

                if candidate_score < (
                    current_score - 0.001
                ):

                    route = candidate
                    improved = True

                    break

            if improved:
                break

    return route


def relocate_improvement(
    route,
    distance_matrix,
    duration_matrix,
):
    """
    Moves one customer to another position.
    """

    if len(route) <= 4:
        return route

    improved = True

    while improved:

        improved = False

        current_score = route_score(
            route,
            distance_matrix,
            duration_matrix,
        )

        n = len(route)

        for i in range(
            1,
            n - 1,
        ):

            customer = route[i]

            shortened = (
                route[:i]
                +
                route[i + 1:]
            )

            for j in range(
                1,
                len(shortened),
            ):

                candidate = (
                    shortened[:j]
                    +
                    [customer]
                    +
                    shortened[j:]
                )

                candidate_score = (
                    route_score(
                        candidate,
                        distance_matrix,
                        duration_matrix,
                    )
                )

                if candidate_score < (
                    current_score - 0.001
                ):

                    route = candidate
                    improved = True

                    break

            if improved:
                break

    return route


def swap_improvement(
    route,
    distance_matrix,
    duration_matrix,
):
    """
    Swaps two customers.
    """

    if len(route) <= 4:
        return route

    improved = True

    while improved:

        improved = False

        current_score = route_score(
            route,
            distance_matrix,
            duration_matrix,
        )

        n = len(route)

        for i in range(
            1,
            n - 2,
        ):

            for j in range(
                i + 1,
                n - 1,
            ):

                candidate = route[:]

                candidate[i], candidate[j] = (
                    candidate[j],
                    candidate[i],
                )

                candidate_score = (
                    route_score(
                        candidate,
                        distance_matrix,
                        duration_matrix,
                    )
                )

                if candidate_score < (
                    current_score - 0.001
                ):

                    route = candidate
                    improved = True

                    break

            if improved:
                break

    return route


# ============================================================
# MAIN ROUTE OPTIMISER
# ============================================================

def optimise_route(
    distance_matrix,
    duration_matrix,
    coords=None,
):
    """
    MAIN OPTIMISER.

    Uploaded file order is completely ignored.

    Every customer is treated as an unordered job.

    Driving time is the main priority.
    """

    number_of_locations = len(
        distance_matrix
    )

    if number_of_locations <= 1:
        return [0]

    if number_of_locations == 2:
        return [0, 1, 0]

    candidates = []

    candidates.append(
        nearest_neighbour_route(
            duration_matrix,
            distance_matrix,
            "time",
        )
    )

    candidates.append(
        nearest_neighbour_route(
            duration_matrix,
            distance_matrix,
            "distance",
        )
    )

    candidates.append(
        nearest_neighbour_route(
            duration_matrix,
            distance_matrix,
            "combined",
        )
    )

    candidates.append(
        lookahead_route(
            duration_matrix,
            distance_matrix,
        )
    )

    candidates.append(
        cheapest_insertion_route(
            duration_matrix,
            distance_matrix,
        )
    )

    if coords is not None:

        candidates.append(
            sweep_route(
                duration_matrix,
                distance_matrix,
                coords,
            )
        )

    unique_candidates = []

    seen = set()

    for route in candidates:

        key = tuple(route)

        if key not in seen:

            seen.add(key)

            unique_candidates.append(
                route
            )

    improved_routes = []

    for route in unique_candidates:

        try:

            route = two_opt(
                route,
                distance_matrix,
                duration_matrix,
            )

            route = relocate_improvement(
                route,
                distance_matrix,
                duration_matrix,
            )

            route = swap_improvement(
                route,
                distance_matrix,
                duration_matrix,
            )

            route = two_opt(
                route,
                distance_matrix,
                duration_matrix,
            )

            improved_routes.append(
                route
            )

        except Exception:

            improved_routes.append(
                route
            )

    best_route = min(
        improved_routes,
        key=lambda r:
        route_score(
            r,
            distance_matrix,
            duration_matrix,
        ),
    )

    return best_route


# ============================================================
# DESTINATION / WHATSAPP
# ============================================================

def destination_for_job(row):
    address = get_address_text(row)

    if address:
        return address

    return normalise_postcode(
        row.get("postcode", "")
    )


def whatsapp_url(phone, message):
    phone = clean_val(phone)

    if not phone:
        return None

    phone = (
        phone.replace(" ", "")
        .replace("-", "")
        .replace("(", "")
        .replace(")", "")
    )

    if phone.startswith("0"):
        phone = "+44" + phone[1:]

    return (
        "https://wa.me/"
        + phone.replace("+", "")
        + "?text="
        + quote(message)
    )


# ============================================================
# LOAD EXISTING DAY
# ============================================================

existing_df = load_day(
    route_date_text
)


# ============================================================
# RESET DAY
# ============================================================

st.sidebar.markdown("---")

if st.sidebar.button(
    "🗑️ Reset this day",
    use_container_width=True,
):

    delete_day(
        route_date_text
    )

    st.session_state.pop(
        "route_data",
        None,
    )

    st.session_state.pop(
        "route_metrics",
        None,
    )

    st.session_state.pop(
        "routing_source",
        None,
    )

    st.rerun()


# ============================================================
# FILE UPLOAD
# ============================================================

st.title(
    "🚐 DanCleanUK Daily Route Optimizer"
)

st.caption(
    "Jobs are treated as an unordered list. "
    "The app calculates the geographical route automatically."
)

uploaded_file = st.file_uploader(
    "Upload today's jobs",
    type=[
        "csv",
        "xlsx",
        "xls",
    ],
)


if uploaded_file is not None:

    try:

        if uploaded_file.name.lower().endswith(
            ".csv"
        ):

            uploaded_df = pd.read_csv(
                uploaded_file
            )

        else:

            uploaded_df = pd.read_excel(
                uploaded_file
            )

        uploaded_df.columns = [
            str(c).strip()
            for c in uploaded_df.columns
        ]

        # ----------------------------------------------------
        # Find postcode column
        # ----------------------------------------------------

        postcode_column = None

        for column in uploaded_df.columns:

            if column.lower() == "postcode":

                postcode_column = column
                break

        if postcode_column is None:

            st.error(
                "The file must contain a Postcode column."
            )

            st.stop()

        if postcode_column != "postcode":

            uploaded_df = uploaded_df.rename(
                columns={
                    postcode_column:
                        "postcode"
                }
            )

        # ----------------------------------------------------
        # Find price column
        # ----------------------------------------------------

        price_column = None

        for column in uploaded_df.columns:

            if column.lower() == "price":

                price_column = column
                break

        if price_column is None:

            st.error(
                "The file must contain a Price column."
            )

            st.stop()

        if price_column != "price":

            uploaded_df = uploaded_df.rename(
                columns={
                    price_column:
                        "price"
                }
            )

        # ----------------------------------------------------
        # Find phone column
        # ----------------------------------------------------

        phone_column = None

        for column in uploaded_df.columns:

            if column.lower() == "phone":

                phone_column = column
                break

        if phone_column is None:

            uploaded_df["phone"] = ""

        elif phone_column != "phone":

            uploaded_df = uploaded_df.rename(
                columns={
                    phone_column:
                        "phone"
                }
            )

        # ----------------------------------------------------
        # Clean data
        # ----------------------------------------------------

        uploaded_df["postcode"] = (
            uploaded_df[
                "postcode"
            ]
            .apply(normalise_postcode)
        )

        uploaded_df["price"] = (
            uploaded_df[
                "price"
            ]
            .apply(
                lambda x:
                safe_float(x)
            )
        )

        uploaded_df["phone"] = (
            uploaded_df[
                "phone"
            ]
            .apply(normalise_phone)
        )

        # ----------------------------------------------------
        # Add stable IDs
        # ----------------------------------------------------

        job_ids = []

        for _, row in uploaded_df.iterrows():

            address = get_address_text(
                row
            )

            job_id = make_job_id(
                row["postcode"],
                address,
                route_date_text,
            )

            job_ids.append(
                job_id
            )

        uploaded_df["job_id"] = job_ids

        uploaded_df[
            "service_date"
        ] = route_date_text

        # ----------------------------------------------------
        # Merge saved information
        # ----------------------------------------------------

        if not existing_df.empty:

            existing_lookup = (
                existing_df
                .set_index("job_id")
                .to_dict("index")
            )

        else:

            existing_lookup = {}

        statuses = []
        payments = []
        payment_times = []
        completed_times = []

        for job_id in job_ids:

            old = existing_lookup.get(
                job_id,
                {},
            )

            statuses.append(
                old.get(
                    "status",
                    "Pending",
                )
            )

            payments.append(
                old.get(
                    "payment",
                    "Unpaid",
                )
            )

            payment_times.append(
                old.get(
                    "payment_time",
                    "",
                )
            )

            completed_times.append(
                old.get(
                    "completed_time",
                    "",
                )
            )

        uploaded_df[
            "status"
        ] = statuses

        uploaded_df[
            "payment"
        ] = payments

        uploaded_df[
            "payment_time"
        ] = payment_times

        uploaded_df[
            "completed_time"
        ] = completed_times

        uploaded_df[
            "address_text"
        ] = uploaded_df.apply(
            get_address_text,
            axis=1,
        )

        uploaded_df[
            "created_at"
        ] = now_text()

        # ----------------------------------------------------
        # Save uploaded jobs
        # ----------------------------------------------------

        save_dataframe(
            uploaded_df
        )

        st.success(
            f"{len(uploaded_df)} jobs loaded."
        )

        st.session_state[
            "uploaded_jobs"
        ] = uploaded_df.copy()

    except Exception as e:

        st.error(
            f"Could not read the file: {e}"
        )


# ============================================================
# LOAD CURRENT DATA
# ============================================================

day_df = load_day(
    route_date_text
)

if day_df.empty:

    st.info(
        "Upload your jobs file to begin."
    )

    st.stop()


# ============================================================
# PLAN / RE-PLAN ROUTE
# ============================================================

st.markdown("---")

if st.button(
    "🧭 PLAN / RE-PLAN ROUTE",
    type="primary",
    use_container_width=True,
):

    progress = st.progress(
        0,
        text="Starting route planning..."
    )

    # --------------------------------------------------------
    # DEPOT
    # --------------------------------------------------------

    depot_row = pd.Series({
        "postcode":
            normalise_postcode(
                depot_postcode
            ),
        "address":
            depot_address,
    })

    progress.progress(
        5,
        text="Finding depot..."
    )

    depot_geo = get_coords(
        depot_row
    )

    if depot_geo is None:

        st.error(
            "Could not find the depot location."
        )

        st.stop()

    # --------------------------------------------------------
    # ACTIVE JOBS
    # --------------------------------------------------------

    work_df = day_df[
        day_df["status"] != "Completed"
    ].copy()

    if work_df.empty:

        st.success(
            "All jobs are already completed."
        )

        st.stop()

    # --------------------------------------------------------
    # GEOCODE JOBS
    # --------------------------------------------------------

    coords = [
        (
            depot_geo["longitude"],
            depot_geo["latitude"],
        )
    ]

    geo_rows = []

    failed_addresses = []

    total_jobs = len(
        work_df
    )

    for position, (_, row) in enumerate(
        work_df.iterrows(),
        start=1,
    ):

        progress_value = int(
            5
            +
            (
                position
                / total_jobs
            )
            * 35
        )

        progress.progress(
            progress_value,
            text=(
                f"Finding address "
                f"{position}/{total_jobs}..."
            ),
        )

        geo = get_coords(
            row
        )

        if geo is None:

            failed_addresses.append(
                row["postcode"]
            )

            continue

        coords.append(
            (
                geo["longitude"],
                geo["latitude"],
            )
        )

        geo_rows.append(
            (
                row,
                geo,
            )
        )

    # --------------------------------------------------------
    # FAILED ADDRESS REPORT
    # --------------------------------------------------------

    if failed_addresses:

        st.error(
            "Could not locate these postcodes: "
            + ", ".join(
                failed_addresses
            )
        )

        st.warning(
            "The app tried Postcodes.io and OpenStreetMap "
            "using the postcode directly. Check these "
            "postcodes in your job file."
        )

        progress.empty()

        st.stop()

    if not geo_rows:

        st.error(
            "No jobs could be geocoded."
        )

        progress.empty()

        st.stop()

    # --------------------------------------------------------
    # BUILD ROUTING DATA
    # --------------------------------------------------------

    routing_rows = []

    for matrix_index, (
        row,
        geo,
    ) in enumerate(
        geo_rows,
        start=1,
    ):

        routing_rows.append({
            "matrix_index":
                matrix_index,
            "job_id":
                row["job_id"],
            "postcode":
                row["postcode"],
            "price":
                safe_float(
                    row["price"]
                ),
            "phone":
                row.get(
                    "phone",
                    "",
                ),
            "status":
                row.get(
                    "status",
                    "Pending",
                ),
            "payment":
                row.get(
                    "payment",
                    "Unpaid",
                ),
            "payment_time":
                row.get(
                    "payment_time",
                    "",
                ),
            "completed_time":
                row.get(
                    "completed_time",
                    "",
                ),
            "address_text":
                row.get(
                    "address_text",
                    "",
                ),
            "latitude":
                geo["latitude"],
            "longitude":
                geo["longitude"],
            "geo_query":
                geo.get(
                    "geo_query",
                    "",
                ),
        })

    routing_df = pd.DataFrame(
        routing_rows
    )

    # --------------------------------------------------------
    # ROAD MATRIX
    # --------------------------------------------------------

    progress.progress(
        45,
        text=(
            "Calculating real road distances "
            "and driving times..."
        ),
    )

    try:

        if not API_KEY:

            raise ValueError(
                "API_KEY is missing."
            )

        distance_matrix, duration_matrix = (
            get_ors_matrix(
                coords,
                API_KEY,
            )
        )

        routing_source = (
            "OpenRouteService live road routing"
        )

    except Exception:

        st.warning(
            "OpenRouteService could not be used. "
            "Using offline geographical routing instead."
        )

        distance_matrix, duration_matrix = (
            offline_matrix(
                coords
            )
        )

        routing_source = (
            "Offline geographical estimate"
        )

    # --------------------------------------------------------
    # OPTIMISE
    # --------------------------------------------------------

    progress.progress(
        60,
        text=(
            "Finding the best geographical route..."
        ),
    )

    optimised_route = optimise_route(
        distance_matrix,
        duration_matrix,
        coords,
    )

    # --------------------------------------------------------
    # METRICS
    # --------------------------------------------------------

    progress.progress(
        85,
        text="Calculating route totals..."
    )

    metrics = route_metrics(
        optimised_route,
        distance_matrix,
        duration_matrix,
        mpg,
        fuel_price,
    )

    # --------------------------------------------------------
    # BUILD FINAL ROUTE DATA
    # --------------------------------------------------------

    final_rows = []

    route_order = 1

    for matrix_index in optimised_route:

        # ----------------------------------------------------
        # DEPOT
        # ----------------------------------------------------

        if matrix_index == 0:

            final_rows.append({
                "route_order":
                    route_order,
                "job_id":
                    "DEPOT",
                "postcode":
                    depot_postcode,
                "address_text":
                    depot_address,
                "price":
                    0.0,
                "phone":
                    "",
                "status":
                    "Depot",
                "payment":
                    "",
                "payment_time":
                    "",
                "completed_time":
                    "",
                "latitude":
                    depot_geo[
                        "latitude"
                    ],
                "longitude":
                    depot_geo[
                        "longitude"
                    ],
                "geo_query":
                    depot_address,
            })

        else:

            row = routing_df[
                routing_df[
                    "matrix_index"
                ]
                == matrix_index
            ]

            if row.empty:
                continue

            row = row.iloc[0]

            final_rows.append({
                "route_order":
                    route_order,
                "job_id":
                    row["job_id"],
                "postcode":
                    row["postcode"],
                "address_text":
                    row["address_text"],
                "price":
                    row["price"],
                "phone":
                    row["phone"],
                "status":
                    row["status"],
                "payment":
                    row["payment"],
                "payment_time":
                    row["payment_time"],
                "completed_time":
                    row["completed_time"],
                "latitude":
                    row["latitude"],
                "longitude":
                    row["longitude"],
                "geo_query":
                    row["geo_query"],
            })

        route_order += 1

    route_data = pd.DataFrame(
        final_rows
    )

    # --------------------------------------------------------
    # SAVE ROUTE ORDER
    # --------------------------------------------------------

    for _, row in route_data.iterrows():

        if row["job_id"] == "DEPOT":
            continue

        save_job({
            "job_id":
                row["job_id"],
            "service_date":
                route_date_text,
            "postcode":
                row["postcode"],
            "price":
                row["price"],
            "phone":
                row["phone"],
            "status":
                row["status"],
            "payment":
                row["payment"],
            "payment_time":
                row["payment_time"],
            "completed_time":
                row["completed_time"],
            "route_order":
                int(
                    row["route_order"]
                ),
            "address_text":
                row["address_text"],
            "latitude":
                row["latitude"],
            "longitude":
                row["longitude"],
            "geo_query":
                row["geo_query"],
            "created_at":
                now_text(),
        })

    # --------------------------------------------------------
    # SAVE SESSION
    # --------------------------------------------------------

    st.session_state[
        "route_data"
    ] = route_data

    st.session_state[
        "route_metrics"
    ] = metrics

    st.session_state[
        "routing_source"
    ] = routing_source

    progress.progress(
        100,
        text="Route complete."
    )

    time.sleep(0.5)

    progress.empty()

    st.success(
        "Route successfully optimised."
    )

    st.rerun()


# ============================================================
# GET ROUTE
# ============================================================

route_data = st.session_state.get(
    "route_data"
)

route_metrics_data = st.session_state.get(
    "route_metrics"
)

routing_source = st.session_state.get(
    "routing_source",
    "Route not yet calculated",
)


# ============================================================
# IF NO ROUTE YET
# ============================================================

if route_data is None:

    saved = load_day(
        route_date_text
    )

    saved = saved[
        saved["route_order"].notna()
    ].copy()

    if not saved.empty:

        saved["route_order"] = (
            saved[
                "route_order"
            ]
            .astype(int)
        )

        route_data = saved.sort_values(
            "route_order"
        )

    else:

        st.info(
            "Press PLAN / RE-PLAN ROUTE to calculate today's route."
        )

        st.stop()


# ============================================================
# DASHBOARD METRICS
# ============================================================

customer_route = route_data[
    route_data["job_id"] != "DEPOT"
].copy()

revenue = customer_route[
    "price"
].apply(
    safe_float
).sum()

if route_metrics_data:

    driving_distance = (
        route_metrics_data[
            "distance_miles"
        ]
    )

    driving_time = (
        route_metrics_data[
            "time_s"
        ]
    )

    fuel_cost = (
        route_metrics_data[
            "fuel_cost"
        ]
    )

    fuel_used = (
        route_metrics_data[
            "litres"
        ]
    )

else:

    driving_distance = 0
    driving_time = 0
    fuel_cost = 0
    fuel_used = 0


st.markdown("---")

col1, col2, col3, col4, col5 = (
    st.columns(5)
)

col1.metric(
    "Revenue",
    f"£{revenue:,.2f}",
)

col2.metric(
    "Driving Distance",
    f"{driving_distance:,.1f} miles",
)

col3.metric(
    "Driving Time",
    format_duration(
        driving_time
    ),
)

col4.metric(
    "Fuel Cost",
    f"£{fuel_cost:,.2f}",
)

col5.metric(
    "Fuel Used",
    f"{fuel_used:,.1f} L",
)

st.caption(
    f"Routing method: {routing_source}"
)


# ============================================================
# LONG ROUTE WARNING
# ============================================================

if driving_time > 8 * 3600:

    st.warning(
        "⚠️ This route is over 8 hours of driving. "
        "The optimiser has calculated the best route it can "
        "from the available jobs, but the day may be too large "
        "to complete comfortably."
    )


# ============================================================
# FAILED / MISSING ADDRESS CHECK
# ============================================================

missing_coords = route_data[
    route_data["latitude"].isna()
    |
    route_data["longitude"].isna()
]

if not missing_coords.empty:

    st.error(
        "Some addresses do not have coordinates:"
    )

    st.dataframe(
        missing_coords[
            [
                "postcode",
                "address_text",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# NEXT STOP
# ============================================================

pending_jobs = route_data[
    (
        route_data["job_id"]
        != "DEPOT"
    )
    &
    (
        route_data["status"]
        != "Completed"
    )
].sort_values(
    "route_order"
)

if not pending_jobs.empty:

    next_stop = pending_jobs.iloc[0]

    st.sidebar.markdown("---")

    st.sidebar.subheader(
        "📍 Next stop"
    )

    st.sidebar.write(
        f"**Stop {int(next_stop['route_order'])}**"
    )

    st.sidebar.write(
        next_stop["address_text"]
        or next_stop["postcode"]
    )

    st.sidebar.link_button(
        "🗺️ Navigate",
        maps_url(
            destination_for_job(
                next_stop
            )
        ),
        use_container_width=True,
    )


# ============================================================
# ROUTE LIST
# ============================================================

st.markdown("---")
st.subheader("🧭 Today's Route")

display_route = route_data.sort_values(
    "route_order"
).copy()

for _, row in display_route.iterrows():

    order = int(
        row["route_order"]
    )

    job_id = row["job_id"]

    # --------------------------------------------------------
    # DEPOT
    # --------------------------------------------------------

    if job_id == "DEPOT":

        st.markdown(
            f"### 🏠 {order}. DEPOT"
        )

        st.caption(
            row["address_text"]
        )

        continue

    # --------------------------------------------------------
    # CUSTOMER
    # --------------------------------------------------------

    completed = (
        row["status"]
        == "Completed"
    )

    if completed:
        icon = "✅"
    else:
        icon = "🧹"

    with st.container():

        col_a, col_b, col_c = (
            st.columns(
                [1, 5, 2]
            )
        )

        with col_a:

            st.markdown(
                f"### {icon}"
            )

            st.markdown(
                f"**{order}**"
            )

        with col_b:

            st.markdown(
                f"**{row['postcode']}**"
            )

            st.caption(
                row["address_text"]
                or row["postcode"]
            )

            st.write(
                f"Price: £{safe_float(row['price']):.2f}"
            )

            if row.get(
                "phone",
                "",
            ):

                st.caption(
                    f"📞 {row['phone']}"
                )

        with col_c:

            destination = (
                destination_for_job(
                    row
                )
            )

            st.link_button(
                "🗺️ Navigate",
                maps_url(
                    destination
                ),
                use_container_width=True,
            )

            phone = clean_val(
                row.get(
                    "phone",
                    "",
                )
            )

            if phone:

                message = (
                    "Hello, this is Dan from "
                    "DanCleanUK. I'm on my way."
                )

                wa = whatsapp_url(
                    phone,
                    message,
                )

                if wa:

                    st.link_button(
                        "💬 WhatsApp",
                        wa,
                        use_container_width=True,
                    )

        # ----------------------------------------------------
        # JOB ACTIONS
        # ----------------------------------------------------

        action_col1, action_col2, action_col3 = (
            st.columns(3)
        )

        with action_col1:

            if not completed:

                if st.button(
                    "✅ Complete",
                    key=f"complete_{job_id}",
                    use_container_width=True,
                ):

                    save_job({
                        "job_id":
                            job_id,
                        "service_date":
                            route_date_text,
                        "postcode":
                            row["postcode"],
                        "price":
                            row["price"],
                        "phone":
                            row["phone"],
                        "status":
                            "Completed",
                        "payment":
                            row["payment"],
                        "payment_time":
                            row["payment_time"],
                        "completed_time":
                            now_text(),
                        "route_order":
                            order,
                        "address_text":
                            row["address_text"],
                        "latitude":
                            row["latitude"],
                        "longitude":
                            row["longitude"],
                        "geo_query":
                            row["geo_query"],
                        "created_at":
                            now_text(),
                    })

                    st.rerun()

            else:

                st.success(
                    "Completed"
                )

        with action_col2:

            payment_value = row.get(
                "payment",
                "Unpaid",
            )

            if payment_value == "Paid":

                st.success(
                    "💷 Paid"
                )

            else:

                if st.button(
                    "💷 Mark Paid",
                    key=f"paid_{job_id}",
                    use_container_width=True,
                ):

                    save_job({
                        "job_id":
                            job_id,
                        "service_date":
                            route_date_text,
                        "postcode":
                            row["postcode"],
                        "price":
                            row["price"],
                        "phone":
                            row["phone"],
                        "status":
                            row["status"],
                        "payment":
                            "Paid",
                        "payment_time":
                            now_text(),
                        "completed_time":
                            row["completed_time"],
                        "route_order":
                            order,
                        "address_text":
                            row["address_text"],
                        "latitude":
                            row["latitude"],
                        "longitude":
                            row["longitude"],
                        "geo_query":
                            row["geo_query"],
                        "created_at":
                            now_text(),
                    })

                    st.rerun()

        with action_col3:

            if st.button(
                "🔄 Re-plan",
                key=f"replan_{job_id}",
                use_container_width=True,
            ):

                st.info(
                    "Press PLAN / RE-PLAN ROUTE above "
                    "to recalculate the complete route."
                )

    st.markdown("---")


# ============================================================
# COMPLETED SUMMARY
# ============================================================

completed_df = route_data[
    (
        route_data["job_id"]
        != "DEPOT"
    )
    &
    (
        route_data["status"]
        == "Completed"
    )
].copy()

if not completed_df.empty:

    st.subheader(
        "✅ Completed Jobs"
    )

    completed_revenue = (
        completed_df["price"]
        .apply(safe_float)
        .sum()
    )

    c1, c2 = st.columns(2)

    c1.metric(
        "Completed Jobs",
        len(completed_df),
    )

    c2.metric(
        "Completed Revenue",
        f"£{completed_revenue:,.2f}",
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

st.markdown("---")

st.subheader(
    "📊 Export"
)

if st.button(
    "📥 Create Excel report",
    use_container_width=True,
):

    workbook = Workbook()

    worksheet = workbook.active
    worksheet.title = "Daily Route"

    export_df = route_data.copy()

    export_df = export_df[
        [
            "route_order",
            "postcode",
            "address_text",
            "price",
            "phone",
            "status",
            "payment",
            "payment_time",
            "completed_time",
        ]
    ]

    for row in dataframe_to_rows(
        export_df,
        index=False,
        header=True,
    ):

        worksheet.append(row)

    # Header formatting
    for cell in worksheet[1]:

        cell.font = Font(
            bold=True
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="D9EAF7",
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    # Column widths
    widths = {
        "A": 12,
        "B": 15,
        "C": 45,
        "D": 12,
        "E": 18,
        "F": 15,
        "G": 15,
        "H": 22,
        "I": 22,
    }

    for column, width in widths.items():

        worksheet.column_dimensions[
            column
        ].width = width

    worksheet.freeze_panes = "A2"

    output = io.BytesIO()

    workbook.save(
        output
    )

    output.seek(0)

    st.download_button(
        label="⬇️ Download Excel",
        data=output,
        file_name=(
            f"DanCleanUK_"
            f"{route_date_text}.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )
