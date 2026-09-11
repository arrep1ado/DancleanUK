import streamlit as st
import pandas as pd
import requests
import io
import time
import math
import random
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# ============================================================
# CONFIG
# ============================================================

APP_VERSION = "10.0"

st.set_page_config(
    page_title="DanCleanUK Route Optimizer",
    page_icon="🚗",
    layout="centered"
)

st.title("🚗 DanCleanUK Daily Route Optimizer")

st.markdown(
    """
    <style>
        html, body {
            overscroll-behavior-y: none;
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SESSION STATE
# ============================================================

if st.session_state.get("app_version") != APP_VERSION:

    for key in list(st.session_state.keys()):
        del st.session_state[key]

    st.session_state.app_version = APP_VERSION

if "geocode_cache" not in st.session_state:
    st.session_state.geocode_cache = {}


# ============================================================
# RESET
# ============================================================

if st.sidebar.button(
    "🔄 Start New Day / Reset",
    use_container_width=True
):

    for key in list(st.session_state.keys()):
        del st.session_state[key]

    st.rerun()


# ============================================================
# SETTINGS
# ============================================================

st.sidebar.title("⚙️ Settings")

DEPOT_POSTCODE = st.sidebar.text_input(
    "Depot Postcode",
    value="NG31 9RA"
)

DEPOT_FULL_ADDRESS = st.sidebar.text_input(
    "Depot Address",
    value="192 Queensway, Grantham NG31 9RA"
)

FUEL_PRICE = st.sidebar.number_input(
    "Fuel Price (£/litre)",
    min_value=0.01,
    value=1.50,
    step=0.01
)

MPG = st.sidebar.number_input(
    "Vehicle MPG",
    min_value=1.0,
    value=30.0,
    step=0.1
)

TAX_RATE = (
    st.sidebar.slider(
        "Tax Deduction (%)",
        0,
        50,
        20
    ) / 100
)


# ============================================================
# ROUTE PRIORITY SETTINGS
# ============================================================

st.sidebar.markdown("---")
st.sidebar.subheader("🧠 Route Priorities")

TIME_PRIORITY = st.sidebar.slider(
    "Driving Time Priority",
    1,
    10,
    10
)

DISTANCE_PRIORITY = st.sidebar.slider(
    "Distance/Fuel Priority",
    1,
    10,
    7
)

CLUSTER_PRIORITY = st.sidebar.slider(
    "Stay Near Nearby Jobs",
    1,
    10,
    9
)

st.sidebar.caption(
    "For a normal Grantham day, leave these near "
    "the defaults. The optimiser will strongly prefer "
    "clearing nearby jobs before travelling to a distant area."
)


# ============================================================
# API KEY
# ============================================================

if "API_KEY" not in st.secrets:

    st.error(
        "API_KEY missing in Streamlit secrets."
    )

    st.info(
        "Add your OpenRouteService API key as API_KEY."
    )

    st.stop()

API_KEY = st.secrets["API_KEY"]


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_val(value):

    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.endswith(".0"):
        try:
            text = str(int(float(text)))
        except Exception:
            pass

    return text


def maps_url(destination):

    return (
        "https://www.google.com/maps/dir/?api=1"
        f"&destination={quote(str(destination))}"
        "&travelmode=driving"
    )


def format_duration(seconds):

    minutes = round(float(seconds) / 60)

    hours = minutes // 60
    mins = minutes % 60

    if hours:
        return f"{hours}h {mins}m"

    return f"{mins}m"


# ============================================================
# BUILD GEOCODING QUERY
# ============================================================

def build_geo_query(row, default_postcode):

    parts = []

    address_columns = [
        "address",
        "street",
        "location",
        "house",
        "house number",
        "house_number",
        "property",
        "property address",
        "address line 1",
        "address1"
    ]

    for column in row.index:

        if column.lower() in address_columns:

            value = clean_val(
                row[column]
            )

            if value:
                parts.append(value)

    postcode = clean_val(
        row.get("Postcode", "")
    )

    if postcode:
        parts.append(postcode)
    else:
        parts.append(default_postcode)

    parts.append("United Kingdom")

    return ", ".join(parts)


# ============================================================
# GEOCODER
# ============================================================

def get_coords(query_string, postcode):

    query = str(query_string).strip()
    postcode = str(postcode).upper().strip()

    cache_key = (
        query + "|" + postcode
    ).lower()

    if cache_key in st.session_state.geocode_cache:

        return st.session_state.geocode_cache[
            cache_key
        ]

    headers = {
        "User-Agent":
            "DanCleanUKRouteOptimizer/10.0"
    }

    # --------------------------------------------------------
    # FULL ADDRESS
    # --------------------------------------------------------

    if query:

        try:

            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": query,
                    "format": "json",
                    "limit": 1,
                    "countrycodes": "gb"
                },
                headers=headers,
                timeout=8
            )

            if (
                response.status_code == 200
                and response.json()
            ):

                item = response.json()[0]

                coords = (
                    float(item["lat"]),
                    float(item["lon"])
                )

                st.session_state.geocode_cache[
                    cache_key
                ] = coords

                return coords

        except Exception:
            pass

    # --------------------------------------------------------
    # POSTCODE FALLBACK - NOMINATIM
    # --------------------------------------------------------

    if postcode:

        try:

            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": f"{postcode}, United Kingdom",
                    "format": "json",
                    "limit": 1,
                    "countrycodes": "gb"
                },
                headers=headers,
                timeout=8
            )

            if (
                response.status_code == 200
                and response.json()
            ):

                item = response.json()[0]

                coords = (
                    float(item["lat"]),
                    float(item["lon"])
                )

                st.session_state.geocode_cache[
                    cache_key
                ] = coords

                return coords

        except Exception:
            pass

    # --------------------------------------------------------
    # POSTCODES.IO
    # --------------------------------------------------------

    pc = postcode.replace(" ", "")

    if pc:

        try:

            response = requests.get(
                f"https://api.postcodes.io/postcodes/{pc}",
                timeout=6
            )

            if response.status_code == 200:

                result = (
                    response.json()
                    .get("result")
                )

                if result:

                    lat = result.get("latitude")
                    lon = result.get("longitude")

                    if (
                        lat is not None
                        and lon is not None
                    ):

                        coords = (
                            float(lat),
                            float(lon)
                        )

                        st.session_state.geocode_cache[
                            cache_key
                        ] = coords

                        return coords

        except Exception:
            pass

    # IMPORTANT:
    # Never silently put a failed customer at the depot.

    return None


# ============================================================
# HAVERSINE
# ============================================================

def haversine_km(
    lat1,
    lon1,
    lat2,
    lon2
):

    R = 6371.0

    dlat = math.radians(
        lat2 - lat1
    )

    dlon = math.radians(
        lon2 - lon1
    )

    a = (
        math.sin(dlat / 2) ** 2
        +
        math.cos(math.radians(lat1))
        *
        math.cos(math.radians(lat2))
        *
        math.sin(dlon / 2) ** 2
    )

    c = (
        2
        * math.asin(
            math.sqrt(a)
        )
    )

    return R * c


# ============================================================
# OFFLINE MATRIX
# ============================================================

def offline_matrix(locations):

    n = len(locations)

    distances = [
        [0.0] * n
        for _ in range(n)
    ]

    durations = [
        [0.0] * n
        for _ in range(n)
    ]

    ROAD_FACTOR = 1.30
    AVERAGE_SPEED = 35.0

    for i in range(n):

        lon1, lat1 = locations[i]

        for j in range(n):

            if i == j:
                continue

            lon2, lat2 = locations[j]

            km = haversine_km(
                lat1,
                lon1,
                lat2,
                lon2
            )

            road_km = km * ROAD_FACTOR

            distances[i][j] = (
                road_km * 1000
            )

            durations[i][j] = (
                road_km
                / AVERAGE_SPEED
                * 3600
            )

    return distances, durations


# ============================================================
# OPENROUTESERVICE MATRIX
# ============================================================

def get_ors_matrix(locations):

    try:

        response = requests.post(
            "https://api.openrouteservice.org/v2/matrix/driving-car",
            json={
                "locations": locations,
                "metrics": [
                    "distance",
                    "duration"
                ],
                "units": "m"
            },
            headers={
                "Authorization": API_KEY,
                "Content-Type": "application/json"
            },
            timeout=60
        )

        if response.status_code != 200:
            return None, None

        data = response.json()

        return (
            data.get("distances"),
            data.get("durations")
        )

    except Exception:

        return None, None


# ============================================================
# ROUTE METRICS
# ============================================================

def route_metrics(
    route,
    distances,
    durations,
    fuel_price,
    mpg
):

    total_distance = 0.0
    total_time = 0.0

    for i in range(
        len(route) - 1
    ):

        a = route[i]
        b = route[i + 1]

        total_distance += (
            distances[a][b]
        )

        total_time += (
            durations[a][b]
        )

    miles = (
        total_distance
        / 1000
        * 0.621371
    )

    litres = (
        miles
        / mpg
        * 4.54609
    )

    fuel_cost = (
        litres
        * fuel_price
    )

    return {
        "distance_m": total_distance,
        "time_s": total_time,
        "miles": miles,
        "litres": litres,
        "fuel_cost": fuel_cost
    }


# ============================================================
# CLUSTER PENALTY
# ============================================================

def calculate_cluster_penalty(
    route,
    distances,
    durations
):

    penalty = 0.0

    # --------------------------------------------------------
    # This is NOT saying "always choose nearest".
    #
    # It asks:
    #
    # "Are we making a big jump while there are still
    # several customers close to where we currently are?"
    #
    # If yes, penalise that behaviour.
    # --------------------------------------------------------

    customers = route[1:-1]

    for position in range(
        1,
        len(route) - 1
    ):

        current = route[position]

        remaining = [
            x
            for x in customers
            if x not in route[:position + 1]
        ]

        if not remaining:
            continue

        next_stop = route[position + 1]

        next_distance = (
            distances[current][next_stop]
        )

        # Find nearby unvisited jobs.
        nearby = []

        for job in remaining:

            d = distances[current][job]

            if d <= 15000:
                nearby.append(d)

        # If there are nearby jobs and we're leaving
        # the area for a much longer journey, penalise it.
        if nearby:

            nearest_local = min(
                nearby
            )

            if next_distance > (
                nearest_local * 2.5
            ):

                penalty += (
                    next_distance
                    - nearest_local
                ) * 3.0

    return penalty


# ============================================================
# ROUTE SCORE
# ============================================================

def route_score(
    route,
    distances,
    durations,
    fuel_price,
    mpg
):

    metrics = route_metrics(
        route,
        distances,
        durations,
        fuel_price,
        mpg
    )

    cluster_penalty = (
        calculate_cluster_penalty(
            route,
            distances,
            durations
        )
    )

    # Time is the most important factor.
    time_score = (
        metrics["time_s"]
        * TIME_PRIORITY
    )

    distance_score = (
        metrics["distance_m"]
        * DISTANCE_PRIORITY
    )

    cluster_score = (
        cluster_penalty
        * CLUSTER_PRIORITY
    )

    return (
        time_score
        + distance_score
        + cluster_score
    )


# ============================================================
# NEAREST FIRST
# ============================================================

def find_first_customer(
    distances,
    durations
):

    best = None
    best_time = float("inf")

    for customer in range(
        1,
        len(distances)
    ):

        travel_time = (
            durations[0][customer]
        )

        if travel_time < best_time:

            best_time = travel_time
            best = customer

    return best


# ============================================================
# GREEDY ROUTE
# ============================================================

def build_greedy_route(
    first_customer,
    distances,
    durations
):

    customer_count = (
        len(distances) - 1
    )

    route = [
        0,
        first_customer
    ]

    remaining = set(
        range(
            1,
            customer_count + 1
        )
    )

    remaining.discard(
        first_customer
    )

    current = first_customer

    while remaining:

        candidates = []

        for candidate in remaining:

            direct_time = (
                durations[current][candidate]
            )

            # Look one step ahead.
            future = [
                x
                for x in remaining
                if x != candidate
            ]

            if future:

                nearest_future = min(
                    future,
                    key=lambda x:
                        durations[candidate][x]
                )

                future_time = (
                    durations[candidate]
                    [nearest_future]
                )

            else:

                future_time = (
                    durations[candidate][0]
                )

            # Mostly nearest next stop,
            # but look ahead to prevent bad jumps.
            score = (
                direct_time * 0.70
                +
                future_time * 0.30
            )

            candidates.append(
                (
                    score,
                    candidate
                )
            )

        candidates.sort(
            key=lambda x: x[0]
        )

        next_customer = (
            candidates[0][1]
        )

        route.append(
            next_customer
        )

        remaining.remove(
            next_customer
        )

        current = next_customer

    route.append(0)

    return route


# ============================================================
# 2-OPT
# ============================================================

def two_opt(
    route,
    distances,
    durations,
    fuel_price,
    mpg
):

    best = route[:]

    best_score = route_score(
        best,
        distances,
        durations,
        fuel_price,
        mpg
    )

    improved = True

    while improved:

        improved = False

        for i in range(
            2,
            len(best) - 2
        ):

            for j in range(
                i + 1,
                len(best) - 1
            ):

                candidate = (
                    best[:i]
                    +
                    best[i:j + 1][::-1]
                    +
                    best[j + 1:]
                )

                # First stop must remain first.
                if candidate[1] != route[1]:
                    continue

                candidate_score = (
                    route_score(
                        candidate,
                        distances,
                        durations,
                        fuel_price,
                        mpg
                    )
                )

                if candidate_score < (
                    best_score - 0.01
                ):

                    best = candidate
                    best_score = candidate_score

                    improved = True

                    break

            if improved:
                break

    return best


# ============================================================
# 3-OPT STYLE RELOCATION
# ============================================================

def relocate_improvement(
    route,
    distances,
    durations,
    fuel_price,
    mpg
):

    best = route[:]

    best_score = route_score(
        best,
        distances,
        durations,
        fuel_price,
        mpg
    )

    improved = True

    while improved:

        improved = False

        # Don't move the first customer.
        for i in range(
            2,
            len(best) - 1
        ):

            customer = best[i]

            shortened = (
                best[:i]
                +
                best[i + 1:]
            )

            for j in range(
                1,
                len(shortened)
            ):

                # Never put anything before the
                # compulsory first customer.
                if j < 2:
                    continue

                candidate = (
                    shortened[:j]
                    + [customer]
                    + shortened[j:]
                )

                if candidate[1] != best[1]:
                    continue

                candidate_score = (
                    route_score(
                        candidate,
                        distances,
                        durations,
                        fuel_price,
                        mpg
                    )
                )

                if candidate_score < (
                    best_score - 0.01
                ):

                    best = candidate
                    best_score = candidate_score

                    improved = True

                    break

            if improved:
                break

    return best


# ============================================================
# CREATE MULTIPLE STARTING ROUTES
# ============================================================

def generate_candidate_routes(
    first_customer,
    distances,
    durations
):

    candidates = []

    # Candidate 1:
    # normal look-ahead greedy.
    candidates.append(
        build_greedy_route(
            first_customer,
            distances,
            durations
        )
    )

    customer_count = (
        len(distances) - 1
    )

    remaining = [
        x
        for x in range(
            1,
            customer_count + 1
        )
        if x != first_customer
    ]

    # --------------------------------------------------------
    # Try several possible second stops.
    #
    # This is extremely useful for your situation.
    #
    # Example:
    #
    # Depot
    # -> nearest Grantham job
    #
    # There might be:
    #   5 more Grantham jobs
    #   4 village jobs
    #   10 distant jobs
    #
    # We test several sensible local second stops instead
    # of blindly committing to one.
    # --------------------------------------------------------

    second_candidates = sorted(
        remaining,
        key=lambda x:
            durations[first_customer][x]
    )

    for second in second_candidates[:15]:

        route = [
            0,
            first_customer,
            second
        ]

        unvisited = set(
            remaining
        )

        unvisited.discard(
            second
        )

        current = second

        while unvisited:

            candidate_scores = []

            for candidate in unvisited:

                direct = (
                    durations[current][candidate]
                )

                future_jobs = [
                    x
                    for x in unvisited
                    if x != candidate
                ]

                if future_jobs:

                    future = min(
                        future_jobs,
                        key=lambda x:
                            durations[candidate][x]
                    )

                    lookahead = (
                        durations[candidate][future]
                    )

                else:

                    lookahead = (
                        durations[candidate][0]
                    )

                score = (
                    direct * 0.65
                    +
                    lookahead * 0.35
                )

                candidate_scores.append(
                    (
                        score,
                        candidate
                    )
                )

            candidate_scores.sort(
                key=lambda x: x[0]
            )

            chosen = (
                candidate_scores[0][1]
            )

            route.append(
                chosen
            )

            unvisited.remove(
                chosen
            )

            current = chosen

        route.append(0)

        candidates.append(
            route
        )

    # --------------------------------------------------------
    # Randomised candidates.
    #
    # This helps escape a poor greedy route.
    # --------------------------------------------------------

    for _ in range(20):

        shuffled = remaining[:]

        random.shuffle(
            shuffled
        )

        route = [
            0,
            first_customer
        ] + shuffled + [0]

        candidates.append(
            route
        )

    return candidates


# ============================================================
# FULL OPTIMISER
# ============================================================

def optimise_route(
    distances,
    durations,
    fuel_price,
    mpg
):

    # --------------------------------------------------------
    # FIRST STOP IS ALWAYS THE CLOSEST CUSTOMER TO DEPOT.
    # --------------------------------------------------------

    first_customer = (
        find_first_customer(
            distances,
            durations
        )
    )

    candidates = (
        generate_candidate_routes(
            first_customer,
            distances,
            durations
        )
    )

    best_route = None
    best_score = float("inf")

    # --------------------------------------------------------
    # Improve every candidate.
    # --------------------------------------------------------

    for candidate in candidates:

        improved = two_opt(
            candidate,
            distances,
            durations,
            fuel_price,
            mpg
        )

        improved = relocate_improvement(
            improved,
            distances,
            durations,
            fuel_price,
            mpg
        )

        improved = two_opt(
            improved,
            distances,
            durations,
            fuel_price,
            mpg
        )

        score = route_score(
            improved,
            distances,
            durations,
            fuel_price,
            mpg
        )

        if score < best_score:

            best_score = score
            best_route = improved

    return (
        best_route,
        first_customer
    )


# ============================================================
# DESTINATION TEXT
# ============================================================

def get_destination(row):

    geo_query = row.get(
        "geo_query",
        ""
    )

    if (
        pd.notna(geo_query)
        and str(geo_query).strip()
    ):

        return str(
            geo_query
        )

    parts = []

    for column in row.index:

        if column.lower() in [
            "address",
            "street",
            "location",
            "house",
            "house number"
        ]:

            value = clean_val(
                row.get(column)
            )

            if value:
                parts.append(value)

    postcode = clean_val(
        row.get("Postcode")
    )

    if postcode:
        parts.append(postcode)

    return ", ".join(parts)


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📁 Upload your day's file",
    type=[
        "csv",
        "xlsx"
    ]
)


if (
    uploaded_file
    and "master_df" not in st.session_state
):

    try:

        if uploaded_file.name.lower().endswith(
            ".xlsx"
        ):

            df = pd.read_excel(
                uploaded_file
            )

        else:

            df = pd.read_csv(
                uploaded_file
            )

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        required = [
            "Postcode",
            "Price",
            "Phone"
        ]

        missing = [
            x
            for x in required
            if x not in df.columns
        ]

        if missing:

            st.error(
                "Missing required columns: "
                + ", ".join(missing)
            )

            st.stop()

        # Remove empty Excel rows.
        df = df.dropna(
            how="all"
        ).copy()

        # ----------------------------------------------------
        # POSTCODE
        # ----------------------------------------------------

        df["Postcode"] = (
            df["Postcode"]
            .fillna("")
            .astype(str)
            .str.upper()
            .str.strip()
        )

        df = df[
            ~df["Postcode"].isin(
                [
                    "",
                    "NAN",
                    "NAT"
                ]
            )
        ].copy()

        # ----------------------------------------------------
        # PRICE
        # ----------------------------------------------------

        df["Price"] = pd.to_numeric(
            df["Price"],
            errors="coerce"
        )

        df = df.dropna(
            subset=["Price"]
        )

        df = df[
            df["Price"] > 0
        ].copy()

        # ----------------------------------------------------
        # PHONE
        # ----------------------------------------------------

        df["Phone"] = (
            df["Phone"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df = df[
            ~df["Phone"]
            .str.lower()
            .isin(
                [
                    "",
                    "nan",
                    "nat"
                ]
            )
        ].copy()

        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------

        df["Status"] = "pending"
        df["Payment"] = "waiting"

        st.session_state.master_df = (
            df.reset_index(drop=True)
        )

        st.rerun()

    except Exception as e:

        st.error(
            f"Error loading file: {e}"
        )


# ============================================================
# WAIT FOR FILE
# ============================================================

if "master_df" not in st.session_state:

    st.info(
        "Upload your day's file to begin."
    )

    st.stop()


df = st.session_state.master_df


# ============================================================
# PLAN ROUTE BUTTON
# ============================================================

if st.button(
    "🚀 PLAN BEST DAILY ROUTE",
    type="primary",
    use_container_width=True
):

    if df.empty:

        st.error(
            "No valid customer jobs found."
        )

        st.stop()

    # --------------------------------------------------------
    # DEPOT
    # --------------------------------------------------------

    with st.spinner(
        "📍 Locating your depot..."
    ):

        depot_coords = get_coords(
            DEPOT_FULL_ADDRESS,
            DEPOT_POSTCODE
        )

    if depot_coords is None:

        st.error(
            "Could not locate the depot."
        )

        st.stop()

    # --------------------------------------------------------
    # GEOCODE CUSTOMERS
    # --------------------------------------------------------

    total = len(df)

    progress = st.progress(
        0,
        text="Locating customer addresses..."
    )

    valid = []
    failed = []

    coords_by_index = {}

    for number, (idx, row) in enumerate(
        df.iterrows(),
        start=1
    ):

        query = build_geo_query(
            row,
            DEPOT_POSTCODE
        )

        coords = get_coords(
            query,
            row["Postcode"]
        )

        if coords is None:

            failed.append(idx)

        else:

            valid.append(idx)

            coords_by_index[idx] = coords

        progress.progress(
            number / total,
            text=(
                f"Locating customer "
                f"{number}/{total}"
            )
        )

        # Nominatim polite rate.
        time.sleep(1)

    progress.empty()

    # --------------------------------------------------------
    # FAILED JOBS
    # --------------------------------------------------------

    if failed:

        st.session_state.failed_jobs = (
            df.loc[failed].copy()
        )

    else:

        st.session_state.failed_jobs = (
            pd.DataFrame()
        )

    if not valid:

        st.error(
            "No customer addresses could be located."
        )

        st.stop()

    # --------------------------------------------------------
    # ROUTING DATA
    # --------------------------------------------------------

    rows = []

    locations = [
        [
            depot_coords[1],
            depot_coords[0]
        ]
    ]

    depot_row = {

        "Postcode":
            DEPOT_POSTCODE,

        "Price":
            0.0,

        "Phone":
            "",

        "Status":
            "depot",

        "Payment":
            "waiting",

        "geo_query":
            DEPOT_FULL_ADDRESS,

        "latitude":
            depot_coords[0],

        "longitude":
            depot_coords[1]
    }

    rows.append(
        depot_row
    )

    for idx in valid:

        original = df.loc[idx]

        row = original.to_dict()

        lat, lon = coords_by_index[idx]

        row["geo_query"] = (
            build_geo_query(
                original,
                DEPOT_POSTCODE
            )
        )

        row["latitude"] = lat
        row["longitude"] = lon

        rows.append(
            row
        )

        locations.append(
            [
                lon,
                lat
            ]
        )

    routing_df = pd.DataFrame(
        rows
    )

    # --------------------------------------------------------
    # MATRIX
    # --------------------------------------------------------

    with st.spinner(
        "🛣️ Getting actual road distances and driving times..."
    ):

        distances, durations = (
            get_ors_matrix(
                locations
            )
        )

    using_offline = False

    if (
        distances is None
        or durations is None
    ):

        using_offline = True

        st.warning(
            "OpenRouteService could not provide the "
            "road matrix. Using an offline estimate."
        )

        distances, durations = (
            offline_matrix(
                locations
            )
        )

    # --------------------------------------------------------
    # OPTIMISE
    # --------------------------------------------------------

    with st.spinner(
        "🧠 Optimising the complete 30–40 stop route..."
    ):

        route, first_customer = (
            optimise_route(
                distances,
                durations,
                FUEL_PRICE,
                MPG
            )
        )

    # --------------------------------------------------------
    # FINAL METRICS
    # --------------------------------------------------------

    metrics = route_metrics(
        route,
        distances,
        durations,
        FUEL_PRICE,
        MPG
    )

    revenue = float(
        routing_df["Price"].sum()
    )

    fuel_cost = (
        metrics["fuel_cost"]
    )

    pre_tax_profit = (
        revenue
        - fuel_cost
    )

    take_home = (
        pre_tax_profit
        * (1 - TAX_RATE)
    )

    # --------------------------------------------------------
    # SAVE ROUTE
    # --------------------------------------------------------

    final_df = (
        routing_df
        .iloc[route]
        .reset_index(drop=True)
        .copy()
    )

    final_df.loc[
        0,
        "Status"
    ] = "depot"

    final_df.loc[
        len(final_df) - 1,
        "Status"
    ] = "depot"

    st.session_state.master_df = (
        final_df
    )

    st.session_state.route_data = {

        "revenue":
            revenue,

        "fuel_cost":
            fuel_cost,

        "take_home":
            take_home,

        "miles":
            metrics["miles"],

        "litres":
            metrics["litres"],

        "time":
            metrics["time_s"],

        "offline":
            using_offline
    }

    st.rerun()


# ============================================================
# DASHBOARD
# ============================================================

if "route_data" in st.session_state:

    data = (
        st.session_state.route_data
    )

    st.markdown("---")

    st.subheader(
        "💰 Daily Route Summary"
    )

    c1, c2 = st.columns(2)

    with c1:

        st.metric(
            "Take-Home",
            f"£{data['take_home']:.2f}"
        )

    with c2:

        st.metric(
            "Revenue",
            f"£{data['revenue']:.2f}"
        )

    c3, c4 = st.columns(2)

    with c3:

        st.metric(
            "Driving Distance",
            f"{data['miles']:.1f} miles"
        )

    with c4:

        st.metric(
            "Driving Time",
            format_duration(
                data["time"]
            )
        )

    c5, c6 = st.columns(2)

    with c5:

        st.metric(
            "Fuel Cost",
            f"£{data['fuel_cost']:.2f}"
        )

    with c6:

        st.metric(
            "Fuel Used",
            f"{data['litres']:.1f} litres"
        )

    if data["offline"]:

        st.warning(
            "This route used estimated distances because "
            "the live road-routing matrix was unavailable."
        )

    else:

        st.success(
            "✅ Route calculated using actual driving "
            "distance and time."
        )


# ============================================================
# FAILED ADDRESSES
# ============================================================

if (
    "failed_jobs" in st.session_state
    and not st.session_state.failed_jobs.empty
):

    st.markdown("---")

    st.error(
        f"{len(st.session_state.failed_jobs)} "
        "customer(s) could not be located."
    )

    st.caption(
        "These jobs were NOT included in the route."
    )

    with st.expander(
        "View unlocated customers"
    ):

        st.dataframe(
            st.session_state.failed_jobs,
            use_container_width=True
        )


# ============================================================
# SIDEBAR NEXT STOP
# ============================================================

if (
    "route_data" in st.session_state
    and not st.session_state.master_df.empty
):

    st.sidebar.markdown("---")

    st.sidebar.subheader(
        "🧭 Route Navigation"
    )

    route_df = (
        st.session_state.master_df
    )

    customers = route_df[
        route_df["Status"]
        .astype(str)
        .str.lower()
        != "depot"
    ]

    pending = customers[
        customers["Status"]
        .astype(str)
        .str.lower()
        == "pending"
    ]

    if not pending.empty:

        next_row = pending.iloc[0]

        destination = (
            get_destination(
                next_row
            )
        )

        st.sidebar.link_button(
            "🚗 Navigate to Next Stop",
            maps_url(
                destination
            ),
            use_container_width=True
        )

        st.sidebar.caption(
            f"Next: {destination}"
        )

        st.sidebar.caption(
            f"{len(pending)} stops remaining"
        )

    else:

        st.sidebar.success(
            "🎉 All customer stops completed!"
        )


# ============================================================
# ROUTE LIST
# ============================================================

st.markdown("---")

st.subheader(
    "📍 Planned Route"
)

route_df = (
    st.session_state.master_df
)


for idx, row in route_df.iterrows():

    # --------------------------------------------------------
    # START DEPOT
    # --------------------------------------------------------

    if idx == 0:

        with st.container(
            border=True
        ):

            st.write(
                "### 🏠 START — GRANTHAM DEPOT"
            )

            st.write(
                DEPOT_FULL_ADDRESS
            )

        continue

    # --------------------------------------------------------
    # RETURN DEPOT
    # --------------------------------------------------------

    if idx == len(route_df) - 1:

        with st.container(
            border=True
        ):

            st.write(
                "### 🏁 FINISH — GRANTHAM DEPOT"
            )

            st.write(
                DEPOT_FULL_ADDRESS
            )

            st.link_button(
                "🚗 Navigate Back to Depot",
                maps_url(
                    DEPOT_FULL_ADDRESS
                ),
                key=f"return_{idx}",
                use_container_width=True
            )

        continue

    # --------------------------------------------------------
    # CUSTOMER
    # --------------------------------------------------------

    postcode = clean_val(
        row.get("Postcode")
    )

    price = float(
        row.get(
            "Price",
            0
        )
    )

    phone = clean_val(
        row.get("Phone")
    )

    status = str(
        row.get(
            "Status",
            "pending"
        )
    ).lower()

    payment = str(
        row.get(
            "Payment",
            "waiting"
        )
    )

    destination = (
        get_destination(
            row
        )
    )

    if status == "completed":

        icon = "✅"
        text = "Completed"

    else:

        icon = "⏳"
        text = "Pending"

    # --------------------------------------------------------
    # EXTRA COLUMNS
    # --------------------------------------------------------

    extra = []

    for column in route_df.columns:

        if column.lower() in [
            "postcode",
            "price",
            "phone",
            "status",
            "payment",
            "latitude",
            "longitude",
            "geo_query"
        ]:
            continue

        value = clean_val(
            row.get(column)
        )

        if value:

            extra.append(
                f"**{column}:** {value}"
            )

    with st.container(
        border=True
    ):

        st.write(
            f"### {icon} STOP {idx} — {postcode}"
        )

        if extra:

            st.write(
                " | ".join(extra)
            )

        st.write(
            f"**Price:** £{price:.2f}"
            f" | **Status:** {text}"
            f" | **Payment:** {payment}"
        )

        col1, col2 = st.columns(2)

        with col1:

            if status == "pending":

                if st.button(
                    "✅ Mark Complete",
                    key=f"complete_{idx}",
                    use_container_width=True
                ):

                    st.session_state.master_df.at[
                        idx,
                        "Status"
                    ] = "completed"

                    st.rerun()

            else:

                st.success(
                    "Completed"
                )

                if phone:

                    message = (
                        "Hi from DanCleanUK! "
                        "Your service is complete today. "
                        f"Total: £{price:.2f}. "
                        "Please pay via bank transfer "
                        "to Mettle - Sort Code: 04-03-33 "
                        "Account: 72515806. "
                        "Thank you!"
                    )

                    whatsapp = (
                        "https://wa.me/"
                        f"{quote(phone)}"
                        "?text="
                        f"{quote(message)}"
                    )

                    st.link_button(
                        "💬 Send WhatsApp",
                        whatsapp,
                        use_container_width=True
                    )

        with col2:

            st.link_button(
                "🚗 Navigate Here",
                maps_url(
                    destination
                ),
                key=f"nav_{idx}",
                use_container_width=True
            )

        pay1, pay2 = st.columns(2)

        with pay1:

            if st.button(
                "💵 Cash Paid",
                key=f"cash_{idx}",
                use_container_width=True
            ):

                st.session_state.master_df.at[
                    idx,
                    "Payment"
                ] = "Cash"

                st.rerun()

        with pay2:

            if st.button(
                "❌ Not Paid",
                key=f"notpaid_{idx}",
                use_container_width=True
            ):

                st.session_state.master_df.at[
                    idx,
                    "Payment"
                ] = "Not Paid"

                st.rerun()


# ============================================================
# EXPORT
# ============================================================

st.sidebar.markdown("---")

st.sidebar.subheader(
    "📊 Export Records"
)

if (
    "master_df" in st.session_state
    and not st.session_state.master_df.empty
):

    export_df = (
        st.session_state.master_df.copy()
    )

    export_df = export_df[
        export_df["Status"]
        .astype(str)
        .str.lower()
        != "depot"
    ].copy()

    for column in [
        "latitude",
        "longitude",
        "geo_query"
    ]:

        if column in export_df.columns:

            export_df.drop(
                columns=[column],
                inplace=True
            )

    output = io.BytesIO()

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Daily Report"

    header_fill = PatternFill(
        start_color="2F4F4F",
        end_color="2F4F4F",
        fill_type="solid"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    for row in dataframe_to_rows(
        export_df,
        index=False,
        header=True
    ):

        worksheet.append(
            row
        )

    for cell in worksheet[1]:

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center"
        )

    for column_cells in worksheet.columns:

        max_length = 0

        letter = (
            column_cells[0]
            .column_letter
        )

        for cell in column_cells:

            try:

                max_length = max(
                    max_length,
                    len(
                        str(
                            cell.value
                        )
                    )
                )

            except Exception:
                pass

        worksheet.column_dimensions[
            letter
        ].width = min(
            max_length + 2,
            50
        )

    workbook.save(
        output
    )

    st.sidebar.download_button(
        "⬇️ Download Excel Report",
        output.getvalue(),
        "DanCleanUK_Records.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
