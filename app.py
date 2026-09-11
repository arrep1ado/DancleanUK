import streamlit as st
import pandas as pd
import requests
import io
import time
import math
from itertools import permutations
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# ============================================================
# APP CONFIG
# ============================================================

APP_VERSION = "7.0"

st.set_page_config(
    page_title="DanCleanUK Optimizer",
    page_icon="🚗",
    layout="centered"
)

st.title("🚗 Daily Route & Profit Optimizer")

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

    for key in [
        "master_df",
        "route_data",
        "unrouted_df",
        "geocode_cache"
    ]:
        st.session_state.pop(key, None)

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
        min_value=0,
        max_value=50,
        value=20
    ) / 100
)

TIME_VALUE = st.sidebar.number_input(
    "Value of Driving Time (£/hour)",
    min_value=0.0,
    value=20.0,
    step=1.0
)

st.sidebar.caption(
    "Higher driving-time value makes the optimiser "
    "prefer faster routes, even when the distance is "
    "slightly longer."
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
# HELPERS
# ============================================================

def clean_val(value):

    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.endswith(".0"):
        try:
            value = str(int(float(value)))
        except Exception:
            pass

    return value


def format_duration(seconds):

    if seconds is None:
        return "0m"

    minutes = round(float(seconds) / 60)

    hours = minutes // 60
    mins = minutes % 60

    if hours:
        return f"{hours}h {mins}m"

    return f"{mins}m"


def maps_url(destination):

    return (
        "https://www.google.com/maps/dir/?api=1"
        f"&destination={quote(str(destination))}"
        "&travelmode=driving"
    )


# ============================================================
# BUILD BEST ADDRESS QUERY
# ============================================================

def build_geo_query(row, default_postcode):

    parts = []

    preferred_columns = [
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

    # Add address information first.
    for column in row.index:

        if column.lower() in preferred_columns:

            value = clean_val(
                row[column]
            )

            if value:
                parts.append(value)

    # Always add postcode.
    postcode = clean_val(
        row.get("Postcode", "")
    )

    if postcode:
        parts.append(postcode)
    else:
        parts.append(default_postcode)

    # Add UK so geocoder doesn't confuse locations.
    parts.append("United Kingdom")

    return ", ".join(parts)


# ============================================================
# GEOCODING
# ============================================================

def get_coords(query_string, postcode_fallback):

    query = str(query_string).strip()
    postcode = str(postcode_fallback).upper().strip()

    cache_key = (
        query + "|" + postcode
    ).lower()

    if cache_key in st.session_state.geocode_cache:
        return st.session_state.geocode_cache[
            cache_key
        ]

    headers = {
        "User-Agent":
            "DanCleanUKOptimizer/7.0"
    }

    # --------------------------------------------------------
    # 1. Full address
    # --------------------------------------------------------

    if query and query.lower() != "nan":

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

                result = response.json()[0]

                coords = (
                    float(result["lat"]),
                    float(result["lon"])
                )

                st.session_state.geocode_cache[
                    cache_key
                ] = coords

                return coords

        except Exception:
            pass

    # --------------------------------------------------------
    # 2. Postcode through Nominatim
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

                result = response.json()[0]

                coords = (
                    float(result["lat"]),
                    float(result["lon"])
                )

                st.session_state.geocode_cache[
                    cache_key
                ] = coords

                return coords

        except Exception:
            pass

    # --------------------------------------------------------
    # 3. Postcodes.io
    # --------------------------------------------------------

    postcode_clean = (
        postcode
        .replace(" ", "")
    )

    if postcode_clean:

        try:

            response = requests.get(
                f"https://api.postcodes.io/postcodes/"
                f"{postcode_clean}",
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

    # --------------------------------------------------------
    # DO NOT PUT FAILED JOB AT DEPOT
    # --------------------------------------------------------

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
        * math.cos(math.radians(lat2))
        * math.sin(dlon / 2) ** 2
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
    AVERAGE_SPEED = 40.0

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

            road_km = (
                km * ROAD_FACTOR
            )

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
            timeout=40
        )

        if response.status_code != 200:
            return None, None

        data = response.json()

        distances = data.get(
            "distances"
        )

        durations = data.get(
            "durations"
        )

        if not distances or not durations:
            return None, None

        return distances, durations

    except Exception:
        return None, None


# ============================================================
# LEG COST
# ============================================================

def leg_cost(
    distance_m,
    duration_s,
    fuel_price,
    mpg,
    time_value
):

    miles = (
        distance_m
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

    time_cost = (
        duration_s
        / 3600
        * time_value
    )

    return (
        fuel_cost
        + time_cost
    )


# ============================================================
# ROUTE TOTAL
# ============================================================

def route_total(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    time_value
):

    total_distance = 0
    total_duration = 0
    total_cost = 0

    for position in range(
        len(route) - 1
    ):

        a = route[position]
        b = route[position + 1]

        distance = distances[a][b]
        duration = durations[a][b]

        total_distance += distance
        total_duration += duration

        total_cost += leg_cost(
            distance,
            duration,
            fuel_price,
            mpg,
            time_value
        )

    return (
        total_cost,
        total_distance,
        total_duration
    )


# ============================================================
# FIND FIRST CUSTOMER
# ============================================================

def find_nearest_first_customer(
    distances,
    durations,
    fuel_price,
    mpg,
    time_value
):

    customer_count = (
        len(distances) - 1
    )

    best_customer = None
    best_score = float("inf")

    for customer in range(
        1,
        customer_count + 1
    ):

        # PRIMARY:
        # driving time from depot
        #
        # SECONDARY:
        # actual operating cost
        #
        # This guarantees the first stop is
        # genuinely near the depot.

        driving_time = (
            durations[0][customer]
        )

        operating_cost = leg_cost(
            distances[0][customer],
            durations[0][customer],
            fuel_price,
            mpg,
            time_value
        )

        score = (
            driving_time,
            operating_cost
        )

        if (
            best_customer is None
            or score < best_score
        ):

            best_customer = customer
            best_score = score

    return best_customer


# ============================================================
# EXACT OPTIMISATION AFTER FIRST STOP
# ============================================================

def optimise_after_first_exact(
    first_customer,
    distances,
    durations,
    fuel_price,
    mpg,
    time_value
):

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
    # No other customers.
    # --------------------------------------------------------

    if not remaining:

        return [
            0,
            first_customer,
            0
        ]

    # --------------------------------------------------------
    # Up to 9 remaining customers:
    # exact permutation search.
    #
    # This is intentionally done AFTER fixing the first
    # customer.
    # --------------------------------------------------------

    if len(remaining) <= 9:

        best_route = None
        best_score = None

        for permutation in permutations(
            remaining
        ):

            route = (
                [0, first_customer]
                + list(permutation)
                + [0]
            )

            score = route_total(
                route,
                distances,
                durations,
                fuel_price,
                mpg,
                time_value
            )

            # Primary objective:
            # total driving time.
            #
            # Secondary:
            # fuel/operating cost.
            #
            # Tertiary:
            # total distance.

            comparison = (
                score[2],
                score[0],
                score[1]
            )

            if (
                best_score is None
                or comparison < best_score
            ):

                best_score = comparison
                best_route = route

        return best_route

    # --------------------------------------------------------
    # Larger route:
    # use nearest-neighbour starts + 2-opt.
    # --------------------------------------------------------

    return optimise_large_route_after_first(
        first_customer,
        distances,
        durations,
        fuel_price,
        mpg,
        time_value
    )


# ============================================================
# NEAREST NEIGHBOUR
# ============================================================

def nearest_neighbour_route(
    first_customer,
    distances
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

        next_customer = min(
            remaining,
            key=lambda x:
                distances[current][x]
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
# 2 OPT
# ============================================================

def two_opt(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    time_value
):

    best = route[:]

    (
        best_cost,
        best_distance,
        best_time
    ) = route_total(
        best,
        distances,
        durations,
        fuel_price,
        mpg,
        time_value
    )

    best_score = (
        best_time,
        best_cost,
        best_distance
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

                (
                    candidate_cost,
                    candidate_distance,
                    candidate_time
                ) = route_total(
                    candidate,
                    distances,
                    durations,
                    fuel_price,
                    mpg,
                    time_value
                )

                candidate_score = (
                    candidate_time,
                    candidate_cost,
                    candidate_distance
                )

                if candidate_score < best_score:

                    best = candidate
                    best_score = candidate_score
                    improved = True

                    break

            if improved:
                break

    return best


# ============================================================
# LARGE ROUTE
# ============================================================

def optimise_large_route_after_first(
    first_customer,
    distances,
    durations,
    fuel_price,
    mpg,
    time_value
):

    candidates = []

    # Main nearest-neighbour route.
    candidates.append(
        nearest_neighbour_route(
            first_customer,
            distances
        )
    )

    customer_count = (
        len(distances) - 1
    )

    # --------------------------------------------------------
    # Also try different SECOND stops.
    #
    # This prevents:
    #
    # Grantham
    # -> nearest
    # -> stupid direction
    #
    # and allows the optimiser to choose a sensible
    # geographical progression after the compulsory
    # first stop.
    # --------------------------------------------------------

    second_candidates = sorted(
        [
            x
            for x in range(
                1,
                customer_count + 1
            )
            if x != first_customer
        ],
        key=lambda x:
            distances[first_customer][x]
    )

    # Test up to 12 promising second stops.
    for second in second_candidates[:12]:

        route = [
            0,
            first_customer,
            second
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

        remaining.discard(
            second
        )

        current = second

        while remaining:

            next_customer = min(
                remaining,
                key=lambda x:
                    distances[current][x]
            )

            route.append(
                next_customer
            )

            remaining.remove(
                next_customer
            )

            current = next_customer

        route.append(0)

        candidates.append(
            route
        )

    # --------------------------------------------------------
    # Improve every candidate.
    # --------------------------------------------------------

    best_route = None
    best_score = None

    for candidate in candidates:

        improved = two_opt(
            candidate,
            distances,
            durations,
            fuel_price,
            mpg,
            time_value
        )

        (
            cost,
            distance,
            duration
        ) = route_total(
            improved,
            distances,
            durations,
            fuel_price,
            mpg,
            time_value
        )

        score = (
            duration,
            cost,
            distance
        )

        if (
            best_score is None
            or score < best_score
        ):

            best_score = score
            best_route = improved

    return best_route


# ============================================================
# MAIN ROUTE OPTIMISER
# ============================================================

def optimise_complete_route(
    distances,
    durations,
    fuel_price,
    mpg,
    time_value
):

    # --------------------------------------------------------
    # STEP 1:
    # Find the closest customer to the depot.
    #
    # THIS IS NOW LOCKED.
    # --------------------------------------------------------

    first_customer = (
        find_nearest_first_customer(
            distances,
            durations,
            fuel_price,
            mpg,
            time_value
        )
    )

    # --------------------------------------------------------
    # STEP 2:
    # Optimise everything AFTER the first stop.
    # --------------------------------------------------------

    route = (
        optimise_after_first_exact(
            first_customer,
            distances,
            durations,
            fuel_price,
            mpg,
            time_value
        )
    )

    return (
        route,
        first_customer
    )


# ============================================================
# DESTINATION
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
        row.get("Postcode", "")
    )

    if postcode:
        parts.append(postcode)

    return ", ".join(parts)


# ============================================================
# UPLOAD
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

        required_columns = [
            "Postcode",
            "Price",
            "Phone"
        ]

        missing = [
            column
            for column in required_columns
            if column not in df.columns
        ]

        if missing:

            st.error(
                "Missing required columns: "
                + ", ".join(missing)
            )

            st.stop()

        # Remove empty rows.
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
            subset=[
                "Price"
            ]
        ).copy()

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
# NO DATA
# ============================================================

if "master_df" not in st.session_state:

    st.info(
        "Upload your day's file to begin."
    )

    st.stop()


df = st.session_state.master_df


# ============================================================
# OPTIMISE
# ============================================================

if st.button(
    "🚀 PLAN ROUTE",
    type="primary",
    use_container_width=True
):

    if df.empty:

        st.error(
            "There are no valid jobs."
        )

        st.stop()

    # --------------------------------------------------------
    # DEPOT LOCATION
    # --------------------------------------------------------

    with st.spinner(
        "📍 Locating depot..."
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
    # CUSTOMER GEOCODING
    # --------------------------------------------------------

    progress = st.progress(
        0,
        text="Locating customer addresses..."
    )

    valid_indices = []
    failed_indices = []
    customer_coords = {}

    total_jobs = len(df)

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

            failed_indices.append(
                idx
            )

        else:

            valid_indices.append(
                idx
            )

            customer_coords[idx] = coords

        progress.progress(
            number / total_jobs,
            text=(
                f"Locating customer "
                f"{number} of {total_jobs}"
            )
        )

        # Respect Nominatim rate limit.
        time.sleep(1.0)

    progress.empty()

    # --------------------------------------------------------
    # FAILED ADDRESSES
    # --------------------------------------------------------

    if failed_indices:

        st.session_state.unrouted_df = (
            df.loc[
                failed_indices
            ].copy()
        )

    else:

        st.session_state.unrouted_df = (
            pd.DataFrame()
        )

    if not valid_indices:

        st.error(
            "No customer addresses could be located."
        )

        st.stop()

    # --------------------------------------------------------
    # BUILD ROUTING DATA
    # --------------------------------------------------------

    routing_rows = []

    locations = [
        [
            depot_coords[1],
            depot_coords[0]
        ]
    ]

    # Depot.
    routing_rows.append(
        {
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
    )

    # Customers.
    for idx in valid_indices:

        row = df.loc[idx]

        coords = customer_coords[idx]

        row_dict = row.to_dict()

        row_dict["geo_query"] = (
            build_geo_query(
                row,
                DEPOT_POSTCODE
            )
        )

        row_dict["latitude"] = (
            coords[0]
        )

        row_dict["longitude"] = (
            coords[1]
        )

        routing_rows.append(
            row_dict
        )

        locations.append(
            [
                coords[1],
                coords[0]
            ]
        )

    routing_df = pd.DataFrame(
        routing_rows
    )

    # --------------------------------------------------------
    # MATRIX
    # --------------------------------------------------------

    with st.spinner(
        "🛣️ Calculating actual driving times..."
    ):

        distances, durations = (
            get_ors_matrix(
                locations
            )
        )

    offline = False

    if (
        distances is None
        or durations is None
    ):

        offline = True

        st.warning(
            "OpenRouteService was unavailable. "
            "Using an offline road-distance estimate."
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
        "🧠 Finding the smartest route..."
    ):

        (
            route,
            first_customer
        ) = optimise_complete_route(
            distances,
            durations,
            FUEL_PRICE,
            MPG,
            TIME_VALUE
        )

    # --------------------------------------------------------
    # TOTALS
    # --------------------------------------------------------

    (
        operating_cost,
        total_meters,
        total_seconds
    ) = route_total(
        route,
        distances,
        durations,
        FUEL_PRICE,
        MPG,
        TIME_VALUE
    )

    total_km = (
        total_meters / 1000
    )

    total_miles = (
        total_km * 0.621371
    )

    fuel_litres = (
        total_miles
        / MPG
        * 4.54609
    )

    fuel_cost = (
        fuel_litres
        * FUEL_PRICE
    )

    driving_time_value = (
        total_seconds
        / 3600
        * TIME_VALUE
    )

    revenue = float(
        routing_df["Price"].sum()
    )

    pre_tax_profit = (
        revenue
        - fuel_cost
    )

    take_home_profit = (
        pre_tax_profit
        * (1 - TAX_RATE)
    )

    # --------------------------------------------------------
    # SAVE ROUTE
    # --------------------------------------------------------

    final_df = (
        routing_df
        .iloc[route]
        .copy()
        .reset_index(drop=True)
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

    # --------------------------------------------------------
    # SAVE ECONOMICS
    # --------------------------------------------------------

    st.session_state.route_data = {

        "revenue":
            revenue,

        "fuel_cost":
            fuel_cost,

        "fuel_litres":
            fuel_litres,

        "driving_time_value":
            driving_time_value,

        "operating_cost":
            operating_cost,

        "take_home_profit":
            take_home_profit,

        "total_miles":
            total_miles,

        "total_km":
            total_km,

        "total_seconds":
            total_seconds,

        "first_customer":
            int(first_customer),

        "offline":
            offline
    }

    st.rerun()


# ============================================================
# DASHBOARD
# ============================================================

if "route_data" in st.session_state:

    data = st.session_state.route_data

    st.markdown("---")

    st.subheader(
        "💰 Daily Route Summary"
    )

    c1, c2 = st.columns(2)

    with c1:

        st.metric(
            "Take-Home Profit",
            f"£{data['take_home_profit']:.2f}"
        )

    with c2:

        st.metric(
            "Customer Revenue",
            f"£{data['revenue']:.2f}"
        )

    c3, c4 = st.columns(2)

    with c3:

        st.metric(
            "Total Driving",
            f"{data['total_miles']:.1f} miles"
        )

    with c4:

        st.metric(
            "Driving Time",
            format_duration(
                data["total_seconds"]
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
            f"{data['fuel_litres']:.1f} L"
        )

    st.success(
        "🚗 First stop is locked to the closest "
        "customer to the Grantham depot."
    )

    with st.expander(
        "🧠 How this route was calculated"
    ):

        st.write(
            """
            **Rule 1 — Start at the depot**

            The route starts from your actual depot.

            **Rule 2 — Find the nearest customer**

            Every customer is compared with the depot.
            The first stop is the customer with the shortest
            actual driving time from the depot.

            **Rule 3 — Optimise the rest**

            Once that first stop is locked, the program
            optimises the remaining customers.

            **Rule 4 — Avoid unnecessary backtracking**

            The optimiser considers the entire remaining
            journey rather than simply picking the nearest
            postcode at every step.

            **Rule 5 — Return to Grantham**

            The route is always completed back at the depot.
            """
        )

        if data["offline"]:

            st.warning(
                "The OpenRouteService road matrix was unavailable, "
                "so an estimated route was used."
            )


# ============================================================
# FAILED JOBS
# ============================================================

if (
    "unrouted_df" in st.session_state
    and not st.session_state.unrouted_df.empty
):

    st.markdown("---")

    st.error(
        f"{len(st.session_state.unrouted_df)} "
        "job(s) could not be located."
    )

    st.caption(
        "These jobs were NOT placed at the depot. "
        "Please check their addresses/postcodes."
    )

    with st.expander(
        "View unlocated jobs"
    ):

        st.dataframe(
            st.session_state.unrouted_df,
            use_container_width=True
        )


# ============================================================
# SIDEBAR NAVIGATION
# ============================================================

if (
    "route_data" in st.session_state
    and not st.session_state.master_df.empty
):

    st.sidebar.markdown("---")

    st.sidebar.title(
        "🧭 Route Navigation"
    )

    route_df = (
        st.session_state.master_df
    )

    customer_df = route_df[
        route_df["Status"]
        .astype(str)
        .str.lower()
        != "depot"
    ]

    pending_df = customer_df[
        customer_df["Status"]
        .astype(str)
        .str.lower()
        == "pending"
    ]

    if not pending_df.empty:

        next_row = pending_df.iloc[0]

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
            f"{len(pending_df)} stops remaining"
        )

    else:

        st.sidebar.success(
            "✅ All customer stops completed!"
        )


# ============================================================
# ROUTE DISPLAY
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
    # DEPOT
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
    # RETURN
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

        status_icon = "✅"
        status_text = "Completed"

    else:

        status_icon = "⏳"
        status_text = "Pending"

    # --------------------------------------------------------
    # EXTRA DATA
    # --------------------------------------------------------

    extra_info = []

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

            extra_info.append(
                f"**{column}:** {value}"
            )

    # --------------------------------------------------------
    # CARD
    # --------------------------------------------------------

    with st.container(
        border=True
    ):

        st.write(
            f"### {status_icon} STOP {idx} — {postcode}"
        )

        if extra_info:

            st.write(
                " | ".join(
                    extra_info
                )
            )

        st.write(
            f"**Price:** £{price:.2f}"
            f" | "
            f"**Status:** {status_text}"
            f" | "
            f"**Payment:** {payment}"
        )

        col1, col2 = st.columns(2)

        # ----------------------------------------------------
        # COMPLETE
        # ----------------------------------------------------

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

                    whatsapp_url = (
                        "https://wa.me/"
                        f"{quote(phone)}"
                        "?text="
                        f"{quote(message)}"
                    )

                    st.link_button(
                        "💬 Send WhatsApp",
                        whatsapp_url,
                        use_container_width=True
                    )

        # ----------------------------------------------------
        # NAVIGATION
        # ----------------------------------------------------

        with col2:

            st.link_button(
                "🚗 Navigate Here",
                maps_url(
                    destination
                ),
                key=f"nav_{idx}",
                use_container_width=True
            )

        # ----------------------------------------------------
        # PAYMENT
        # ----------------------------------------------------

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
# EXCEL EXPORT
# ============================================================

st.sidebar.markdown("---")

st.sidebar.title(
    "📊 Export Records"
)

if (
    "master_df" in st.session_state
    and not st.session_state.master_df.empty
):

    export_df = (
        st.session_state.master_df.copy()
    )

    # Remove depot rows.
    if "Status" in export_df.columns:

        export_df = export_df[
            export_df["Status"]
            .astype(str)
            .str.lower()
            != "depot"
        ].copy()

    # Remove technical columns.
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

    # Auto-size columns.
    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = (
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
            column_letter
        ].width = min(
            max_length + 2,
            50
        )

    workbook.save(
        output
    )

    st.sidebar.download_button(
        "⬇️ Download Professional Report",
        output.getvalue(),
        "DanCleanUK_Records.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
