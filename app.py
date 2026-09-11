import streamlit as st
import pandas as pd
import requests
import io
import time
import math
from urllib.parse import quote
from functools import lru_cache

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# ============================================================
# APP VERSION
# ============================================================

APP_VERSION = "5.0"


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="DanCleanUK Optimizer",
    page_icon="🚗",
    layout="centered"
)

st.title("🚗 Daily Route & Profit Optimizer")


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>
    html, body {
        overscroll-behavior-y: none;
    }

    .route-number {
        font-size: 1.25rem;
        font-weight: 700;
    }

    .route-summary {
        padding: 10px;
        border-radius: 8px;
        background-color: #f5f5f5;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SESSION STATE VERSION
# ============================================================

if st.session_state.get("app_version") != APP_VERSION:

    st.session_state.pop("route_data", None)
    st.session_state.pop("master_df", None)
    st.session_state.pop("unrouted_df", None)
    st.session_state.pop("geocode_cache", None)

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
# SIDEBAR SETTINGS
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

# NEW:
# This gives the optimiser a financial value for your time.
TIME_VALUE = st.sidebar.number_input(
    "Value of Driving Time (£/hour)",
    min_value=0.0,
    value=20.0,
    step=1.0,
    help=(
        "Used to balance driving time against fuel cost. "
        "Example: £20 means every hour spent driving "
        "is treated as costing £20."
    )
)


# ============================================================
# API KEY
# ============================================================

if "API_KEY" not in st.secrets:

    st.error(
        "API_KEY missing in Streamlit secrets. "
        "Please add your OpenRouteService API key."
    )

    st.stop()

API_KEY = st.secrets["API_KEY"]


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_val(value):

    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.endswith(".0"):

        try:
            value = str(int(float(value)))
        except ValueError:
            pass

    return value


def format_duration(seconds):

    try:
        seconds = float(seconds)
    except Exception:
        return "0m"

    minutes = max(
        0,
        round(seconds / 60)
    )

    hours = minutes // 60
    mins = minutes % 60

    if hours:
        return f"{hours}h {mins}m"

    return f"{mins}m"


def google_maps_url(destination):

    return (
        "https://www.google.com/maps/dir/"
        "?api=1"
        f"&destination={quote(str(destination))}"
        "&travelmode=driving"
    )


# ============================================================
# BUILD ADDRESS QUERY
# ============================================================

def build_geo_query(row, default_postcode):

    parts = []

    useful_columns = [
        "address",
        "street",
        "location",
        "house",
        "house number",
        "property",
        "property address"
    ]

    for column in row.index:

        if column.lower() in useful_columns:

            value = clean_val(row[column])

            if value:
                parts.append(value)

    postcode = clean_val(
        row.get("Postcode", "")
    )

    if postcode:
        parts.append(postcode)
    else:
        parts.append(default_postcode)

    return ", ".join(parts)


# ============================================================
# GEOCODING
# ============================================================

def get_coords(query_string, postcode_fallback):

    """
    Attempts:

    1. Full address through Nominatim
    2. Postcode through Nominatim
    3. postcodes.io

    Failed addresses return None.

    IMPORTANT:
    We never pretend a failed customer is at the depot.
    """

    query = str(
        query_string
    ).strip()

    postcode = str(
        postcode_fallback
    ).upper().strip()

    cache_key = (
        query
        + "|"
        + postcode
    ).lower()

    if cache_key in st.session_state.geocode_cache:

        return st.session_state.geocode_cache[
            cache_key
        ]

    headers = {
        "User-Agent":
            "DanCleanUKOptimizer/5.0"
    }

    # --------------------------------------------------------
    # FULL ADDRESS
    # --------------------------------------------------------

    if query and query.lower() != "nan":

        full_query = query

        if (
            "united kingdom"
            not in query.lower()
            and not query.lower().endswith("uk")
        ):

            full_query = (
                f"{query}, United Kingdom"
            )

        try:

            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": full_query,
                    "format": "json",
                    "limit": 1,
                    "countrycodes": "gb"
                },
                headers=headers,
                timeout=6
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
    # POSTCODE NOMINATIM
    # --------------------------------------------------------

    if postcode:

        try:

            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q":
                        f"{postcode}, United Kingdom",
                    "format": "json",
                    "limit": 1,
                    "countrycodes": "gb"
                },
                headers=headers,
                timeout=6
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
    # POSTCODES.IO
    # --------------------------------------------------------

    if postcode:

        postcode_clean = (
            postcode
            .replace(" ", "")
        )

        try:

            response = requests.get(
                "https://api.postcodes.io/postcodes/"
                f"{postcode_clean}",
                timeout=5
            )

            if response.status_code == 200:

                result = (
                    response.json()
                    .get("result")
                )

                if result:

                    lat = result.get(
                        "latitude"
                    )

                    lon = result.get(
                        "longitude"
                    )

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
        math.cos(
            math.radians(lat1)
        )
        *
        math.cos(
            math.radians(lat2)
        )
        *
        math.sin(dlon / 2) ** 2
    )

    c = (
        2
        *
        math.asin(
            math.sqrt(a)
        )
    )

    return R * c


# ============================================================
# OFFLINE MATRIX
# ============================================================

def build_offline_matrix(locations):

    count = len(locations)

    distance_matrix = [
        [0.0 for _ in range(count)]
        for _ in range(count)
    ]

    duration_matrix = [
        [0.0 for _ in range(count)]
        for _ in range(count)
    ]

    # Approximate road distance from straight-line distance.
    ROAD_FACTOR = 1.30

    # Conservative average speed.
    AVG_SPEED_KMH = 40.0

    for i in range(count):

        lon1, lat1 = locations[i]

        for j in range(count):

            if i == j:
                continue

            lon2, lat2 = locations[j]

            straight_km = haversine_km(
                lat1,
                lon1,
                lat2,
                lon2
            )

            road_km = (
                straight_km
                * ROAD_FACTOR
            )

            distance_matrix[i][j] = (
                road_km * 1000
            )

            duration_matrix[i][j] = (
                road_km
                / AVG_SPEED_KMH
                * 3600
            )

    return (
        distance_matrix,
        duration_matrix
    )


# ============================================================
# OPENROUTESERVICE MATRIX
# ============================================================

def get_route_matrix(locations):

    """
    Returns:

        distance matrix in metres
        duration matrix in seconds
    """

    try:

        body = {
            "locations": locations,
            "metrics": [
                "distance",
                "duration"
            ],
            "units": "m"
        }

        response = requests.post(
            "https://api.openrouteservice.org/v2/matrix/driving-car",
            json=body,
            headers={
                "Authorization": API_KEY,
                "Content-Type": "application/json"
            },
            timeout=30
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

        if (
            not distances
            or not durations
        ):

            return None, None

        return (
            distances,
            durations
        )

    except Exception:

        return None, None


# ============================================================
# ROUTE OBJECTIVE
# ============================================================

def route_totals(
    route,
    distance_matrix,
    duration_matrix,
    fuel_price,
    mpg,
    time_value
):

    total_distance = 0.0
    total_duration = 0.0

    for i in range(
        len(route) - 1
    ):

        a = route[i]
        b = route[i + 1]

        total_distance += float(
            distance_matrix[a][b]
        )

        total_duration += float(
            duration_matrix[a][b]
        )

    miles = (
        total_distance
        / 1000
        * 0.621371
    )

    fuel_litres = (
        miles
        / mpg
        * 4.54609
    )

    fuel_cost = (
        fuel_litres
        * fuel_price
    )

    driving_time_cost = (
        total_duration
        / 3600
        * time_value
    )

    total_operating_cost = (
        fuel_cost
        + driving_time_cost
    )

    return {
        "meters":
            total_distance,

        "seconds":
            total_duration,

        "miles":
            miles,

        "fuel_litres":
            fuel_litres,

        "fuel_cost":
            fuel_cost,

        "driving_time_cost":
            driving_time_cost,

        "operating_cost":
            total_operating_cost
    }


def route_score(
    route,
    distance_matrix,
    duration_matrix,
    fuel_price,
    mpg,
    time_value
):

    totals = route_totals(
        route,
        distance_matrix,
        duration_matrix,
        fuel_price,
        mpg,
        time_value
    )

    return totals["operating_cost"]


# ============================================================
# NEAREST NEIGHBOUR INITIAL ROUTE
# ============================================================

def nearest_neighbour_route(
    start_customer,
    distance_matrix
):

    count = len(
        distance_matrix
    )

    customers = set(
        range(1, count)
    )

    route = [0]

    current = 0

    # Optionally force a particular first customer.
    if (
        start_customer is not None
        and start_customer in customers
    ):

        route.append(
            start_customer
        )

        customers.remove(
            start_customer
        )

        current = start_customer

    while customers:

        next_node = min(
            customers,
            key=lambda x:
                distance_matrix[current][x]
        )

        route.append(
            next_node
        )

        customers.remove(
            next_node
        )

        current = next_node

    route.append(0)

    return route


# ============================================================
# CHEAPEST INSERTION
# ============================================================

def cheapest_insertion_route(
    distance_matrix,
    duration_matrix,
    fuel_price,
    mpg,
    time_value
):

    """
    Builds a sensible initial route by repeatedly inserting
    the customer where it adds the least total route cost.
    """

    count = len(
        distance_matrix
    )

    if count <= 1:
        return [0, 0]

    customers = list(
        range(1, count)
    )

    # Start with customer having the cheapest
    # round trip from depot.
    first = min(
        customers,
        key=lambda x:
            distance_matrix[0][x]
            +
            distance_matrix[x][0]
    )

    route = [
        0,
        first,
        0
    ]

    customers.remove(first)

    while customers:

        best_customer = None
        best_position = None
        best_increase = float("inf")

        for customer in customers:

            for position in range(
                1,
                len(route)
            ):

                a = route[position - 1]
                b = route[position]

                old_route = (
                    distance_matrix[a][b]
                )

                new_route = (
                    distance_matrix[a][customer]
                    +
                    distance_matrix[customer][b]
                )

                distance_increase = (
                    new_route
                    - old_route
                )

                old_time = (
                    duration_matrix[a][b]
                )

                new_time = (
                    duration_matrix[a][customer]
                    +
                    duration_matrix[customer][b]
                )

                time_increase = (
                    new_time
                    - old_time
                )

                # Convert distance to fuel cost.
                miles_added = (
                    distance_increase
                    / 1000
                    * 0.621371
                )

                litres_added = (
                    miles_added
                    / mpg
                    * 4.54609
                )

                fuel_added = (
                    litres_added
                    * fuel_price
                )

                time_added = (
                    time_increase
                    / 3600
                    * time_value
                )

                increase = (
                    fuel_added
                    + time_added
                )

                if increase < best_increase:

                    best_increase = increase
                    best_customer = customer
                    best_position = position

        route.insert(
            best_position,
            best_customer
        )

        customers.remove(
            best_customer
        )

    return route


# ============================================================
# 2-OPT
# ============================================================

def two_opt(
    route,
    distance_matrix,
    duration_matrix,
    fuel_price,
    mpg,
    time_value
):

    """
    Improves the COMPLETE route.

    Unlike the previous version, this is intentional:
    we are NOT preserving nearest-neighbour order.

    We are trying to minimise the whole journey including
    the return to Grantham.
    """

    best_route = route[:]

    best_score = route_score(
        best_route,
        distance_matrix,
        duration_matrix,
        fuel_price,
        mpg,
        time_value
    )

    improved = True

    while improved:

        improved = False

        # Don't reverse depot.
        for i in range(
            1,
            len(best_route) - 2
        ):

            for j in range(
                i + 1,
                len(best_route) - 1
            ):

                candidate = (
                    best_route[:i]
                    +
                    best_route[i:j + 1][::-1]
                    +
                    best_route[j + 1:]
                )

                candidate_score = route_score(
                    candidate,
                    distance_matrix,
                    duration_matrix,
                    fuel_price,
                    mpg,
                    time_value
                )

                if (
                    candidate_score
                    <
                    best_score - 0.0001
                ):

                    best_route = candidate
                    best_score = candidate_score
                    improved = True

                    break

            if improved:
                break

    return best_route


# ============================================================
# EXACT OPTIMISATION FOR SMALL ROUTES
# ============================================================

def exact_route_dynamic_programming(
    distance_matrix,
    duration_matrix,
    fuel_price,
    mpg,
    time_value
):

    """
    Exact travelling-salesman style optimisation.

    Used for smaller days.

    State:
        (visited_customers, current_customer)

    It evaluates every possible customer sequence
    without blindly following nearest-neighbour.

    This guarantees the best route according to the
    operating-cost objective for the given matrix.
    """

    n = len(
        distance_matrix
    ) - 1

    if n == 0:
        return [0, 0]

    # Customers are represented by bits.
    full_mask = (
        (1 << n) - 1
    )

    # dp[(mask, current)] = (cost, previous)
    dp = {}

    # Start from depot.
    for customer_index in range(n):

        node = customer_index + 1

        route = [
            0,
            node
        ]

        cost = route_score(
            route,
            distance_matrix,
            duration_matrix,
            fuel_price,
            mpg,
            time_value
        )

        # We don't actually want the incomplete route's
        # return-to-depot cost here.
        # So calculate only the first leg.
        first_leg = route_totals(
            [0, node],
            distance_matrix,
            duration_matrix,
            fuel_price,
            mpg,
            time_value
        )["operating_cost"]

        mask = (
            1 << customer_index
        )

        dp[
            (mask, node)
        ] = (
            first_leg,
            0
        )

    # --------------------------------------------------------
    # BUILD STATES
    # --------------------------------------------------------

    for mask in range(
        1,
        full_mask + 1
    ):

        for current_customer in range(
            1,
            n + 1
        ):

            state = (
                mask,
                current_customer
            )

            if state not in dp:
                continue

            current_cost, _ = dp[state]

            for next_customer in range(
                1,
                n + 1
            ):

                bit = (
                    1
                    <<
                    (next_customer - 1)
                )

                if mask & bit:
                    continue

                new_mask = (
                    mask | bit
                )

                a = current_customer
                b = next_customer

                # Incremental operating cost.
                distance_m = (
                    distance_matrix[a][b]
                )

                duration_s = (
                    duration_matrix[a][b]
                )

                miles = (
                    distance_m
                    / 1000
                    * 0.621371
                )

                fuel_litres = (
                    miles
                    / mpg
                    * 4.54609
                )

                fuel_cost = (
                    fuel_litres
                    * fuel_price
                )

                time_cost = (
                    duration_s
                    / 3600
                    * time_value
                )

                increment = (
                    fuel_cost
                    + time_cost
                )

                new_cost = (
                    current_cost
                    + increment
                )

                new_state = (
                    new_mask,
                    next_customer
                )

                if (
                    new_state not in dp
                    or
                    new_cost
                    <
                    dp[new_state][0]
                ):

                    dp[new_state] = (
                        new_cost,
                        current_customer
                    )

    # --------------------------------------------------------
    # CLOSE ROUTE BACK TO DEPOT
    # --------------------------------------------------------

    best_cost = float("inf")
    best_last = None

    for current in range(
        1,
        n + 1
    ):

        state = (
            full_mask,
            current
        )

        if state not in dp:
            continue

        current_cost = (
            dp[state][0]
        )

        distance_m = (
            distance_matrix[current][0]
        )

        duration_s = (
            duration_matrix[current][0]
        )

        miles = (
            distance_m
            / 1000
            * 0.621371
        )

        fuel_litres = (
            miles
            / mpg
            * 4.54609
        )

        fuel_cost = (
            fuel_litres
            * fuel_price
        )

        time_cost = (
            duration_s
            / 3600
            * time_value
        )

        final_cost = (
            current_cost
            + fuel_cost
            + time_cost
        )

        if final_cost < best_cost:

            best_cost = final_cost
            best_last = current

    # --------------------------------------------------------
    # RECONSTRUCT ROUTE
    # --------------------------------------------------------

    route_reversed = []

    mask = full_mask
    current = best_last

    while current != 0:

        route_reversed.append(
            current
        )

        previous = dp[
            (mask, current)
        ][1]

        mask ^= (
            1
            <<
            (current - 1)
        )

        current = previous

    route = [
        0
    ] + list(
        reversed(
            route_reversed
        )
    ) + [
        0
    ]

    return route


# ============================================================
# MAIN OPTIMISER
# ============================================================

def optimise_complete_route(
    distance_matrix,
    duration_matrix,
    fuel_price,
    mpg,
    time_value
):

    customer_count = (
        len(distance_matrix) - 1
    )

    # --------------------------------------------------------
    # SMALL ROUTES
    #
    # Exact optimisation.
    #
    # 11 customers is already computationally significant,
    # so exact mode is limited to 10.
    # --------------------------------------------------------

    if customer_count <= 10:

        return exact_route_dynamic_programming(
            distance_matrix,
            duration_matrix,
            fuel_price,
            mpg,
            time_value
        )

    # --------------------------------------------------------
    # LARGER ROUTES
    #
    # Generate several different starting routes.
    # Improve each with 2-opt.
    # Keep the best COMPLETE route.
    # --------------------------------------------------------

    candidates = []

    # --------------------------------------------------------
    # 1. Nearest-neighbour from every customer
    # --------------------------------------------------------

    for first_customer in range(
        1,
        customer_count + 1
    ):

        route = nearest_neighbour_route(
            first_customer,
            distance_matrix
        )

        candidates.append(
            route
        )

    # --------------------------------------------------------
    # 2. Cheapest insertion
    # --------------------------------------------------------

    candidates.append(
        cheapest_insertion_route(
            distance_matrix,
            duration_matrix,
            fuel_price,
            mpg,
            time_value
        )
    )

    # --------------------------------------------------------
    # IMPROVE ALL CANDIDATES
    # --------------------------------------------------------

    best_route = None
    best_score = float("inf")

    for route in candidates:

        improved = two_opt(
            route,
            distance_matrix,
            duration_matrix,
            fuel_price,
            mpg,
            time_value
        )

        score = route_score(
            improved,
            distance_matrix,
            duration_matrix,
            fuel_price,
            mpg,
            time_value
        )

        if score < best_score:

            best_score = score
            best_route = improved

    return best_route


# ============================================================
# MAP DESTINATION
# ============================================================

def get_map_destination(row):

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
                row[column]
            )

            if value:
                parts.append(value)

    postcode = clean_val(
        row.get(
            "Postcode",
            ""
        )
    )

    if postcode:
        parts.append(postcode)

    return ", ".join(parts)


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "📁 Upload your day's file",
    type=["csv", "xlsx"],
    help="Required columns: Postcode, Price, Phone"
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

        # ----------------------------------------------------
        # REQUIRED COLUMNS
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # CLEAN EMPTY ROWS
        # ----------------------------------------------------

        df = (
            df
            .dropna(how="all")
            .copy()
        )

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
            ~df["Postcode"]
            .isin([
                "",
                "NAN",
                "NAT"
            ])
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
            .isin([
                "",
                "nan",
                "nat"
            ])
        ].copy()

        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------

        df["Status"] = "pending"
        df["Payment"] = "waiting"

        df = (
            df
            .reset_index(drop=True)
        )

        st.session_state.master_df = df

        st.rerun()

    except Exception as e:

        st.error(
            f"Error loading file: {e}"
        )


# ============================================================
# MAIN APPLICATION
# ============================================================

if "master_df" in st.session_state:

    df = st.session_state.master_df


    # ========================================================
    # PLAN ROUTE
    # ========================================================

    if st.button(
        "🚀 PLAN MOST EFFICIENT ROUTE",
        type="primary",
        use_container_width=True
    ):

        if df.empty:

            st.error(
                "There are no customer jobs."
            )

            st.stop()

        # ----------------------------------------------------
        # LOCATE DEPOT
        # ----------------------------------------------------

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

        # ----------------------------------------------------
        # LOCATE CUSTOMERS
        # ----------------------------------------------------

        valid_rows = []
        failed_rows = []

        customer_coordinates = {}

        progress = st.progress(
            0,
            text="Locating customer addresses..."
        )

        total_customers = len(df)

        for position, (original_index, row) in enumerate(
            df.iterrows()
        ):

            query = build_geo_query(
                row,
                DEPOT_POSTCODE
            )

            postcode = clean_val(
                row["Postcode"]
            )

            coords = get_coords(
                query,
                postcode
            )

            if coords is None:

                failed_rows.append(
                    {
                        "index":
                            original_index,

                        "postcode":
                            postcode,

                        "query":
                            query
                    }
                )

            else:

                valid_rows.append(
                    original_index
                )

                customer_coordinates[
                    original_index
                ] = coords

            progress.progress(
                (position + 1)
                / total_customers,
                text=(
                    f"Locating address "
                    f"{position + 1} of "
                    f"{total_customers}"
                )
            )

            # Respect public geocoder.
            time.sleep(1.0)

        progress.empty()

        # ----------------------------------------------------
        # FAILED ADDRESSES
        # ----------------------------------------------------

        if failed_rows:

            failed_df = df.loc[
                [
                    item["index"]
                    for item in failed_rows
                ]
            ].copy()

            failed_df["Geo Error"] = [
                item["query"]
                for item in failed_rows
            ]

            st.session_state.unrouted_df = (
                failed_df
            )

        else:

            st.session_state.unrouted_df = (
                pd.DataFrame()
            )

        # ----------------------------------------------------
        # CHECK
        # ----------------------------------------------------

        if not valid_rows:

            st.error(
                "No customer addresses could be located."
            )

            st.stop()

        # ----------------------------------------------------
        # BUILD ROUTING DATA
        # ----------------------------------------------------

        routing_rows = []

        coordinates = [
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
        for original_index in valid_rows:

            original_row = (
                df.loc[
                    original_index
                ].copy()
            )

            coords = (
                customer_coordinates[
                    original_index
                ]
            )

            query = build_geo_query(
                original_row,
                DEPOT_POSTCODE
            )

            row_dict = (
                original_row.to_dict()
            )

            row_dict["geo_query"] = (
                query
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

            coordinates.append(
                [
                    coords[1],
                    coords[0]
                ]
            )

        routing_df = pd.DataFrame(
            routing_rows
        )

        # ----------------------------------------------------
        # ROUTING MATRIX
        # ----------------------------------------------------

        with st.spinner(
            "🛣️ Calculating real road distances and driving times..."
        ):

            (
                distance_matrix,
                duration_matrix
            ) = get_route_matrix(
                coordinates
            )

        using_offline = False

        if (
            distance_matrix is None
            or duration_matrix is None
        ):

            using_offline = True

            st.warning(
                "⚠️ Live road routing was unavailable. "
                "Using the offline fallback."
            )

            (
                distance_matrix,
                duration_matrix
            ) = build_offline_matrix(
                coordinates
            )

        # ----------------------------------------------------
        # OPTIMISE COMPLETE ROUTE
        # ----------------------------------------------------

        customer_count = (
            len(routing_df) - 1
        )

        if customer_count <= 10:

            optimisation_text = (
                "🔬 Searching the complete set of "
                "possible routes..."
            )

        else:

            optimisation_text = (
                "🧠 Testing multiple route patterns "
                "and improving the complete journey..."
            )

        with st.spinner(
            optimisation_text
        ):

            route_indices = (
                optimise_complete_route(
                    distance_matrix,
                    duration_matrix,
                    FUEL_PRICE,
                    MPG,
                    TIME_VALUE
                )
            )

        # ----------------------------------------------------
        # TOTALS
        # ----------------------------------------------------

        totals = route_totals(
            route_indices,
            distance_matrix,
            duration_matrix,
            FUEL_PRICE,
            MPG,
            TIME_VALUE
        )

        total_miles = totals[
            "miles"
        ]

        total_seconds = totals[
            "seconds"
        ]

        fuel_litres = totals[
            "fuel_litres"
        ]

        fuel_cost = totals[
            "fuel_cost"
        ]

        driving_time_cost = totals[
            "driving_time_cost"
        ]

        operating_cost = totals[
            "operating_cost"
        ]

        # ----------------------------------------------------
        # REVENUE
        # ----------------------------------------------------

        revenue = (
            routing_df["Price"]
            .sum()
        )

        pre_tax_profit = (
            revenue
            - fuel_cost
        )

        take_home_profit = (
            pre_tax_profit
            *
            (1 - TAX_RATE)
        )

        # ----------------------------------------------------
        # REORDER DATA
        # ----------------------------------------------------

        route_df = (
            routing_df
            .iloc[
                route_indices
            ]
            .copy()
            .reset_index(drop=True)
        )

        route_df["route_index"] = (
            range(
                len(route_df)
            )
        )

        # ----------------------------------------------------
        # DEPOT STATUS
        # ----------------------------------------------------

        route_df.loc[
            0,
            "Status"
        ] = "depot"

        route_df.loc[
            len(route_df) - 1,
            "Status"
        ] = "depot"

        # ----------------------------------------------------
        # SAVE
        # ----------------------------------------------------

        st.session_state.master_df = (
            route_df
        )

        st.session_state.route_data = {

            "total_miles":
                float(total_miles),

            "total_km":
                float(
                    total_miles
                    / 0.621371
                ),

            "total_seconds":
                float(total_seconds),

            "fuel_litres":
                float(fuel_litres),

            "fuel_cost":
                float(fuel_cost),

            "driving_time_cost":
                float(driving_time_cost),

            "operating_cost":
                float(operating_cost),

            "revenue":
                float(revenue),

            "pre_tax_profit":
                float(pre_tax_profit),

            "take_home_profit":
                float(take_home_profit),

            "customer_count":
                int(customer_count),

            "time_value":
                float(TIME_VALUE),

            "offline":
                bool(using_offline)
        }

        st.rerun()


    # ========================================================
    # ECONOMICS DASHBOARD
    # ========================================================

    if "route_data" in st.session_state:

        data = st.session_state.get(
            "route_data",
            {}
        )

        take_home_profit = float(
            data.get(
                "take_home_profit",
                0
            )
        )

        revenue = float(
            data.get(
                "revenue",
                0
            )
        )

        total_miles = float(
            data.get(
                "total_miles",
                0
            )
        )

        total_seconds = float(
            data.get(
                "total_seconds",
                0
            )
        )

        fuel_cost = float(
            data.get(
                "fuel_cost",
                0
            )
        )

        fuel_litres = float(
            data.get(
                "fuel_litres",
                0
            )
        )

        driving_time_cost = float(
            data.get(
                "driving_time_cost",
                0
            )
        )

        operating_cost = float(
            data.get(
                "operating_cost",
                0
            )
        )

        customer_count = int(
            data.get(
                "customer_count",
                0
            )
        )

        time_value = float(
            data.get(
                "time_value",
                TIME_VALUE
            )
        )

        offline = bool(
            data.get(
                "offline",
                False
            )
        )

        # ----------------------------------------------------
        # HEADER
        # ----------------------------------------------------

        st.markdown("---")

        st.subheader(
            "💰 Today's Route Economics"
        )

        # ----------------------------------------------------
        # PRIMARY FIGURES
        # ----------------------------------------------------

        col1, col2 = st.columns(2)

        with col1:

            st.metric(
                "Take-Home Profit",
                f"£{take_home_profit:.2f}"
            )

        with col2:

            st.metric(
                "Customer Revenue",
                f"£{revenue:.2f}"
            )

        col3, col4 = st.columns(2)

        with col3:

            st.metric(
                "Total Driving",
                f"{total_miles:.1f} miles"
            )

        with col4:

            st.metric(
                "Driving Time",
                format_duration(
                    total_seconds
                )
            )

        col5, col6 = st.columns(2)

        with col5:

            st.metric(
                "Fuel Cost",
                f"£{fuel_cost:.2f}"
            )

        with col6:

            st.metric(
                "Fuel Used",
                f"{fuel_litres:.1f} L"
            )

        # ----------------------------------------------------
        # OPTIMISATION INFORMATION
        # ----------------------------------------------------

        with st.expander(
            "💡 How the route was optimised"
        ):

            st.write(
                f"""
                The optimiser considers the **complete journey**,
                including the return to the Grantham depot.

                **Value of your driving time:** £{time_value:.2f}/hour

                **Fuel cost:** £{fuel_cost:.2f}

                **Time cost:** £{driving_time_cost:.2f}

                **Combined operating cost used by optimiser:**
                £{operating_cost:.2f}

                The route is therefore not simply:
                "go to whichever address is nearest."

                It asks:
                **"Which complete sequence of jobs gives the
                lowest overall driving cost?"**
                """
            )

            if customer_count <= 10:

                st.success(
                    "🔬 This route used exact optimisation "
                    "for the available customer sequence."
                )

            else:

                st.info(
                    "🧠 This route used multiple starting "
                    "patterns followed by route improvement."
                )

        if offline:

            st.warning(
                "⚠️ Live road routing was unavailable. "
                "Distance and time are estimates."
            )


    # ========================================================
    # UNROUTED JOBS
    # ========================================================

    if (
        "unrouted_df" in st.session_state
        and not st.session_state.unrouted_df.empty
    ):

        st.markdown("---")

        st.error(
            f"⚠️ "
            f"{len(st.session_state.unrouted_df)} "
            "customer(s) could not be located."
        )

        with st.expander(
            "View addresses needing checking"
        ):

            for _, row in (
                st.session_state
                .unrouted_df
                .iterrows()
            ):

                postcode = clean_val(
                    row.get(
                        "Postcode"
                    )
                )

                geo_error = clean_val(
                    row.get(
                        "Geo Error"
                    )
                )

                st.write(
                    f"**{postcode}** — {geo_error}"
                )

            st.caption(
                "These jobs were NOT included in the "
                "optimised route and were NOT placed "
                "at the depot."
            )


    # ========================================================
    # SIDEBAR NAVIGATION
    # ========================================================

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

            next_row = (
                pending_df.iloc[0]
            )

            destination = (
                get_map_destination(
                    next_row
                )
            )

            nav_url = (
                google_maps_url(
                    destination
                )
            )

            st.sidebar.link_button(
                "🚗 Navigate to Next Stop",
                nav_url,
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


    # ========================================================
    # ROUTE DISPLAY
    # ========================================================

    st.markdown("---")

    st.subheader(
        "📍 Optimised Route"
    )

    for idx, row in (
        st.session_state.master_df
        .iterrows()
    ):

        # ----------------------------------------------------
        # START DEPOT
        # ----------------------------------------------------

        if idx == 0:

            with st.container(
                border=True
            ):

                st.write(
                    "### 🏠 START — GRANtham DEPOT"
                )

                st.write(
                    DEPOT_FULL_ADDRESS
                )

            continue

        # ----------------------------------------------------
        # RETURN DEPOT
        # ----------------------------------------------------

        if (
            idx
            ==
            len(
                st.session_state.master_df
            ) - 1
        ):

            return_url = (
                google_maps_url(
                    DEPOT_FULL_ADDRESS
                )
            )

            with st.container(
                border=True
            ):

                st.write(
                    "### 🏁 FINISH — DEPOT"
                )

                st.write(
                    DEPOT_FULL_ADDRESS
                )

                st.link_button(
                    "🚗 Navigate Back to Depot",
                    return_url,
                    key=f"return_{idx}",
                    use_container_width=True
                )

            continue

        # ----------------------------------------------------
        # CUSTOMER
        # ----------------------------------------------------

        postcode = clean_val(
            row.get(
                "Postcode"
            )
        )

        price = float(
            row.get(
                "Price",
                0
            )
        )

        phone = clean_val(
            row.get(
                "Phone"
            )
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
            get_map_destination(
                row
            )
        )

        map_url = (
            google_maps_url(
                destination
            )
        )

        # ----------------------------------------------------
        # EXTRA DATA
        # ----------------------------------------------------

        extra_parts = []

        for column in (
            st.session_state
            .master_df
            .columns
        ):

            if column.lower() in [
                "postcode",
                "price",
                "phone",
                "status",
                "payment",
                "latitude",
                "longitude",
                "geo_query",
                "route_index"
            ]:
                continue

            if (
                "date"
                in column.lower()
                or
                "time"
                in column.lower()
            ):
                continue

            value = clean_val(
                row.get(
                    column
                )
            )

            if value:

                extra_parts.append(
                    f"**{column}:** {value}"
                )

        # ----------------------------------------------------
        # STATUS
        # ----------------------------------------------------

        if status == "completed":

            status_icon = "✅"
            status_text = "Completed"

        else:

            status_icon = "⏳"
            status_text = "Pending"

        # ----------------------------------------------------
        # CUSTOMER CARD
        # ----------------------------------------------------

        with st.container(
            border=True
        ):

            st.write(
                f"### {status_icon} "
                f"STOP {idx} — {postcode}"
            )

            if extra_parts:

                st.write(
                    " | ".join(
                        extra_parts
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

            # ------------------------------------------------
            # COMPLETE
            # ------------------------------------------------

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

            # ------------------------------------------------
            # NAVIGATE
            # ------------------------------------------------

            with col2:

                st.link_button(
                    "🚗 Navigate Here",
                    map_url,
                    key=f"navigate_{idx}",
                    use_container_width=True
                )

            # ------------------------------------------------
            # PAYMENT
            # ------------------------------------------------

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

st.sidebar.title(
    "📊 Export Monthly Records"
)

if (
    "master_df" in st.session_state
    and not st.session_state.master_df.empty
):

    export_df = (
        st.session_state.master_df.copy()
    )

    # Remove depot.
    if "Status" in export_df.columns:

        export_df = export_df[
            export_df["Status"]
            .astype(str)
            .str.lower()
            != "depot"
        ].copy()

    # Remove technical fields.
    for column in [
        "latitude",
        "longitude",
        "geo_query",
        "route_index"
    ]:

        if column in export_df.columns:

            export_df = (
                export_df
                .drop(
                    columns=[column]
                )
            )

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

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

    # Auto-size.
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

else:

    st.sidebar.info(
        "Upload a file to begin."
    )
