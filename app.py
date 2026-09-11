import streamlit as st
import pandas as pd
import requests
import io
import time
import math
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows


# ============================================================
# APP VERSION
# ============================================================

APP_VERSION = "4.0"


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
# MOBILE / UI CSS
# ============================================================

st.markdown(
    """
    <style>
    html, body {
        overscroll-behavior-y: none;
    }

    .route-stop {
        font-size: 1.15rem;
        font-weight: 700;
    }

    .small-muted {
        color: #777;
        font-size: 0.85rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SESSION STATE VERSION CONTROL
# ============================================================

if st.session_state.get("app_version") != APP_VERSION:

    # Old versions may contain incompatible route_data.
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
    """
    Safely convert Excel/CSV values into clean strings.
    """

    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.endswith(".0"):

        try:
            value = str(
                int(float(value))
            )
        except ValueError:
            pass

    return value


def format_duration(seconds):
    """
    Convert seconds into a readable duration.
    """

    if seconds is None:
        return "0m"

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

    if hours > 0:
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
# BUILD BEST GEOCODING QUERY
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

    return ", ".join(parts)


# ============================================================
# GEOCODER
# ============================================================

def get_coords(query_string, postcode_fallback):

    """
    Multi-level geocoder.

    IMPORTANT:
    Failed addresses return None.

    We NEVER pretend that an unknown customer
    is located at the depot.
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

    # --------------------------------------------------------
    # CACHE
    # --------------------------------------------------------

    if cache_key in st.session_state.geocode_cache:

        return st.session_state.geocode_cache[
            cache_key
        ]

    headers = {
        "User-Agent":
            "DanCleanUKOptimizer/4.0"
    }

    # --------------------------------------------------------
    # 1. FULL ADDRESS VIA NOMINATIM
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
                timeout=5
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
    # 2. POSTCODE VIA NOMINATIM
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
                timeout=5
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
    # 3. POSTCODES.IO
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
                timeout=4
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

    # --------------------------------------------------------
    # FAILED
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
# OFFLINE ROUTING FALLBACK
# ============================================================

def build_offline_matrix(locations):

    """
    Offline fallback.

    locations format:
        [longitude, latitude]

    This is less accurate than real road routing,
    but is better than failing completely.
    """

    count = len(locations)

    distance_matrix = [
        [0.0 for _ in range(count)]
        for _ in range(count)
    ]

    duration_matrix = [
        [0.0 for _ in range(count)]
        for _ in range(count)
    ]

    # Approximate road multiplier.
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
    Gets REAL road distance and driving duration
    from OpenRouteService.

    distances:
        metres

    durations:
        seconds
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
            timeout=20
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

        return (
            distances,
            durations
        )

    except Exception:
        return None, None


# ============================================================
# NEAREST-FIRST ROUTE
# ============================================================

def build_nearest_first_route(
    distance_matrix,
    duration_matrix
):

    """
    MAIN ROUTING RULE.

    Starting at the depot:

        1. Find every unvisited customer.
        2. Choose the customer with the shortest
           REAL driving distance.
        3. If distances are effectively tied,
           use driving time as the tie-breaker.
        4. Repeat from the new location.

    IMPORTANT:

    There is NO 2-opt afterwards.

    Therefore the route produced here remains
    nearest-first all the way through.
    """

    count = len(
        distance_matrix
    )

    unvisited = set(
        range(1, count)
    )

    route = [0]

    current = 0

    while unvisited:

        candidates = []

        for candidate in unvisited:

            distance = float(
                distance_matrix[
                    current
                ][candidate]
            )

            duration = float(
                duration_matrix[
                    current
                ][candidate]
            )

            candidates.append(
                (
                    distance,
                    duration,
                    candidate
                )
            )

        # ----------------------------------------------------
        # SORT BY:
        #
        # 1. REAL DRIVING DISTANCE
        # 2. REAL DRIVING TIME
        #
        # This guarantees nearest address first.
        # ----------------------------------------------------

        candidates.sort(
            key=lambda item: (
                item[0],
                item[1]
            )
        )

        next_customer = (
            candidates[0][2]
        )

        route.append(
            next_customer
        )

        unvisited.remove(
            next_customer
        )

        current = next_customer

    # Return to depot.
    route.append(0)

    return route


# ============================================================
# ROUTE TOTALS
# ============================================================

def calculate_route_totals(
    route,
    distance_matrix,
    duration_matrix
):

    total_distance = 0.0
    total_duration = 0.0

    for i in range(
        len(route) - 1
    ):

        start = route[i]
        end = route[i + 1]

        total_distance += float(
            distance_matrix[
                start
            ][end]
        )

        total_duration += float(
            duration_matrix[
                start
            ][end]
        )

    return (
        total_distance,
        total_duration
    )


# ============================================================
# DESTINATION FOR MAPS
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

        # ----------------------------------------------------
        # LOAD FILE
        # ----------------------------------------------------

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

        # Clean column names.
        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        # ----------------------------------------------------
        # VALIDATE REQUIRED COLUMNS
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
        # REMOVE COMPLETELY EMPTY ROWS
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
            (
                df["Postcode"] != ""
            )
            &
            (
                df["Postcode"] != "NAN"
            )
            &
            (
                df["Postcode"] != "NAT"
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
        # INITIAL STATUS
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
# MAIN APP
# ============================================================

if "master_df" in st.session_state:

    df = st.session_state.master_df


    # ========================================================
    # PLAN ROUTE BUTTON
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
        # DEPOT LOCATION
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
                "Could not locate the depot. "
                "Please check the depot address/postcode."
            )

            st.stop()

        # ----------------------------------------------------
        # CUSTOMER GEOCODING
        # ----------------------------------------------------

        valid_rows = []
        failed_rows = []

        coordinates = [
            [
                depot_coords[1],
                depot_coords[0]
            ]
        ]

        geo_queries = []

        with st.spinner(
            "📍 Finding every customer address..."
        ):

            for original_index, row in df.iterrows():

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

                    geo_queries.append(
                        query
                    )

                    coordinates.append(
                        [
                            coords[1],
                            coords[0]
                        ]
                    )

                # Polite geocoder delay.
                time.sleep(0.10)

        # ----------------------------------------------------
        # FAILED JOBS
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
        # NO VALID CUSTOMERS
        # ----------------------------------------------------

        if not valid_rows:

            st.error(
                "None of the customer addresses "
                "could be located."
            )

            st.stop()

        # ----------------------------------------------------
        # CREATE ROUTING DATAFRAME
        # ----------------------------------------------------

        routing_rows = []

        # Depot row.
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

        # Customer rows.
        for original_index in valid_rows:

            original_row = (
                df.loc[
                    original_index
                ].copy()
            )

            query = build_geo_query(
                original_row,
                DEPOT_POSTCODE
            )

            postcode = clean_val(
                original_row["Postcode"]
            )

            coords = get_coords(
                query,
                postcode
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

        routing_df = pd.DataFrame(
            routing_rows
        )

        # ----------------------------------------------------
        # REAL ROAD MATRIX
        # ----------------------------------------------------

        with st.spinner(
            "🛣️ Calculating real road distances and driving times..."
        ):

            distance_matrix, duration_matrix = (
                get_route_matrix(
                    coordinates
                )
            )

        using_offline = False

        if (
            distance_matrix is None
            or duration_matrix is None
        ):

            using_offline = True

            st.warning(
                "⚠️ Live routing was unavailable. "
                "Using the offline distance/time fallback."
            )

            (
                distance_matrix,
                duration_matrix
            ) = build_offline_matrix(
                coordinates
            )

        # ----------------------------------------------------
        # BUILD STRICT NEAREST-FIRST ROUTE
        # ----------------------------------------------------

        route_indices = (
            build_nearest_first_route(
                distance_matrix,
                duration_matrix
            )
        )

        # ----------------------------------------------------
        # CALCULATE TOTAL ROUTE
        # ----------------------------------------------------

        (
            total_meters,
            total_seconds
        ) = calculate_route_totals(
            route_indices,
            distance_matrix,
            duration_matrix
        )

        total_km = (
            total_meters / 1000
        )

        total_miles = (
            total_km * 0.621371
        )

        # ----------------------------------------------------
        # FUEL
        # ----------------------------------------------------

        fuel_litres = (
            total_miles
            / MPG
            * 4.54609
        )

        fuel_cost = (
            fuel_litres
            * FUEL_PRICE
        )

        # ----------------------------------------------------
        # REVENUE
        # ----------------------------------------------------

        customer_revenue = (
            routing_df["Price"].sum()
        )

        pre_tax_profit = (
            customer_revenue
            - fuel_cost
        )

        take_home_profit = (
            pre_tax_profit
            * (1 - TAX_RATE)
        )

        customer_count = (
            len(route_indices) - 2
        )

        # ----------------------------------------------------
        # REORDER DATA
        # ----------------------------------------------------

        route_df = (
            routing_df
            .iloc[route_indices]
            .copy()
            .reset_index(drop=True)
        )

        route_df["route_index"] = (
            range(len(route_df))
        )

        # First and last rows are depot.
        route_df.loc[
            0,
            "Status"
        ] = "depot"

        route_df.loc[
            len(route_df) - 1,
            "Status"
        ] = "depot"

        # ----------------------------------------------------
        # SAVE ROUTE
        # ----------------------------------------------------

        st.session_state.master_df = (
            route_df
        )

        st.session_state.route_data = {

            "total_miles":
                float(total_miles),

            "total_km":
                float(total_km),

            "total_seconds":
                float(total_seconds),

            "fuel_litres":
                float(fuel_litres),

            "fuel_cost":
                float(fuel_cost),

            "revenue":
                float(customer_revenue),

            "pre_tax_profit":
                float(pre_tax_profit),

            "take_home_profit":
                float(take_home_profit),

            "customer_count":
                int(customer_count),

            "offline":
                bool(using_offline)
        }

        st.rerun()


    # ========================================================
    # ROUTE ECONOMICS
    # ========================================================

    if "route_data" in st.session_state:

        data = st.session_state.get(
            "route_data",
            {}
        )

        # ----------------------------------------------------
        # SAFE VALUES
        # ----------------------------------------------------

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

        customer_count = int(
            data.get(
                "customer_count",
                0
            )
        )

        offline = bool(
            data.get(
                "offline",
                False
            )
        )

        # ----------------------------------------------------
        # DASHBOARD
        # ----------------------------------------------------

        st.markdown("---")

        st.subheader(
            "💰 Today's Route Economics"
        )

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

        st.caption(
            f"{customer_count} customer stops"
        )

        if offline:

            st.warning(
                "⚠️ These route figures use the "
                "offline routing fallback."
            )

        # ----------------------------------------------------
        # ROUTING RULE
        # ----------------------------------------------------

        st.info(
            "🧭 **Routing rule:** Starting from the "
            "depot, the app always chooses the "
            "**nearest remaining customer by real "
            "driving distance**. Driving time breaks "
            "ties. The route is not rearranged afterwards."
        )


    # ========================================================
    # UNROUTED CUSTOMERS
    # ========================================================

    if (
        "unrouted_df" in st.session_state
        and not st.session_state.unrouted_df.empty
    ):

        st.markdown("---")

        st.error(
            f"⚠️ {len(st.session_state.unrouted_df)} "
            "customer(s) could not be located."
        )

        with st.expander(
            "View customers needing address checking"
        ):

            for _, failed_row in (
                st.session_state
                .unrouted_df
                .iterrows()
            ):

                postcode = clean_val(
                    failed_row.get(
                        "Postcode"
                    )
                )

                geo_error = clean_val(
                    failed_row.get(
                        "Geo Error"
                    )
                )

                st.write(
                    f"**{postcode}** — {geo_error}"
                )

            st.caption(
                "These customers have NOT been "
                "silently placed at the depot."
            )


    # ========================================================
    # SIDEBAR ROUTE NAVIGATION
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

        # Customer rows only.
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

            next_destination = (
                get_map_destination(
                    next_row
                )
            )

            next_url = (
                google_maps_url(
                    next_destination
                )
            )

            st.sidebar.link_button(
                "🚗 Navigate to Next Stop",
                next_url,
                use_container_width=True
            )

            st.sidebar.caption(
                f"Next: {next_destination}"
            )

            st.sidebar.caption(
                f"{len(pending_df)} stops remaining"
            )

        else:

            st.sidebar.success(
                "✅ All customer stops completed!"
            )


    # ========================================================
    # MAIN ROUTE CARDS
    # ========================================================

    st.markdown("---")

    st.subheader(
        "📍 Planned Route"
    )

    for idx, row in (
        st.session_state.master_df
        .iterrows()
    ):

        # ----------------------------------------------------
        # DEPOT START
        # ----------------------------------------------------

        if idx == 0:

            with st.container(
                border=True
            ):

                st.write(
                    "### 🏠 START — DEPOT"
                )

                st.write(
                    DEPOT_FULL_ADDRESS
                )

            continue

        # ----------------------------------------------------
        # DEPOT RETURN
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
                    key=f"return_depot_{idx}",
                    use_container_width=True
                )

            continue

        # ----------------------------------------------------
        # CUSTOMER DATA
        # ----------------------------------------------------

        postcode = clean_val(
            row.get(
                "Postcode"
            )
        )

        price = row.get(
            "Price",
            0
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
        # EXTRA CUSTOMER INFORMATION
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
                f"**Price:** £{float(price):.2f}"
                f"  |  "
                f"**Status:** {status_text}"
                f"  |  "
                f"**Payment:** {payment}"
            )

            col1, col2 = st.columns(2)

            # ------------------------------------------------
            # COMPLETE / WHATSAPP
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
                            f"Total: £{price}. "
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
            # NAVIGATION
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
        st.session_state
        .master_df
        .copy()
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
        "geo_query",
        "route_index"
    ]:

        if column in export_df.columns:

            export_df = (
                export_df
                .drop(columns=[column])
            )

    # --------------------------------------------------------
    # CREATE XLSX
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

        cell.fill = (
            header_fill
        )

        cell.font = (
            header_font
        )

        cell.alignment = Alignment(
            horizontal="center"
        )

    # Auto-width columns.
    for column_cells in worksheet.columns:

        max_length = 0

        column_letter = (
            column_cells[0]
            .column_letter
        )

        for cell in column_cells:

            try:

                cell_length = len(
                    str(cell.value)
                )

                max_length = max(
                    max_length,
                    cell_length
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
