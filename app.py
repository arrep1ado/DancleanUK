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
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="DanCleanUK Optimizer",
    page_icon="🚗",
    layout="centered"
)

st.title("🚗 Daily Route & Profit Optimizer")

st.markdown(
    """
    <style>
    body, html {
        overscroll-behavior-y: none;
    }

    .route-number {
        background-color: #2F4F4F;
        color: white;
        padding: 4px 9px;
        border-radius: 50%;
        font-weight: bold;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# RESET
# ============================================================

if st.sidebar.button("🔄 Start New Day / Reset"):
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

TAX_RATE = st.sidebar.slider(
    "Tax Deduction (%)",
    0,
    50,
    20
) / 100

# ============================================================
# API KEY
# ============================================================

if "API_KEY" not in st.secrets:
    st.error(
        "API_KEY is missing from Streamlit secrets. "
        "Please add your OpenRouteService API key."
    )
    st.stop()

API_KEY = st.secrets["API_KEY"]

# ============================================================
# SESSION STATE
# ============================================================

if "geocode_cache" not in st.session_state:
    st.session_state.geocode_cache = {}

# ============================================================
# EXPORT
# ============================================================

st.sidebar.markdown("---")
st.sidebar.title("📊 Export Records")

if "master_df" in st.session_state and not st.session_state.master_df.empty:

    export_df = st.session_state.master_df.copy()

    if "Status" in export_df.columns:
        export_df = export_df[
            export_df["Status"].str.lower() != "depot"
        ].copy()

    for col in [
        "latitude",
        "longitude",
        "geo_query",
        "route_index"
    ]:
        if col in export_df.columns:
            export_df = export_df.drop(columns=[col])

    output = io.BytesIO()

    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Report"

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
        ws.append(row)

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center"
        )

    wb.save(output)

    st.sidebar.download_button(
        "⬇️ Download Professional XLSX",
        output.getvalue(),
        "DanCleanUK_Records.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.sidebar.info("Upload a file to begin.")

# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(
    "Upload your day's file",
    type=["csv", "xlsx"],
    help="Required columns: Postcode, Price, Phone"
)

if uploaded_file and "master_df" not in st.session_state:

    try:

        if uploaded_file.name.lower().endswith(".xlsx"):
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_csv(uploaded_file)

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        required = ["Postcode", "Price", "Phone"]

        missing = [
            col for col in required
            if col not in df.columns
        ]

        if missing:
            st.error(
                "Missing required columns: "
                + ", ".join(missing)
            )
            st.stop()

        # Remove completely empty rows
        df = df.dropna(how="all").copy()

        # ----------------------------------------------------
        # POSTCODE CLEANING
        # ----------------------------------------------------

        df["Postcode"] = (
            df["Postcode"]
            .fillna("")
            .astype(str)
            .str.upper()
            .str.strip()
        )

        df = df[
            (df["Postcode"] != "") &
            (df["Postcode"] != "NAN") &
            (df["Postcode"] != "NAT")
        ].copy()

        # ----------------------------------------------------
        # PRICE CLEANING
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
        # PHONE CLEANING
        # ----------------------------------------------------

        df["Phone"] = (
            df["Phone"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df = df[
            ~df["Phone"].str.lower().isin(
                ["", "nan", "nat"]
            )
        ].copy()

        # ----------------------------------------------------
        # INITIAL STATUS
        # ----------------------------------------------------

        df["Status"] = "pending"
        df["Payment"] = "waiting"

        df = df.reset_index(drop=True)

        st.session_state.master_df = df

        st.rerun()

    except Exception as e:

        st.error(
            f"Error loading file: {e}"
        )

# ============================================================
# HELPERS
# ============================================================

def clean_val(value):
    """
    Safely convert Excel values to readable strings.
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


def build_geo_query(row, default_postcode):
    """
    Build the best possible address for geocoding.
    """

    parts = []

    preferred_columns = [
        "address",
        "street",
        "location",
        "house",
        "house number"
    ]

    for col in row.index:

        if col.lower() in preferred_columns:

            value = clean_val(
                row[col]
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
# GEOCODING
# ============================================================

def get_coords(query_string, postcode_fallback):
    """
    Multi-tier geocoder.

    IMPORTANT:
    Failed geocoding returns None instead of silently
    placing the customer at the depot.
    """

    query = str(
        query_string
    ).strip()

    postcode = str(
        postcode_fallback
    ).upper().strip()

    cache_key = (
        f"{query}|{postcode}"
    ).lower()

    if cache_key in st.session_state.geocode_cache:
        return st.session_state.geocode_cache[
            cache_key
        ]

    headers = {
        "User-Agent":
            "DanCleanUKOptimizer/2.0"
    }

    # --------------------------------------------------------
    # NOMINATIM ADDRESS
    # --------------------------------------------------------

    if query and query.lower() != "nan":

        full_query = (
            query
            if (
                "uk" in query.lower()
                or "united kingdom"
                in query.lower()
            )
            else f"{query}, United Kingdom"
        )

        try:

            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": full_query,
                    "format": "json",
                    "limit": 1
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
    # NOMINATIM POSTCODE
    # --------------------------------------------------------

    if postcode:

        try:

            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": f"{postcode}, United Kingdom",
                    "format": "json",
                    "limit": 1
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

        # ----------------------------------------------------
        # POSTCODES.IO
        # ----------------------------------------------------

        postcode_clean = (
            postcode.replace(" ", "")
        )

        try:

            response = requests.get(
                f"https://api.postcodes.io/postcodes/{postcode_clean}",
                timeout=4
            )

            if response.status_code == 200:

                result = response.json().get(
                    "result"
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
# OFFLINE DISTANCE / TIME
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


def build_offline_matrix(locations):

    """
    Fallback matrix.

    locations are [lon, lat].
    """

    n = len(locations)

    distance = [
        [0.0 for _ in range(n)]
        for _ in range(n)
    ]

    duration = [
        [0.0 for _ in range(n)]
        for _ in range(n)
    ]

    ROAD_FACTOR = 1.30

    # Conservative average UK driving speed
    # for fallback calculations.
    AVG_SPEED_KMH = 40.0

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
            ) * ROAD_FACTOR

            distance[i][j] = km * 1000

            duration[i][j] = (
                km / AVG_SPEED_KMH
            ) * 3600

    return distance, duration


# ============================================================
# OPENROUTESERVICE MATRIX
# ============================================================

def get_route_matrix(locations):

    """
    Get real driving distance + duration.

    Returns:
        distances in metres
        durations in seconds
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

        if response.status_code == 200:

            data = response.json()

            distances = data.get(
                "distances"
            )

            durations = data.get(
                "durations"
            )

            if distances and durations:
                return distances, durations

        return None, None

    except Exception:
        return None, None


# ============================================================
# ROUTE OPTIMIZER
# ============================================================

def build_nearest_first_route(
    distance_matrix,
    duration_matrix
):

    """
    CORE ROUTING RULE:

    Starting at the depot, always select the
    nearest / quickest reachable unvisited customer.

    Priority:
        1. Driving time
        2. Driving distance

    This intentionally does NOT run 2-opt afterwards.

    Therefore the route always follows the
    nearest-next rule.
    """

    total_stops = len(
        distance_matrix
    )

    unvisited = set(
        range(1, total_stops)
    )

    route = [0]

    current = 0

    while unvisited:

        candidates = []

        for candidate in unvisited:

            travel_time = duration_matrix[
                current
            ][candidate]

            travel_distance = distance_matrix[
                current
            ][candidate]

            candidates.append(
                (
                    travel_time,
                    travel_distance,
                    candidate
                )
            )

        # Fastest reachable stop first.
        # Distance breaks ties.
        candidates.sort(
            key=lambda x: (
                x[0],
                x[1]
            )
        )

        next_stop = candidates[0][2]

        route.append(
            next_stop
        )

        unvisited.remove(
            next_stop
        )

        current = next_stop

    # Return to depot
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

    total_distance = 0
    total_duration = 0

    for i in range(
        len(route) - 1
    ):

        a = route[i]
        b = route[i + 1]

        total_distance += (
            distance_matrix[a][b]
        )

        total_duration += (
            duration_matrix[a][b]
        )

    return (
        total_distance,
        total_duration
    )


def format_duration(seconds):

    minutes = round(
        seconds / 60
    )

    hours = minutes // 60
    mins = minutes % 60

    if hours:
        return f"{hours}h {mins}m"

    return f"{mins}m"


# ============================================================
# GOOGLE MAPS
# ============================================================

def google_maps_url(destination):

    return (
        "https://www.google.com/maps/dir/"
        "?api=1"
        f"&destination={quote(str(destination))}"
        "&travelmode=driving"
    )


def get_map_destination(row):

    if (
        "geo_query" in row
        and pd.notna(row["geo_query"])
        and str(row["geo_query"]).strip()
    ):
        return str(
            row["geo_query"]
        )

    parts = []

    for col in row.index:

        if col.lower() in [
            "address",
            "street",
            "location",
            "house"
        ]:

            value = clean_val(
                row[col]
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
# ROUTING
# ============================================================

if "master_df" in st.session_state:

    df = st.session_state.master_df

    # --------------------------------------------------------
    # OPTIMIZE BUTTON
    # --------------------------------------------------------

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

        depot_coords = get_coords(
            DEPOT_FULL_ADDRESS,
            DEPOT_POSTCODE
        )

        if depot_coords is None:

            st.error(
                "Could not locate the depot."
            )
            st.stop()

        queries = []
        coordinates = [
            [
                depot_coords[1],
                depot_coords[0]
            ]
        ]

        valid_rows = []
        failed_rows = []

        # ----------------------------------------------------
        # GEOCODE CUSTOMERS
        # ----------------------------------------------------

        with st.spinner(
            "📍 Finding every customer location..."
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
                        (
                            original_index,
                            query,
                            postcode
                        )
                    )

                else:

                    valid_rows.append(
                        original_index
                    )

                    queries.append(query)

                    coordinates.append(
                        [
                            coords[1],
                            coords[0]
                        ]
                    )

                # Respect geocoding services
                time.sleep(0.1)

        # ----------------------------------------------------
        # SHOW FAILED ADDRESSES
        # ----------------------------------------------------

        if failed_rows:

            st.warning(
                f"{len(failed_rows)} customer(s) "
                "could not be located and were NOT "
                "included in the route."
            )

            with st.expander(
                "View addresses that need checking"
            ):

                for _, query, postcode in failed_rows:

                    st.write(
                        f"**{postcode}** — {query}"
                    )

        if not valid_rows:

            st.error(
                "No customer addresses could be located."
            )
            st.stop()

        # ----------------------------------------------------
        # BUILD ROUTING DATA
        # ----------------------------------------------------

        routing_rows = []

        depot_row = {
            "Postcode": DEPOT_POSTCODE,
            "Price": 0.0,
            "Phone": "",
            "Status": "depot",
            "Payment": "waiting",
            "geo_query": DEPOT_FULL_ADDRESS,
            "latitude": depot_coords[0],
            "longitude": depot_coords[1]
        }

        routing_rows.append(
            depot_row
        )

        for i, original_index in enumerate(
            valid_rows
        ):

            row = df.loc[
                original_index
            ].copy()

            coords = st.session_state.geocode_cache[
                (
                    f"{build_geo_query(row, DEPOT_POSTCODE)}"
                    f"|{clean_val(row['Postcode'])}"
                ).lower()
            ]

            row_data = row.to_dict()

            row_data["geo_query"] = (
                build_geo_query(
                    row,
                    DEPOT_POSTCODE
                )
            )

            row_data["latitude"] = coords[0]
            row_data["longitude"] = coords[1]

            routing_rows.append(
                row_data
            )

        routing_df = pd.DataFrame(
            routing_rows
        )

        # ----------------------------------------------------
        # GET REAL DRIVING MATRIX
        # ----------------------------------------------------

        with st.spinner(
            "🛣️ Calculating real driving times and distances..."
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
                "OpenRouteService did not return a "
                "driving matrix. Using offline routing "
                "fallback."
            )

            distance_matrix, duration_matrix = (
                build_offline_matrix(
                    coordinates
                )
            )

        # ----------------------------------------------------
        # NEAREST-FIRST ROUTE
        # ----------------------------------------------------

        route_indices = build_nearest_first_route(
            distance_matrix,
            duration_matrix
        )

        # ----------------------------------------------------
        # CALCULATE TOTALS
        # ----------------------------------------------------

        total_meters, total_seconds = (
            calculate_route_totals(
                route_indices,
                distance_matrix,
                duration_matrix
            )
        )

        total_km = (
            total_meters / 1000
        )

        total_miles = (
            total_km * 0.621371
        )

        total_hours = (
            total_seconds / 3600
        )

        fuel_litres = (
            total_miles / MPG
            * 4.54609
        )

        fuel_cost = (
            fuel_litres
            * FUEL_PRICE
        )

        revenue = routing_df[
            "Price"
        ].sum()

        pre_tax_profit = (
            revenue
            - fuel_cost
        )

        take_home_profit = (
            pre_tax_profit
            * (1 - TAX_RATE)
        )

        customer_count = len(
            route_indices
        ) - 2

        # ----------------------------------------------------
        # REORDER DATAFRAME
        # ----------------------------------------------------

        route_df = routing_df.iloc[
            route_indices
        ].copy().reset_index(
            drop=True
        )

        route_df["route_index"] = (
            range(len(route_df))
        )

        route_df.loc[
            0,
            "Status"
        ] = "depot"

        route_df.loc[
            len(route_df) - 1,
            "Status"
        ] = "depot"

        route_df["Payment"] = (
            route_df["Payment"]
            if "Payment" in route_df.columns
            else "waiting"
        )

        # ----------------------------------------------------
        # SAVE
        # ----------------------------------------------------

        st.session_state.master_df = (
            route_df
        )

        st.session_state.route_data = {

            "total_miles":
                total_miles,

            "total_km":
                total_km,

            "total_seconds":
                total_seconds,

            "fuel_litres":
                fuel_litres,

            "fuel_cost":
                fuel_cost,

            "revenue":
                revenue,

            "pre_tax_profit":
                pre_tax_profit,

            "take_home_profit":
                take_home_profit,

            "customer_count":
                customer_count,

            "offline":
                using_offline
        }

        st.rerun()

    # ========================================================
    # DASHBOARD
    # ========================================================

    if "route_data" in st.session_state:

        data = st.session_state.route_data

        st.markdown("---")

        st.subheader(
            "💰 Today's Route Economics"
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
                "Estimated Driving Time",
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

        st.caption(
            f"{data['customer_count']} customer stops"
        )

        if data["offline"]:
            st.warning(
                "Route was calculated using the "
                "offline fallback rather than live "
                "driving-road data."
            )

    # ========================================================
    # ROUTE NAVIGATION SIDEBAR
    # ========================================================

    if (
        "route_data" in st.session_state
        and not st.session_state.master_df.empty
    ):

        st.sidebar.markdown("---")
        st.sidebar.title("🧭 Route Navigation")

        current_df = (
            st.session_state.master_df
        )

        customer_df = current_df[
            current_df["Status"].str.lower()
            != "depot"
        ]

        pending_df = customer_df[
            customer_df["Status"].str.lower()
            == "pending"
        ]

        if not pending_df.empty:

            next_row = pending_df.iloc[0]

            next_destination = (
                get_map_destination(next_row)
            )

            next_url = google_maps_url(
                next_destination
            )

            st.sidebar.link_button(
                "🚗 Navigate to Next Stop",
                next_url,
                use_container_width=True
            )

            st.sidebar.caption(
                f"Next stop: {next_destination}"
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

    for idx, row in (
        st.session_state.master_df.iterrows()
    ):

        is_first = (
            idx == 0
        )

        is_last = (
            idx
            ==
            len(st.session_state.master_df) - 1
        )

        # ----------------------------------------------------
        # START DEPOT
        # ----------------------------------------------------

        if is_first:

            with st.container(
                border=True
            ):

                st.write(
                    "### 🏠 Start Depot"
                )

                st.write(
                    DEPOT_FULL_ADDRESS
                )

            continue

        # ----------------------------------------------------
        # RETURN DEPOT
        # ----------------------------------------------------

        if is_last:

            return_url = google_maps_url(
                DEPOT_FULL_ADDRESS
            )

            with st.container(
                border=True
            ):

                st.write(
                    "### 🏁 Return to Depot"
                )

                st.write(
                    DEPOT_FULL_ADDRESS
                )

                st.link_button(
                    "🚗 Navigate Back to Depot",
                    return_url,
                    key=f"return_{idx}"
                )

            continue

        # ----------------------------------------------------
        # CUSTOMER
        # ----------------------------------------------------

        postcode = clean_val(
            row.get("Postcode")
        )

        price = row.get(
            "Price",
            0
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
            get_map_destination(row)
        )

        map_url = google_maps_url(
            destination
        )

        # ----------------------------------------------------
        # EXTRA INFO
        # ----------------------------------------------------

        extra_parts = []

        for col in (
            st.session_state.master_df.columns
        ):

            if col.lower() in [
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
                "date" in col.lower()
                or "time" in col.lower()
            ):
                continue

            value = clean_val(
                row.get(col)
            )

            if value:
                extra_parts.append(
                    f"**{col}:** {value}"
                )

        # ----------------------------------------------------
        # CARD
        # ----------------------------------------------------

        icon = (
            "✅"
            if status == "completed"
            else "⏳"
        )

        with st.container(
            border=True
        ):

            st.write(
                f"### {icon} Stop {idx} — {postcode}"
            )

            if extra_parts:
                st.write(
                    " | ".join(
                        extra_parts
                    )
                )

            st.write(
                f"**Price:** £{float(price):.2f}  |  "
                f"**Status:** "
                f"{status.title()}  |  "
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
            # PAYMENT
            # ------------------------------------------------

            with col2:

                st.link_button(
                    "🚗 Navigate Here",
                    map_url,
                    key=f"nav_{idx}",
                    use_container_width=True
                )

            p1, p2 = st.columns(2)

            with p1:

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

            with p2:

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

else:

    st.info(
        "📁 Upload your day's file to begin."
    )
