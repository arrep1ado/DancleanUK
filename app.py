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
# DAN CLEAN UK - DAILY ROUTE OPTIMIZER
# Version 11.0
# ============================================================

APP_VERSION = "11.0"
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

    return pd.DataFrame([dict(row) for row in rows])


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


def get_coords(query_string, postcode):
    query = str(query_string).strip()
    postcode = normalise_postcode(postcode)
    key = cache_key_for(query, postcode)

    if key in st.session_state.geocode_cache:
        return st.session_state.geocode_cache[key]

    headers = {
        "User-Agent": "DanCleanUKRouteOptimizer/11.0"
    }

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
                timeout=10,
            )

            if response.status_code == 200:
                data = response.json()
                if data:
                    coords = (
                        float(data[0]["lat"]),
                        float(data[0]["lon"]),
                    )
                    st.session_state.geocode_cache[key] = coords
                    return coords
        except Exception:
            pass

    if postcode:
        try:
            response = requests.get(
                "https://nominatim.openstreetmap.org/search",
                params={
                    "q": f"{postcode}, United Kingdom",
                    "format": "json",
                    "limit": 1,
                    "countrycodes": "gb",
                },
                headers=headers,
                timeout=10,
            )

            if response.status_code == 200:
                data = response.json()
                if data:
                    coords = (
                        float(data[0]["lat"]),
                        float(data[0]["lon"]),
                    )
                    st.session_state.geocode_cache[key] = coords
                    return coords
        except Exception:
            pass

    pc = postcode.replace(" ", "")
    if pc:
        try:
            response = requests.get(
                f"https://api.postcodes.io/postcodes/{pc}",
                timeout=8,
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


def calculate_cluster_penalty(route, distances):
    penalty = 0.0
    customers = route[1:-1]

    for position in range(1, len(route) - 1):
        current = route[position]

        remaining = [
            x for x in customers
            if x not in route[:position + 1]
        ]

        if not remaining:
            continue

        next_stop = route[position + 1]
        next_distance = distances[current][next_stop]

        nearby = [
            distances[current][job]
            for job in remaining
            if distances[current][job] <= 15000
        ]

        if nearby:
            nearest_local = min(nearby)

            if next_distance > nearest_local * 2.5:
                penalty += (next_distance - nearest_local) * 3.0

    return penalty


def route_score(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
):
    metrics = route_metrics(
        route,
        distances,
        durations,
        fuel_price,
        mpg,
    )

    cluster_penalty = calculate_cluster_penalty(
        route,
        distances,
    )

    return (
        metrics["time_s"] * TIME_PRIORITY
        + metrics["distance_m"] * DISTANCE_PRIORITY
        + cluster_penalty * CLUSTER_PRIORITY
    )


def build_greedy_route(first_customer, distances, durations):
    customer_count = len(distances) - 1

    route = [0, first_customer]

    remaining = set(range(1, customer_count + 1))
    remaining.discard(first_customer)

    current = first_customer

    while remaining:
        candidates = []

        for candidate in remaining:
            direct_time = durations[current][candidate]

            future = [
                x for x in remaining if x != candidate
            ]

            if future:
                nearest_future = min(
                    future,
                    key=lambda x: durations[candidate][x],
                )
                future_time = durations[candidate][nearest_future]
            else:
                future_time = durations[candidate][0]

            score = direct_time * 0.70 + future_time * 0.30
            candidates.append((score, candidate))

        candidates.sort(key=lambda x: x[0])
        next_customer = candidates[0][1]

        route.append(next_customer)
        remaining.remove(next_customer)
        current = next_customer

    route.append(0)
    return route


def two_opt(route, distances, durations, fuel_price, mpg):
    best = route[:]
    best_score = route_score(
        best,
        distances,
        durations,
        fuel_price,
        mpg,
    )

    improved = True

    while improved:
        improved = False

        for i in range(1, len(best) - 2):
            for j in range(i + 1, len(best) - 1):
                candidate = (
                    best[:i]
                    + best[i:j + 1][::-1]
                    + best[j + 1:]
                )

                candidate_score = route_score(
                    candidate,
                    distances,
                    durations,
                    fuel_price,
                    mpg,
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
):
    best = route[:]
    best_score = route_score(
        best,
        distances,
        durations,
        fuel_price,
        mpg,
    )

    improved = True

    while improved:
        improved = False

        for i in range(1, len(best) - 1):
            customer = best[i]
            shortened = best[:i] + best[i + 1:]

            for j in range(1, len(shortened)):
                candidate = (
                    shortened[:j]
                    + [customer]
                    + shortened[j:]
                )

                candidate_score = route_score(
                    candidate,
                    distances,
                    durations,
                    fuel_price,
                    mpg,
                )

                if candidate_score < best_score - 0.01:
                    best = candidate
                    best_score = candidate_score
                    improved = True
                    break

            if improved:
                break

    return best


def generate_candidate_routes(distances, durations):
    customer_count = len(distances) - 1

    if customer_count <= 0:
        return []

    candidates = []

    # Do not blindly force the nearest customer first.
    # Test several strong starting candidates.
    start_candidates = sorted(
        range(1, customer_count + 1),
        key=lambda x: durations[0][x],
    )

    # Include the nearest, several nearby starts and a few
    # geographically expensive starts so the optimiser can
    # discover a better global route.
    selected = start_candidates[:min(10, customer_count)]

    if customer_count > 10:
        selected += [
            start_candidates[-1],
            start_candidates[len(start_candidates) // 2],
        ]

    selected = list(dict.fromkeys(selected))

    for first_customer in selected:
        candidates.append(
            build_greedy_route(
                first_customer,
                distances,
                durations,
            )
        )

    # A few randomised routes add diversity for larger days.
    customer_indexes = list(range(1, customer_count + 1))

    for _ in range(min(15, max(3, customer_count // 2))):
        shuffled = customer_indexes[:]
        random.shuffle(shuffled)
        candidates.append([0] + shuffled + [0])

    return candidates


def optimise_route(distances, durations, fuel_price, mpg):
    candidates = generate_candidate_routes(
        distances,
        durations,
    )

    best_route = None
    best_score = float("inf")

    for candidate in candidates:
        improved = two_opt(
            candidate,
            distances,
            durations,
            fuel_price,
            mpg,
        )

        improved = relocate_improvement(
            improved,
            distances,
            durations,
            fuel_price,
            mpg,
        )

        improved = two_opt(
            improved,
            distances,
            durations,
            fuel_price,
            mpg,
        )

        score = route_score(
            improved,
            distances,
            durations,
            fuel_price,
            mpg,
        )

        if score < best_score:
            best_score = score
            best_route = improved

    return best_route


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
            existing_day["status"].astype(str).str.lower() != "depot"
        ].copy()

        if not customer_rows.empty:
            completed = (
                customer_rows["status"]
                .astype(str)
                .str.lower()
                .eq("completed")
                .sum()
            )
            st.session_state.route_data = {
                "revenue": float(customer_rows["price"].sum()),
                "fuel_cost": 0.0,
                "take_home": float(
                    customer_rows["price"].sum()
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
