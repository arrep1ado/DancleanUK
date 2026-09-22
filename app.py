import io
import base64
import json
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
# Version 26.10
# ============================================================

APP_VERSION = "27.7.1"
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
            notes TEXT NOT NULL DEFAULT '',
            created_at TEXT NOT NULL
        )
        """
    )
    columns = {row[1] for row in conn.execute("PRAGMA table_info(jobs)").fetchall()}
    if "notes" not in columns:
        conn.execute("ALTER TABLE jobs ADD COLUMN notes TEXT NOT NULL DEFAULT ''")
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
            geo_query, notes, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            geo_query=excluded.geo_query,
            notes=excluded.notes
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
            str(row.get("Notes", row.get("notes", "")) or ""),
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
            "notes": "Notes",
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
        "Notes": "",
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
    loaded["Notes"] = loaded["Notes"].fillna("")

    return loaded


def delete_day(service_date):
    conn = db_connect()
    conn.execute(
        "DELETE FROM jobs WHERE service_date = ?",
        (service_date,),
    )
    conn.commit()
    conn.close()


def _supabase_config():
    """Return private server-side Supabase settings from Streamlit Secrets."""
    try:
        url = str(st.secrets["SUPABASE_URL"]).strip().rstrip("/")
        key = str(st.secrets["SUPABASE_SECRET_KEY"]).strip()
    except Exception:
        return None, None
    if not url.startswith("https://") or not key.startswith("sb_secret_"):
        return None, None
    return url, key


def _supabase_headers(prefer=None):
    url, key = _supabase_config()
    if not url or not key:
        return None
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        headers["Prefer"] = prefer
    return headers


def _json_safe(value):
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def init_saved_routes_db():
    # V27.2 stores locked route snapshots in Supabase. The existing local
    # SQLite jobs database is deliberately left unchanged.
    return True


def save_route_snapshot(service_date, df, route_data):
    if df is None or df.empty or not route_data:
        return False

    routed = df[
        df["route_order"].notna()
        & (df["Status"].astype(str).str.lower() != "depot")
    ].copy()
    if routed.empty:
        return False

    routed["_saved_order"] = pd.to_numeric(routed["route_order"], errors="coerce")
    routed = routed.dropna(subset=["_saved_order"]).sort_values("_saved_order")
    if len(routed) != int(route_data.get("jobs", len(routed))):
        return False

    # Store complete job rows as well as the exact order. This means a locked
    # route can be reconstructed even if Streamlit's local SQLite is restarted.
    job_rows = []
    for _, row in routed.drop(columns=["_saved_order"], errors="ignore").iterrows():
        job_rows.append({str(k): _json_safe(v) for k, v in row.to_dict().items()})

    route_payload = {
        "route_job_ids": routed["job_id"].astype(str).tolist(),
        "jobs_data": job_rows,
        "litres": float(route_data.get("litres", 0.0)),
        "time_s": float(route_data.get("time", 0.0)),
        "saved_at": now_text(),
        "app_version": APP_VERSION,
    }
    existing_snapshot = load_route_snapshot(service_date)
    if existing_snapshot:
        existing_data = existing_snapshot.get("route_data") or {}
        for key in ("report_b64", "report_filename", "report_saved_at"):
            if existing_data.get(key):
                route_payload[key] = existing_data[key]
    total_minutes = int(round(float(route_data.get("time", 0.0)) / 60.0))
    payload = {
        "route_date": str(service_date),
        "route_data": route_payload,
        "total_jobs": int(route_data.get("jobs", len(routed))),
        "total_miles": round(float(route_data.get("miles", 0.0)), 2),
        "total_minutes": total_minutes,
        "revenue": round(float(route_data.get("revenue", 0.0)), 2),
        "fuel_cost": round(float(route_data.get("fuel_cost", 0.0)), 2),
        "take_home": round(float(route_data.get("take_home", 0.0)), 2),
        "updated_at": datetime.now().astimezone().isoformat(),
    }

    url, _ = _supabase_config()
    headers = _supabase_headers("resolution=merge-duplicates,return=minimal")
    if not url or not headers:
        return False
    try:
        response = requests.post(
            f"{url}/rest/v1/saved_routes?on_conflict=route_date",
            headers=headers,
            json=payload,
            timeout=20,
        )
        return response.status_code in (200, 201, 204)
    except requests.RequestException:
        return False


def load_route_snapshot(service_date):
    url, _ = _supabase_config()
    headers = _supabase_headers()
    if not url or not headers:
        return None
    try:
        response = requests.get(
            f"{url}/rest/v1/saved_routes",
            headers=headers,
            params={"route_date": f"eq.{service_date}", "select": "*", "limit": "1"},
            timeout=20,
        )
        if response.status_code != 200:
            return None
        rows = response.json()
        if not rows:
            return None
        row = rows[0]
        data = row.get("route_data") or {}
        row["route_data"] = data
        row["report_b64"] = data.get("report_b64") or ""
        row["report_filename"] = data.get("report_filename") or ""
        row["route_job_ids"] = [str(x) for x in data.get("route_job_ids", [])]
        row["jobs_data"] = data.get("jobs_data", [])
        row["litres"] = float(data.get("litres", 0.0) or 0.0)
        row["time_s"] = float(data.get("time_s", (row.get("total_minutes") or 0) * 60) or 0.0)
        row["jobs"] = int(row.get("total_jobs") or len(row["route_job_ids"]))
        row["miles"] = float(row.get("total_miles") or 0.0)
        row["saved_at"] = data.get("saved_at") or row.get("updated_at") or row.get("created_at") or ""
        return row
    except (requests.RequestException, ValueError, TypeError):
        return None


def delete_route_snapshot(service_date):
    url, _ = _supabase_config()
    headers = _supabase_headers("return=minimal")
    if not url or not headers:
        return False
    try:
        response = requests.delete(
            f"{url}/rest/v1/saved_routes",
            headers=headers,
            params={"route_date": f"eq.{service_date}"},
            timeout=20,
        )
        return response.status_code in (200, 204)
    except requests.RequestException:
        return False


def apply_saved_route_snapshot(service_date):
    snapshot = load_route_snapshot(service_date)
    if snapshot is None:
        return False, "No saved route is available for this date."

    route_ids = [str(x) for x in snapshot.get("route_job_ids", [])]
    if not route_ids:
        return False, "The saved route does not contain any customer stops."

    day_df = load_day(service_date)
    jobs_data = snapshot.get("jobs_data") or []

    # The permanent saved route is the master working-day record. Its job rows
    # carry the latest completion/payment/notes state between phone and laptop.
    # Local SQLite is only a working cache and must never overwrite newer
    # Supabase state after a reconnect or device change.
    if jobs_data:
        day_df = pd.DataFrame(jobs_data)
    elif day_df.empty:
        return False, "The permanent route exists, but its customer records are unavailable."
    else:
        current_ids = set(day_df["job_id"].astype(str))
        missing = [job_id for job_id in route_ids if job_id not in current_ids]
        if missing:
            return False, "The saved route no longer matches the jobs stored for this date."

    for column, default in {
        "Status": "pending", "Payment": "Waiting", "PaymentTime": "",
        "CompletedTime": "", "address_text": "", "geo_query": "", "Notes": "",
        "WhatsAppSent": False, "WhatsAppTime": "",
    }.items():
        if column not in day_df.columns:
            day_df[column] = default
        day_df[column] = day_df[column].fillna(default)

    order_map = {job_id: order for order, job_id in enumerate(route_ids)}
    day_df["route_order"] = day_df["job_id"].astype(str).map(order_map)
    day_df = day_df[day_df["job_id"].astype(str).isin(route_ids)].copy()
    day_df = day_df.sort_values("route_order").reset_index(drop=True)
    save_dataframe(day_df)

    st.session_state.master_df = day_df
    st.session_state.route_data = {
        "revenue": float(snapshot.get("revenue") or 0.0),
        "fuel_cost": float(snapshot.get("fuel_cost") or 0.0),
        "take_home": float(snapshot.get("take_home") or 0.0),
        "miles": float(snapshot.get("miles") or 0.0),
        "litres": float(snapshot.get("litres") or 0.0),
        "time": float(snapshot.get("time_s") or 0.0),
        "offline": False,
        "jobs": int(snapshot.get("jobs") or len(route_ids)),
        "completed": int(day_df["Status"].astype(str).str.lower().eq("completed").sum()),
        "persisted_only": False,
        "saved_route": True,
        "saved_at": clean_val(snapshot.get("saved_at")),
    }
    st.session_state.pop("failed_jobs", None)
    return True, "Permanent saved route loaded."


init_db()
init_saved_routes_db()


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
st.sidebar.subheader("💬 Customer Messages")
BUSINESS_NAME = st.sidebar.text_input("Business name", value="DanCleanUK")

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



def get_terminated_postcode_coords(postcode):
    """Return the last known coordinate for a terminated UK postcode.

    This coordinate is VALIDATION-ONLY. It is never used as a customer
    routing point and never supplies route distance/time. It lets genuine
    addresses with old postcodes pass geographic validation while still
    rejecting a street/address match that is in a completely different area.
    """
    postcode = normalise_postcode(postcode)
    if not postcode:
        return None

    cache_key = f"__TERMINATED_POSTCODE__|{postcode}".lower()
    if cache_key in st.session_state.geocode_cache:
        return st.session_state.geocode_cache.get(cache_key)

    for pc in dict.fromkeys([postcode, postcode.replace(" ", "")]):
        try:
            response = requests.get(
                f"https://api.postcodes.io/terminated_postcodes/{quote(pc)}",
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

    # Cache the miss for this session so every customer with the same bad/
    # unknown postcode does not repeatedly call the service.
    st.session_state.geocode_cache[cache_key] = None
    return None

def photon_exact_geocode(query, postcode, expected_house_number="", expected_street=""):
    """Try Photon/OSM address data for a true house-level coordinate.

    Photon can expose address points that are not returned by the Nominatim
    query used by the normal path.  It is an additional exact-address source,
    not a postcode fallback.  Results are accepted only when the house number
    and street agree with the customer's address.
    """
    expected_house_number = clean_val(expected_house_number)
    expected_street = clean_val(expected_street).lower()
    postcode = normalise_postcode(postcode)

    queries = []
    base = str(query or "").strip()
    if base:
        queries.append(base)
    if expected_house_number and expected_street and postcode:
        queries.append(
            f"{expected_house_number} {expected_street}, {postcode}, United Kingdom"
        )

    for search_query in dict.fromkeys(queries):
        try:
            response = requests.get(
                "https://photon.komoot.io/api/",
                params={
                    "q": search_query,
                    "limit": 20,
                },
                headers={
                    "User-Agent": "DanCleanUKRouteOptimizer/26.10"
                },
                timeout=20,
            )

            if response.status_code != 200:
                continue

            data = response.json() or {}
            for feature in data.get("features") or []:
                geometry = feature.get("geometry") or {}
                coords = geometry.get("coordinates") or []
                if len(coords) < 2:
                    continue

                try:
                    lon = float(coords[0])
                    lat = float(coords[1])
                except Exception:
                    continue

                if not (-90 <= lat <= 90 and -180 <= lon <= 180):
                    continue

                props = feature.get("properties") or {}
                returned_house = clean_val(
                    props.get("housenumber")
                    or props.get("house_number")
                )
                returned_street = clean_val(
                    props.get("street")
                    or props.get("name")
                ).lower()
                label = clean_val(
                    props.get("name")
                    or props.get("street")
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

        except Exception:
            continue

    return None


def _postcode_outcode(postcode):
    postcode = normalise_postcode(postcode)
    if not postcode:
        return ""
    return postcode.split()[0].upper()


def get_outcode_coords(postcode):
    """Return the official centroid for a UK postcode outcode.

    This is validation-only. It is useful when a full postcode is retired,
    mistyped, or unavailable from the live postcode service. It is never used
    as a customer routing coordinate and never supplies road distance/time.
    """
    outcode = _postcode_outcode(postcode)
    if not outcode:
        return None

    cache_key = f"__OUTCODE__|{outcode}".lower()
    cached = st.session_state.geocode_cache.get(cache_key)
    if cached is not None:
        return cached

    try:
        response = requests.get(
            f"https://api.postcodes.io/outcodes/{quote(outcode)}",
            timeout=10,
        )
        if response.status_code != 200:
            return None
        result = response.json().get("result") or {}
        lat = result.get("latitude")
        lon = result.get("longitude")
        if lat is None or lon is None:
            return None
        coords = (float(lat), float(lon))
        st.session_state.geocode_cache[cache_key] = coords
        return coords
    except Exception:
        return None


def reverse_postcode_outcode(lat, lon):
    """Return the nearest live UK postcode outcode for a coordinate.

    This is a validation check only. It never supplies route distance/time and
    never changes the road matrix. It prevents a geocoder from matching the
    right house/street name in the wrong part of the country.
    """
    try:
        response = requests.get(
            "https://api.postcodes.io/postcodes",
            params={"lat": float(lat), "lon": float(lon), "limit": 1},
            timeout=10,
        )
        if response.status_code != 200:
            return ""
        results = response.json().get("result") or []
        if not results:
            return ""
        nearest = normalise_postcode(results[0].get("postcode", ""))
        return _postcode_outcode(nearest)
    except Exception:
        return ""


def get_coords(query_string, postcode):
    """Locate a customer safely before any ORS road calculation.

    House-level geocoder results are accepted only when they agree with the
    supplied postcode geography. Cached coordinates are revalidated too, so a
    bad coordinate from an earlier lookup cannot survive inside the session.

    If the supplied postcode cannot be verified and the candidate coordinate
    does not even belong to the same postcode outcode, the customer is left
    unlocated rather than routing hundreds of real-road miles to a wrong place.
    """
    query = str(query_string or "").strip()
    postcode = normalise_postcode(postcode)
    key = cache_key_for(query, postcode)

    headers = {"User-Agent": "DanCleanUKRouteOptimizer/26.10"}

    house_match = re.search(r"(?<!\d)(\d+[A-Za-z]?)\b", query)
    expected_house = house_match.group(1) if house_match else ""

    expected_street = ""
    if house_match:
        tail = query[house_match.end():]
        expected_street = tail.split(",")[0].strip()

    postcode_anchor = get_postcode_coords(postcode)
    terminated_postcode_anchor = (
        None if postcode_anchor is not None
        else get_terminated_postcode_coords(postcode)
    )
    expected_outcode = _postcode_outcode(postcode)
    outcode_anchor = get_outcode_coords(postcode)

    def safe_exact(coords):
        if coords is None:
            return None
        try:
            lat, lon = float(coords[0]), float(coords[1])
        except Exception:
            return None

        if not (-90 <= lat <= 90 and -180 <= lon <= 180):
            return None

        # Strongest check: a live official postcode point exists. Exact house
        # coordinates must remain geographically close to that postcode.
        if postcode_anchor is not None:
            if haversine_km(
                postcode_anchor[0], postcode_anchor[1], lat, lon
            ) > 5.0:
                return None
            return (lat, lon)

        # A terminated postcode still has a last-known official coordinate.
        # Use it ONLY to validate the geocoded house/street. This is the key
        # distinction: old postcodes can remain useful evidence of the area,
        # but are never substituted as the customer's routing coordinate.
        if terminated_postcode_anchor is not None:
            if haversine_km(
                terminated_postcode_anchor[0],
                terminated_postcode_anchor[1],
                lat,
                lon,
            ) <= 5.0:
                return (lat, lon)
            return None

        # If neither a live nor terminated full postcode can be verified,
        # validate against the official OUTCODE geography. A nearby address
        # may legitimately reverse-geocode to an adjacent outcode, so exact
        # outcode equality is too strict.
        if outcode_anchor is not None:
            if haversine_km(
                outcode_anchor[0], outcode_anchor[1], lat, lon
            ) <= 20.0:
                return (lat, lon)
            return None

        # Last validation option when the outcode service has no centroid.
        # Exact outcode agreement is still useful, but failure means we do not
        # guess.
        if expected_outcode:
            candidate_outcode = reverse_postcode_outcode(lat, lon)
            if candidate_outcode == expected_outcode:
                return (lat, lon)

        return None

    # Revalidate cached coordinates under the CURRENT safety rules. This is
    # essential because a previously cached wrong match must not bypass fixes.
    cached = st.session_state.geocode_cache.get(key)
    if cached is not None:
        checked_cached = safe_exact(cached)
        if checked_cached is not None:
            return checked_cached
        st.session_state.geocode_cache.pop(key, None)

    exact_ors = safe_exact(
        ors_exact_geocode(
            query,
            postcode,
            expected_house_number=expected_house,
            expected_street=expected_street,
        )
    )
    if exact_ors is not None:
        st.session_state.geocode_cache[key] = exact_ors
        return exact_ors

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

    photon = safe_exact(
        photon_exact_geocode(
            query,
            postcode,
            expected_house_number=expected_house,
            expected_street=expected_street,
        )
    )
    if photon is not None:
        st.session_state.geocode_cache[key] = photon
        return photon

    # If a geocoder has no house-number record, try the named STREET itself.
    # This is still a real map coordinate and ORS still supplies every road
    # distance/time.  Crucially, the street coordinate must pass the same
    # postcode/outcode geography check, so a same-named street in another
    # part of the country cannot be accepted.
    if expected_street:
        street_queries = []
        if postcode:
            street_queries.append(f"{expected_street}, {postcode}, United Kingdom")
        street_queries.append(f"{expected_street}, United Kingdom")

        seen_street_queries = set()
        for street_query in street_queries:
            sq_key = street_query.strip().lower()
            if not sq_key or sq_key in seen_street_queries:
                continue
            seen_street_queries.add(sq_key)

            street_coords = nominatim_search(
                street_query,
                headers,
                expected_house_number=None,
                expected_street=expected_street,
                postcode=postcode,
            )
            street_coords = safe_exact(street_coords)
            if street_coords is not None:
                st.session_state.geocode_cache[key] = street_coords
                return street_coords
            time.sleep(0.35)

    # Final safe postcode-level fallback. A LIVE postcode centroid is always
    # acceptable when the exact house/street is unavailable.
    if postcode_anchor is not None:
        st.session_state.geocode_cache[key] = postcode_anchor
        return postcode_anchor

    # A terminated postcode can still be useful for older customer records,
    # but only when the imported address does NOT contain an additional
    # locality/town that could conflict with that historic postcode.
    # Example safe shape: "45 Dysart Road, NG31 7AN, United Kingdom".
    # Example unsafe shape: "7 Church Street, Muston, OLD POSTCODE, UK".
    # In the unsafe case we keep the hard stop rather than silently routing to
    # the wrong village/town.
    if terminated_postcode_anchor is not None:
        address_without_country = re.sub(
            r",?\s*united kingdom\s*$",
            "",
            query,
            flags=re.IGNORECASE,
        )
        address_without_pc = re.sub(
            re.escape(postcode),
            "",
            address_without_country,
            flags=re.IGNORECASE,
        ).strip(" ,")

        # Split the remaining address. One component means house + street only;
        # two or more components usually means an explicit locality/town was
        # supplied and must be resolved rather than ignored.
        address_parts = [
            part.strip()
            for part in address_without_pc.split(",")
            if part.strip()
        ]

        if len(address_parts) <= 1:
            st.session_state.geocode_cache[key] = terminated_postcode_anchor
            return terminated_postcode_anchor

    # No verified location: do not guess.
    return None



def _route_group_street(value):
    value = clean_val(value).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _route_house_number(value):
    m = re.search(r"(?<!\d)(\d+)(?:[A-Za-z])?\b", clean_val(value))
    return int(m.group(1)) if m else None


def deconflict_same_postcode_houses(df, valid_rows, depot_coords):
    """Keep duplicate house records stable without inventing road geometry.

    Public geocoders sometimes return the same coordinate for several houses
    on one postcode/street. Those jobs remain separate customer records, but
    fabricated coordinate offsets can change the road matrix and push nearby
    houses to opposite ends of a route.

    Leave the validated geocoder/postcode coordinate unchanged. No postcode,
    street, depot, current job list, or route order is hard-coded here.
    """
    return df


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



def matrix_values_are_valid(distances, durations, expected_size):
    """Reject incomplete/non-finite/negative ORS matrices before optimisation."""
    if not isinstance(distances, list) or not isinstance(durations, list):
        return False

    if len(distances) != expected_size or len(durations) != expected_size:
        return False

    for matrix in (distances, durations):
        for row in matrix:
            if not isinstance(row, list) or len(row) != expected_size:
                return False
            for value in row:
                try:
                    value = float(value)
                except Exception:
                    return False
                if not math.isfinite(value) or value < 0:
                    return False

    return True


def suspicious_matrix_nodes(locations, distances, durations):
    """Find coordinates whose live-road legs are implausible for their geometry.

    This is deliberately generic: it knows nothing about Grantham, Queensway,
    today's postcodes, or a target route length. It compares ORS road distance
    with straight-line distance between the SAME two coordinates.

    A leg is suspicious only when the road detour is both very large in
    absolute terms and extreme relative to the straight-line separation.
    Requiring repeated suspicious legs before blaming a customer avoids
    changing a valid route because of one unusual bridge/road restriction.
    """
    n = len(locations)
    strikes = [0] * n
    worst_excess = [0.0] * n

    for i in range(n):
        lon1, lat1 = locations[i]
        for j in range(i + 1, n):
            lon2, lat2 = locations[j]
            straight_km = haversine_km(lat1, lon1, lat2, lon2)

            try:
                road_ij = float(distances[i][j]) / 1000.0
                road_ji = float(distances[j][i]) / 1000.0
                time_ij = float(durations[i][j])
                time_ji = float(durations[j][i])
            except Exception:
                strikes[i] += 2
                strikes[j] += 2
                continue

            # Use the smaller direction so normal one-way systems do not
            # trigger the validator merely because one direction is longer.
            road_km = min(road_ij, road_ji)
            drive_s = min(time_ij, time_ji)

            # Zero/near-zero geometry can legitimately have a small road snap.
            # What we are looking for is a many-tens-of-km detour between
            # coordinates that are geographically close.
            # Generic sanity envelope. For nearby coordinates, a road route
            # tens of kilometres longer than the geometry is suspicious. For
            # genuinely distant jobs, proportional allowance grows naturally.
            # No town, postcode, daily mileage target, or route order is used.
            allowed_km = max(20.0, straight_km * 5.0 + 8.0)
            excessive_road = road_km > allowed_km

            # Also reject impossible average speeds on a substantial leg.
            bad_speed = False
            if road_km >= 5.0 and drive_s > 0:
                speed_kph = road_km / (drive_s / 3600.0)
                bad_speed = speed_kph < 3.0 or speed_kph > 160.0

            if excessive_road or bad_speed:
                strikes[i] += 1
                strikes[j] += 1
                excess = max(0.0, road_km - allowed_km)
                worst_excess[i] = max(worst_excess[i], excess)
                worst_excess[j] = max(worst_excess[j], excess)

    # Depot is index 0 and is never auto-replaced here.
    suspects = [
        idx
        for idx in range(1, n)
        if strikes[idx] >= 2
    ]

    # Generic isolation check: if a customer has no reasonably local road
    # connection relative to its nearest geometric neighbour, flag it. This
    # catches a single wildly misplaced/snap-broken customer even when the
    # pairwise strike count would otherwise be too low.
    for i in range(1, n):
        geometric = []
        for j in range(n):
            if i == j:
                continue
            lon1, lat1 = locations[i]
            lon2, lat2 = locations[j]
            straight_km = haversine_km(lat1, lon1, lat2, lon2)
            geometric.append((straight_km, j))

        geometric.sort(key=lambda item: (item[0], item[1]))
        for straight_km, j in geometric[:3]:
            road_km = min(
                float(distances[i][j]),
                float(distances[j][i]),
            ) / 1000.0
            allowed_km = max(20.0, straight_km * 5.0 + 8.0)
            if road_km > allowed_km:
                strikes[i] += 1
                worst_excess[i] = max(
                    worst_excess[i],
                    road_km - allowed_km,
                )

    suspects = [
        idx
        for idx in range(1, n)
        if strikes[idx] >= 2
    ]
    suspects.sort(
        key=lambda idx: (strikes[idx], worst_excess[idx], idx),
        reverse=True,
    )
    return suspects


def get_validated_ors_matrix(locations, postcode_fallbacks, max_repairs=4):
    """Build a live ORS matrix and repair bad address snaps generically.

    Exact address coordinates remain preferred. If the live matrix shows that
    one customer coordinate repeatedly creates implausible road detours, only
    that customer is moved to its official postcode centroid and the matrix is
    requested again. This makes daily address changes safe without hard-coding
    any route, postcode, town, or customer.
    """
    working_locations = [
        [float(lon), float(lat)]
        for lon, lat in locations
    ]
    repaired = []

    for _ in range(max_repairs + 1):
        distances, durations = get_ors_matrix(working_locations)

        if not matrix_values_are_valid(
            distances,
            durations,
            len(working_locations),
        ):
            return None, None, working_locations, repaired

        suspects = suspicious_matrix_nodes(
            working_locations,
            distances,
            durations,
        )

        repair_idx = None
        for idx in suspects:
            fallback = postcode_fallbacks[idx]
            if fallback is None:
                continue

            fallback_lon, fallback_lat = fallback
            current_lon, current_lat = working_locations[idx]

            # Do not waste an API retry if the postcode point is effectively
            # identical to the coordinate already being used.
            if haversine_km(
                current_lat,
                current_lon,
                fallback_lat,
                fallback_lon,
            ) < 0.03:
                continue

            repair_idx = idx
            break

        if repair_idx is None:
            # Suspicious live-road data remains, but none of the flagged
            # customer points has a usable, materially different postcode
            # fallback. Do NOT silently accept this matrix as valid.
            return None, None, working_locations, repaired

        fallback_lon, fallback_lat = postcode_fallbacks[repair_idx]
        working_locations[repair_idx] = [
            float(fallback_lon),
            float(fallback_lat),
        ]
        repaired.append(repair_idx)

    # Final matrix after the allowed repairs.
    distances, durations = get_ors_matrix(working_locations)
    if not matrix_values_are_valid(
        distances,
        durations,
        len(working_locations),
    ):
        return None, None, working_locations, repaired

    # If it is still structurally suspicious, do not present it as a valid
    # live-road result. The caller must stop rather than substitute estimates.
    if suspicious_matrix_nodes(
        working_locations,
        distances,
        durations,
    ):
        return None, None, working_locations, repaired

    return distances, durations, working_locations, repaired


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






def route_leg_diagnostics(route, locations, distances, durations):
    """Return final-route leg diagnostics for validation/display."""
    rows = []
    for pos in range(len(route) - 1):
        a = route[pos]
        b = route[pos + 1]
        lon1, lat1 = locations[a]
        lon2, lat2 = locations[b]
        straight_km = haversine_km(lat1, lon1, lat2, lon2)
        road_km = float(distances[a][b]) / 1000.0
        seconds = float(durations[a][b])
        ratio = road_km / straight_km if straight_km >= 0.03 else None
        rows.append({
            "leg": pos + 1,
            "from_node": a,
            "to_node": b,
            "road_miles": road_km * 0.621371,
            "minutes": seconds / 60.0,
            "straight_miles": straight_km * 0.621371,
            "ratio": ratio,
        })
    return rows


def build_daily_pocket_route(
    distances,
    durations,
    locations,
    fuel_price,
    mpg,
):
    """Dynamic daily driver route using only today's road matrix.

    The current local neighbourhood is derived from the nearest unfinished
    road-time moves from the current position. Nearby jobs are cleared before
    a materially longer transition is made. There are no postcode/town rules
    and no transitive connected-component clustering.
    """
    customer_count = len(distances) - 1
    if customer_count <= 0:
        return [0, 0]

    remaining = set(range(1, customer_count + 1))
    route = [0]
    current = 0

    def edge_cost(a, b):
        miles = float(distances[a][b]) / 1000.0 * 0.621371
        litres = miles / float(mpg) * 4.54609
        return (
            litres * float(fuel_price)
            + (float(durations[a][b]) / 3600.0)
            * DRIVING_TIME_VALUE_PER_HOUR
        )

    while remaining:
        # Dynamic local scale from CURRENT position only. This avoids the
        # transitive "A near B, B near C, therefore A/C same pocket" problem.
        current_times = sorted(
            float(durations[current][x]) for x in remaining
        )
        nearest_time = current_times[0]
        local_limit = max(
            nearest_time * 2.25,
            nearest_time + 6.0 * 60.0,
        )

        local = [
            x for x in remaining
            if float(durations[current][x]) <= local_limit
        ]
        if not local:
            local = list(remaining)

        def candidate_key(candidate):
            direct = edge_cost(current, candidate)
            future_local = [x for x in local if x != candidate]

            if future_local:
                continuation = min(
                    edge_cost(candidate, x)
                    for x in future_local
                )
            else:
                future = remaining - {candidate}
                continuation = (
                    min(edge_cost(candidate, x) for x in future)
                    if future else edge_cost(candidate, 0)
                )

            return (
                direct + 0.15 * continuation,
                float(durations[current][candidate]),
                float(distances[current][candidate]),
                candidate,
            )

        chosen = min(local, key=candidate_key)
        route.append(chosen)
        remaining.remove(chosen)
        current = chosen

    route.append(0)
    return route






def pocket_preserving_polish(
    route,
    distances,
    durations,
    locations,
    fuel_price,
    mpg,
):
    """Conservative final polish.

    Only accepts a relocation when it improves the real economic objective AND
    does not create a larger local jump around the moved customer than the
    original route. This prevents a global optimiser from undoing the
    driver-style pocket behaviour.
    """
    if not route or len(route) < 5:
        return route[:] if route else route

    best = route[:]
    best_metrics = route_metrics(best, distances, durations, fuel_price, mpg)
    best_score = (
        float(best_metrics["fuel_cost"])
        + (float(best_metrics["time_s"]) / 3600.0)
        * DRIVING_TIME_VALUE_PER_HOUR
    )

    # One conservative deterministic pass is intentional.
    original = best[:]
    for i in range(1, len(original) - 1):
        customer = original[i]
        if customer not in best[1:-1]:
            continue
        old_i = best.index(customer)
        old_prev = best[old_i - 1]
        old_next = best[old_i + 1]
        old_local = (
            float(durations[old_prev][customer])
            + float(durations[customer][old_next])
        )

        shortened = best[:old_i] + best[old_i + 1:]
        chosen = None

        for j in range(1, len(shortened)):
            prev_node = shortened[j - 1]
            next_node = shortened[j]
            new_local = (
                float(durations[prev_node][customer])
                + float(durations[customer][next_node])
            )

            # Do not move a job into a materially more remote local position.
            if new_local > old_local * 1.20 + 120.0:
                continue

            candidate = shortened[:j] + [customer] + shortened[j:]
            metrics = route_metrics(
                candidate, distances, durations, fuel_price, mpg
            )
            score = (
                float(metrics["fuel_cost"])
                + (float(metrics["time_s"]) / 3600.0)
                * DRIVING_TIME_VALUE_PER_HOUR
            )
            key = (
                round(score, 9),
                round(float(metrics["time_s"]), 6),
                round(float(metrics["distance_m"]), 6),
                j,
            )
            if score + 1e-9 < best_score and (
                chosen is None or key < chosen[0]
            ):
                chosen = (key, candidate, metrics)

        if chosen is not None:
            _, best, best_metrics = chosen
            best_score = (
                float(best_metrics["fuel_cost"])
                + (float(best_metrics["time_s"]) / 3600.0)
                * DRIVING_TIME_VALUE_PER_HOUR
            )

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



def intensive_route_polish(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
    max_rounds=8,
):
    """V26.11 best-improvement whole-route search.

    Searches several complementary neighbourhoods against the SAME verified
    ORS matrix and the SAME fuel + driving-time objective as V26.10:
      * 2-opt segment reversals
      * single-customer swaps
      * Or-opt relocation of 1..5 consecutive customers

    The depot is fixed at both ends.  A round accepts only the single best
    strict improvement found across every tested move, then starts again.
    This is intentionally geography-agnostic: it contains no postcode, town,
    specific village, postcode, or other hard-coded ordering rule.
    """
    if not route or len(route) < 5:
        return route[:] if route else route

    customer_count = len(route) - 2
    expected = list(range(1, customer_count + 1))

    def valid(candidate):
        return (
            candidate
            and candidate[0] == 0
            and candidate[-1] == 0
            and len(candidate) == customer_count + 2
            and sorted(candidate[1:-1]) == expected
        )

    best = route[:]
    if not valid(best):
        return route[:]
    best_key = _practical_route_key(
        best, distances, durations, fuel_price, mpg
    )

    for _ in range(max_rounds):
        round_best = None
        round_key = best_key
        n = len(best)

        # 2-opt: reverse a complete internal section.
        for i in range(1, n - 2):
            for j in range(i + 1, n - 1):
                candidate = best[:]
                candidate[i:j + 1] = reversed(candidate[i:j + 1])
                key = _practical_route_key(
                    candidate, distances, durations, fuel_price, mpg
                )
                if key < round_key:
                    round_key = key
                    round_best = candidate

        # Swap two individual customers. This can escape a local minimum that
        # neither a simple relocation nor one 2-opt reversal can improve.
        for i in range(1, n - 2):
            for j in range(i + 1, n - 1):
                candidate = best[:]
                candidate[i], candidate[j] = candidate[j], candidate[i]
                key = _practical_route_key(
                    candidate, distances, durations, fuel_price, mpg
                )
                if key < round_key:
                    round_key = key
                    round_best = candidate

        # Or-opt: relocate short consecutive runs. V26.10 searched up to three
        # jobs; V26.11 extends this to five so a whole small geographic pocket
        # can move together instead of being split across the day.
        for block_size in range(1, min(5, customer_count) + 1):
            for i in range(1, n - block_size):
                if i + block_size > n - 1:
                    continue
                block = best[i:i + block_size]
                remainder = best[:i] + best[i + block_size:]
                for j in range(1, len(remainder)):
                    candidate = remainder[:j] + block + remainder[j:]
                    if candidate == best:
                        continue
                    key = _practical_route_key(
                        candidate, distances, durations, fuel_price, mpg
                    )
                    if key < round_key:
                        round_key = key
                        round_best = candidate

        if round_best is None or not valid(round_best):
            break
        best = round_best
        best_key = round_key

    return best


def deterministic_multistart_polish(
    route,
    distances,
    durations,
    fuel_price,
    mpg,
):
    """V26.13 deterministic multi-start escape search.

    V26.12 showed that simply making one search more aggressive can land on a
    different local optimum without improving the live route.  V26.13 instead
    protects the V26.11 route, creates a small set of fundamentally different
    complete-day starting tours, fully polishes each against the SAME verified
    ORS matrix, and keeps a result only when its complete practical route key
    is strictly better than the protected input route.

    No town, postcode, coordinate, cluster or customer-specific rule is used.
    The number of starts is deliberately capped so daily optimisation remains
    practical in Streamlit.
    """
    if not route or len(route) < 5:
        return route[:] if route else route

    customer_count = len(route) - 2
    expected = list(range(1, customer_count + 1))

    def valid(candidate):
        return (
            candidate
            and candidate[0] == 0
            and candidate[-1] == 0
            and len(candidate) == customer_count + 2
            and sorted(candidate[1:-1]) == expected
        )

    protected = route[:]
    if not valid(protected):
        return protected

    customers = protected[1:-1]
    seeds = [protected]

    def add_seed(order):
        candidate = [0] + list(order) + [0]
        if valid(candidate) and candidate not in seeds:
            seeds.append(candidate)

    # Reverse sweep: useful on asymmetric road networks because A->B need not
    # have exactly the same time/cost as B->A.
    add_seed(reversed(customers))

    # Rotate the complete customer sweep at deterministic cut points.  These
    # starts preserve most good adjacencies while changing where the day enters
    # and leaves the sweep, which can expose a different local optimum.
    n = len(customers)
    cuts = sorted(set(x for x in (n // 4, n // 2, (3 * n) // 4) if 0 < x < n))
    for cut in cuts:
        add_seed(customers[cut:] + customers[:cut])

    # Two broad section reversals create genuinely different basins without
    # randomisation.  Keeping this list small avoids the V26.12 runtime cost.
    if n >= 8:
        mid = n // 2
        add_seed(list(reversed(customers[:mid])) + customers[mid:])
        add_seed(customers[:mid] + list(reversed(customers[mid:])))

    best = protected
    best_key = _practical_route_key(best, distances, durations, fuel_price, mpg)

    # The protected V26.11 route is seed zero. Other seeds get a bounded deep
    # polish.  Nothing is accepted unless the complete final key beats it.
    for seed in seeds[1:]:
        candidate = surgical_route_polish(
            seed, distances, durations, fuel_price, mpg, max_passes=2
        )
        candidate = intensive_route_polish(
            candidate,
            distances,
            durations,
            fuel_price,
            mpg,
            max_rounds=5,
        )
        if not valid(candidate):
            continue
        key = _practical_route_key(candidate, distances, durations, fuel_price, mpg)
        if key < best_key:
            best = candidate
            best_key = key

    return best

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
    # Bank details stay out of the visible app and source code. They come only
    # from private Streamlit Secrets and are inserted into the WhatsApp text.
    account_name = clean_val(st.secrets.get("PAYMENT_ACCOUNT_NAME", ""))
    bank_name = clean_val(st.secrets.get("PAYMENT_BANK_NAME", ""))
    sort_code = clean_val(st.secrets.get("PAYMENT_SORT_CODE", ""))
    account_number = clean_val(st.secrets.get("PAYMENT_ACCOUNT_NUMBER", ""))
    payment_text = "Please use the bank details on your DanCleanUK payment flyer."
    if account_name and bank_name and sort_code and account_number:
        payment_text = (f"Please pay by bank transfer to {account_name} - "
                        f"Bank: {bank_name}, Sort Code: {sort_code}, "
                        f"Account Number: {account_number}.")
    message = (f"Hi from {BUSINESS_NAME}! Your service is complete today. "
               f"Total: £{price:.2f}. {payment_text} Thank you!")
    return "https://wa.me/" + quote(normalise_phone(phone)) + "?text=" + quote(message)


# ============================================================
# LOAD EXISTING DAY
# ============================================================

existing_day = load_day(service_date_str)

# V27.6 reconnect recovery: if Streamlit/Chrome restarted and the permanent
# route exists, rebuild the working day from Supabase automatically. This
# never optimises or changes the saved stop order.
if "master_df" not in st.session_state and existing_day.empty:
    reconnect_snapshot = load_route_snapshot(service_date_str)
    if reconnect_snapshot is not None:
        ok, _ = apply_saved_route_snapshot(service_date_str)
        if ok:
            existing_day = st.session_state.master_df.copy()

if (
    not existing_day.empty
    and "master_df" not in st.session_state
):
    st.session_state.master_df = existing_day.copy()

    # If the laptop has explicitly saved a finished route, restore its exact
    # summary and stop order instead of creating a zero-mile placeholder.
    saved_snapshot = load_route_snapshot(service_date_str)
    if saved_snapshot is not None and "route_data" not in st.session_state:
        apply_saved_route_snapshot(service_date_str)
        existing_day = st.session_state.master_df.copy()

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
    # V27.7.1 safety fix: starting a new working session must NEVER delete
    # either the local day record or the permanent Supabase route snapshot.
    # It only clears the active in-memory route so a different date/day can
    # be selected. Returning to this date will offer/load its saved route.
    for key in [
        "master_df",
        "route_data",
        "failed_jobs",
    ]:
        st.session_state.pop(key, None)
    st.rerun()


# ============================================================
# SAVED ROUTE / PHONE WORK MODE
# ============================================================

saved_snapshot = load_route_snapshot(service_date_str)

if saved_snapshot is not None:
    st.success(
        "📱 A saved route is available for this date. "
        "On your phone/tablet, load it instead of optimising again."
    )
    if st.button(
        "📱 LOAD SAVED ROUTE — NO OPTIMISATION",
        type="primary",
        use_container_width=True,
    ):
        ok, message = apply_saved_route_snapshot(service_date_str)
        if ok:
            st.session_state["saved_route_notice"] = message
            st.rerun()
        else:
            st.error(message)

if st.session_state.pop("saved_route_notice", None):
    st.success("✅ Saved laptop route loaded exactly as stored.")

# Driver Mode is automatic when a locked laptop route has been loaded.
# It changes presentation only; routing and saved-route data stay untouched.
driver_mode = bool(
    st.session_state.get("route_data", {}).get("saved_route")
)

if driver_mode:
    st.markdown(
        """
        <style>
        /* Mobile-safe action rows: columns must share the viewport instead of
           keeping a desktop minimum width and overflowing off-screen. */
        div[data-testid="stHorizontalBlock"] {
            gap: 0.35rem !important;
            width: 100% !important;
        }
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
            flex: 1 1 0 !important;
            width: 0 !important;
            min-width: 0 !important;
            max-width: 100% !important;
        }
        div[data-testid="stButton"],
        div[data-testid="stLinkButton"] {
            width: 100% !important;
            min-width: 0 !important;
        }
        div[data-testid="stButton"] button,
        div[data-testid="stLinkButton"] a {
            width: 100% !important;
            min-width: 0 !important;
            max-width: 100% !important;
            min-height: 2.75rem;
            padding-left: 0.2rem !important;
            padding-right: 0.2rem !important;
            white-space: nowrap !important;
            overflow: hidden !important;
            text-overflow: ellipsis !important;
        }
        @media (max-width: 640px) {
            h1 { font-size: 1.8rem !important; margin-bottom: 0.35rem !important; }
            h2 { font-size: 1.45rem !important; }
            h3 { font-size: 1.25rem !important; }
            div[data-testid="stVerticalBlockBorderWrapper"] { margin-bottom: 0.35rem; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = None
if not driver_mode:
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
            # Notes may be supplied in the daily route spreadsheet. Keep them
            # with the customer from import -> locked route -> phone -> report.
            if "Notes" not in df.columns:
                df["Notes"] = ""
            else:
                df["Notes"] = df["Notes"].fillna("").astype(str).str.strip()
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
                        "Notes",
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
                        "Notes",
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
                df["Notes"] = df["Notes"].fillna("")
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

if (
    not driver_mode
    and st.button(
        "🚀 PLAN / RE-PLAN BEST DAILY ROUTE",
        type="primary",
        use_container_width=True,
    )
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

    if failed_rows:
        st.error(
            f"{len(failed_rows)} customer address(es) could not be verified. "
            "The route has NOT been calculated, because an incomplete or "
            "mislocated daily route would be inaccurate."
        )
        cols = [c for c in ["Address", "Postcode", "Phone"] if c in df.columns]
        st.dataframe(df.loc[failed_rows, cols], use_container_width=True)
        st.info(
            "Correct the address/postcode shown above and calculate again. "
            "The app will not guess a location or substitute straight-line mileage."
        )
        st.stop()

    if not valid_rows:
        st.error("No customer addresses could be located.")
        st.stop()

    # Public geocoders can collapse several houses sharing one postcode to the
    # same centroid. Separate those local houses before building the road matrix.
    df = deconflict_same_postcode_houses(df, valid_rows, depot_coords)

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

    # ------------------------------------------------------------------
    # V26.4 ROUTING ENGINE INPUT
    # ------------------------------------------------------------------
    # Route from the validated house-level coordinates already resolved by
    # get_coords(). Official postcode centroids are kept only as recovery
    # anchors if live-road validation later identifies a suspicious customer.
    #
    # No town/postcode ordering, fixed route, expected mileage or test-day
    # customer rule exists here.

    locations = [[float(depot_coords[1]), float(depot_coords[0])]]
    postcode_fallbacks = [None]

    for node_idx, df_idx in enumerate(valid_rows, start=1):
        row = routing_df.iloc[node_idx]

        # Exact address coordinate produced by the app's existing geocoder,
        # which already validates exact matches against the postcode area.
        locations.append([
            float(row["longitude"]),
            float(row["latitude"]),
        ])

        pc = normalise_postcode(row.get("Postcode", ""))
        anchor = get_postcode_coords(pc)
        if anchor is None:
            # Rows using a verified historic postcode fallback are allowed to
            # use that same official historic point for matrix repair. This is
            # only reached for customers already accepted by get_coords().
            anchor = get_terminated_postcode_coords(pc)

        if anchor is None:
            postcode_fallbacks.append(None)
        else:
            lat, lon = float(anchor[0]), float(anchor[1])
            postcode_fallbacks.append([lon, lat])

    def build_duplicate_groups(route_locations):
        grouped = {}
        for node_idx in range(1, len(route_locations)):
            lon, lat = route_locations[node_idx]
            key = (round(float(lon), 6), round(float(lat), 6))
            grouped.setdefault(key, []).append(node_idx)
        return [
            members
            for members in grouped.values()
            if len(members) > 1
        ]

    duplicate_groups = build_duplicate_groups(locations)

    # First choice: a genuine ORS matrix from validated house coordinates.
    with st.spinner(
        "🛣️ Getting actual road distances and driving times..."
    ):
        distances, durations = get_ors_matrix(locations)

    using_offline = False
    live_matrix_verified = False
    corrected_nodes = []

    if matrix_values_are_valid(distances, durations, len(locations)):
        suspects = suspicious_matrix_nodes(
            locations,
            distances,
            durations,
        )

        # Repair coordinates, NOT road legs. Each suspicious customer is moved
        # only to its official postcode anchor, then the WHOLE road matrix is
        # requested again from ORS. No road distance/time is fabricated.
        if suspects:
            retry_locations = [coord[:] for coord in locations]
            for node_idx in suspects:
                if node_idx <= 0 or node_idx >= len(postcode_fallbacks):
                    continue
                fallback = postcode_fallbacks[node_idx]
                if fallback is None:
                    continue

                old_lon, old_lat = retry_locations[node_idx]
                new_lon, new_lat = fallback
                shift_km = haversine_km(
                    old_lat, old_lon, new_lat, new_lon
                )
                if shift_km >= 0.03:
                    retry_locations[node_idx] = [new_lon, new_lat]
                    corrected_nodes.append(node_idx)

            if corrected_nodes:
                with st.spinner(
                    "🛣️ Rechecking suspicious road locations..."
                ):
                    retry_d, retry_t = get_ors_matrix(retry_locations)

                if matrix_values_are_valid(
                    retry_d, retry_t, len(retry_locations)
                ):
                    retry_suspects = suspicious_matrix_nodes(
                        retry_locations,
                        retry_d,
                        retry_t,
                    )
                    if not retry_suspects:
                        locations = retry_locations
                        distances = retry_d
                        durations = retry_t
                        live_matrix_verified = True

                        labels = []
                        for node_idx in corrected_nodes:
                            row = routing_df.iloc[node_idx]
                            address = clean_val(row.get("Address", ""))
                            postcode = clean_val(row.get("Postcode", ""))
                            labels.append(address or postcode or f"stop {node_idx}")

                        st.info(
                            "Live-road validation safely corrected the routing "
                            "location for: " + " / ".join(labels)
                        )
                # If retry remains suspicious, do not label it verified/live.
            else:
                # Suspicious ORS matrix but no safe coordinate repair available.
                live_matrix_verified = False
        else:
            live_matrix_verified = True

    if not live_matrix_verified:
        # V26.4 production rule: NO straight-line/estimated routing fallback.
        # If genuine ORS road data cannot be verified, stop rather than invent
        # mileage or driving time.
        st.error(
            "A verified real-road route could not be calculated. "
            "No estimated or straight-line route has been substituted."
        )
        st.info(
            "Check the customer addresses/postcodes shown above, then calculate "
            "the route again. The app will only continue when verified ORS "
            "driving-road distances and times are available."
        )
        st.stop()

    using_offline = False

    # Authoritative duplicate groups are based on the exact coordinates used
    # by this route matrix.
    duplicate_groups = build_duplicate_groups(locations)

    with st.spinner(
        f"🧠 Optimising {len(valid_rows)} customer stops..."
    ):
        # V26.10 hybrid search.  Build the conservative driver-style route AND
        # the complete-day route, then compare them on the same verified ORS
        # matrix.  This removes the V26.9 red flag where the final route could
        # preserve a local pocket even when it caused a large later backtrack.
        route_candidates = []

        pocket_route = build_daily_pocket_route(
            distances,
            durations,
            locations,
            FUEL_PRICE,
            MPG,
        )
        if pocket_route:
            pocket_route = pocket_preserving_polish(
                pocket_route,
                distances,
                durations,
                locations,
                FUEL_PRICE,
                MPG,
            )
            route_candidates.append(pocket_route)

            # Let the whole-route polish test whether a job trapped in the
            # wrong pocket can be relocated without changing any coordinates.
            route_candidates.append(
                surgical_route_polish(
                    pocket_route,
                    distances,
                    durations,
                    FUEL_PRICE,
                    MPG,
                    max_passes=4,
                )
            )

            route_candidates.append(
                intensive_route_polish(
                    route_candidates[-1],
                    distances,
                    durations,
                    FUEL_PRICE,
                    MPG,
                    max_rounds=8,
                )
            )

        global_route = optimise_route(
            distances,
            durations,
            FUEL_PRICE,
            MPG,
            locations=locations,
        )
        if global_route:
            route_candidates.append(global_route)
            route_candidates.append(
                surgical_route_polish(
                    global_route,
                    distances,
                    durations,
                    FUEL_PRICE,
                    MPG,
                    max_passes=4,
                )
            )

            # V26.11: deeper best-improvement search from both the raw global
            # route and its V26.10 polished result.  The final selector below
            # still keeps V26.10's candidate whenever these do not beat it.
            route_candidates.append(
                intensive_route_polish(
                    global_route,
                    distances,
                    durations,
                    FUEL_PRICE,
                    MPG,
                    max_rounds=8,
                )
            )
            route_candidates.append(
                intensive_route_polish(
                    route_candidates[-2],
                    distances,
                    durations,
                    FUEL_PRICE,
                    MPG,
                    max_rounds=8,
                )
            )

        # Reject any malformed candidate before comparison. Depot stays fixed,
        # every customer must appear exactly once, and no job may disappear.
        expected_customers = list(range(1, len(distances)))
        safe_candidates = []
        seen_candidates = set()
        for candidate in route_candidates:
            if not candidate:
                continue
            if candidate[0] != 0 or candidate[-1] != 0:
                continue
            if len(candidate) != len(distances) + 1:
                continue
            if sorted(candidate[1:-1]) != expected_customers:
                continue
            candidate_tuple = tuple(candidate)
            if candidate_tuple in seen_candidates:
                continue
            seen_candidates.add(candidate_tuple)
            safe_candidates.append(candidate)

        route = (
            min(
                safe_candidates,
                key=lambda r: _practical_route_key(
                    r, distances, durations, FUEL_PRICE, MPG
                ),
            )
            if safe_candidates
            else None
        )

        # V26.13: keep the complete V26.11 winner protected, then try a small
        # deterministic set of genuinely different starts.  The multi-start
        # result can replace it only when the SAME verified ORS matrix and the
        # SAME complete-route objective say it is strictly better.
        if route:
            route = deterministic_multistart_polish(
                route,
                distances,
                durations,
                FUEL_PRICE,
                MPG,
            )

    if not route:
        st.error("The route optimiser could not create a safe complete route.")
        st.stop()

    # Keep same-coordinate jobs consecutive without globally relocating
    # their area. The first occurrence determines the group's position.
    if duplicate_groups:
        group_for = {}
        for gid, group in enumerate(duplicate_groups):
            for node in group:
                group_for[node] = gid

        emitted = set()
        grouped_route = [route[0]]

        for node in route[1:-1]:
            gid = group_for.get(node)
            if gid is None:
                grouped_route.append(node)
                continue
            if gid in emitted:
                continue

            members = [x for x in route[1:-1] if group_for.get(x) == gid]
            grouped_route.extend(members)
            emitted.add(gid)

        grouped_route.append(route[-1])

        # Safety: grouping must never lose or duplicate a customer.
        if (
            grouped_route[0] == 0
            and grouped_route[-1] == 0
            and len(grouped_route) == len(route)
            and sorted(grouped_route[1:-1]) == sorted(route[1:-1])
            and _practical_route_key(
                grouped_route, distances, durations, FUEL_PRICE, MPG
            ) <= _practical_route_key(
                route, distances, durations, FUEL_PRICE, MPG
            )
        ):
            # V26.10: never force duplicate grouping if it would make the
            # verified complete route worse.
            route = grouped_route

    metrics = route_metrics(
        route,
        distances,
        durations,
        FUEL_PRICE,
        MPG,
    )

    # V26.2 final-leg audit. This makes any remaining bad leg visible instead
    # of hiding it inside the day's total.
    leg_audit = route_leg_diagnostics(
        route,
        locations,
        distances,
        durations,
    )

    with st.expander("🔎 Routing diagnostics", expanded=False):
        diagnostic_rows = []
        for item in leg_audit:
            a = item["from_node"]
            b = item["to_node"]

            def node_label(node_idx):
                if node_idx == 0:
                    return "DEPOT"
                row = routing_df.iloc[node_idx]
                address = clean_val(row.get("Address", ""))
                postcode = clean_val(row.get("Postcode", ""))
                return f"{address} ({postcode})" if address else postcode

            diagnostic_rows.append({
                "Leg": item["leg"],
                "From": node_label(a),
                "To": node_label(b),
                "Road miles": round(item["road_miles"], 2),
                "Drive min": round(item["minutes"], 1),
                "Straight miles": round(item["straight_miles"], 2),
                "Road / straight": (
                    round(item["ratio"], 2)
                    if item["ratio"] is not None else ""
                ),
            })

        st.dataframe(
            pd.DataFrame(diagnostic_rows),
            use_container_width=True,
            hide_index=True,
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

    if driver_mode:
        st.markdown(
            f"""
            <div style="border:1px solid rgba(128,128,128,.35);border-radius:12px;
                        padding:.75rem .9rem;margin:.25rem 0 .65rem 0;">
              <div style="font-size:1.05rem;font-weight:700;margin-bottom:.45rem;">📱 Driver Mode</div>
              <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:.45rem;text-align:center;">
                <div><b>{max(total_jobs - completed_jobs, 0)}</b><br><span style="font-size:.78rem;opacity:.75">Remaining</span></div>
                <div><b>{route_data['miles']:.1f} mi</b><br><span style="font-size:.78rem;opacity:.75">Distance</span></div>
                <div><b>{format_duration(route_data['time'])}</b><br><span style="font-size:.78rem;opacity:.75">Driving</span></div>
                <div><b>{completed_jobs}</b><br><span style="font-size:.78rem;opacity:.75">Done</span></div>
                <div><b>{paid_jobs}</b><br><span style="font-size:.78rem;opacity:.75">Paid</span></div>
                <div><b>£{route_data['take_home']:.2f}</b><br><span style="font-size:.78rem;opacity:.75">Take-home</span></div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        c1, c2 = st.columns(2)
        with c1:
            st.metric("Take-Home", f"£{route_data['take_home']:.2f}")
        with c2:
            st.metric("Revenue", f"£{route_data['revenue']:.2f}")

        c3, c4 = st.columns(2)
        with c3:
            st.metric("Driving Distance", f"{route_data['miles']:.1f} miles")
        with c4:
            st.metric("Driving Time", format_duration(route_data["time"]))

        c5, c6 = st.columns(2)
        with c5:
            st.metric("Fuel Cost", f"£{route_data['fuel_cost']:.2f}")
        with c6:
            st.metric("Fuel Used", f"{route_data['litres']:.1f} litres")

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
            "✅ Route calculated using verified ORS driving-road distance and driving time."
        )


# ============================================================
# SAVE FINISHED ROUTE FOR PHONE / TABLET
# ============================================================

if route_data and not route_data.get("persisted_only") and not driver_mode:
    route_ready_to_save = (
        route_data.get("jobs", 0) > 0
        and df["route_order"].notna().sum() >= int(route_data.get("jobs", 0))
    )
    if st.button(
        "💾 SAVE / LOCK THIS ROUTE FOR PHONE",
        use_container_width=True,
        disabled=not route_ready_to_save,
    ):
        if save_route_snapshot(service_date_str, df, route_data):
            st.success(
                "✅ Route locked. Open the same date on your phone/tablet and "
                "press LOAD SAVED ROUTE — NO OPTIMISATION."
            )
        else:
            st.error("The route could not be locked safely. No saved route was changed.")

    if route_data.get("saved_route"):
        st.info(
            "🔒 You are using the saved route. No optimisation was run on this device."
        )

if driver_mode:
    st.info("🔒 Saved laptop route · no optimisation on this device")

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
# WORKING-DAY PROGRESS
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

# V27.6: navigation lives on each customer card. The duplicate sidebar
# navigation/next-stop block was intentionally removed.


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
            "notes",
            "_route_sort",
        }

        for column in row.index:
            column_key = str(column).lower()
            if column_key in ignored:
                continue
            if driver_mode and column_key in {"address", "dates", "date"}:
                continue

            value = clean_val(row.get(column))
            if value:
                extra.append(
                    f"**{column}:** {value}"
                )

        with st.container(border=True):
            street_address = clean_val(row.get("Address")) or clean_val(row.get("address_text"))
            if street_address:
                st.write(f"### {icon} STOP {display_number} — {street_address}")
                if postcode:
                    st.write(f"**{postcode}**")
            else:
                st.write(f"### {icon} STOP {display_number} — {postcode}")

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
                    payment_selected = payment in {"Cash", "Bank Transfer", "Not Paid"}
                    if st.button(
                        "✅ Complete",
                        key=f"complete_{row['job_id']}",
                        use_container_width=True,
                        disabled=not payment_selected,
                        help=None if payment_selected else "Choose Cash, Bank or Unpaid first.",
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
                        current_route_data = st.session_state.get("route_data")
                        if current_route_data and current_route_data.get("saved_route"):
                            save_route_snapshot(service_date_str, df, current_route_data)
                        st.rerun()
                else:
                    st.success("Completed")

            with col2:
                st.link_button(
                    "🚗 Navigate",
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
                    current_route_data = st.session_state.get("route_data")
                    if current_route_data and current_route_data.get("saved_route"):
                        save_route_snapshot(service_date_str, df, current_route_data)
                    st.rerun()

            with pay2:
                if st.button(
                    "🏦 Bank",
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
                    df.at[master_idx, "WhatsAppSent"] = False
                    df.at[master_idx, "WhatsAppTime"] = ""

                    save_job(df.loc[master_idx])
                    st.session_state.master_df = df
                    current_route_data = st.session_state.get("route_data")
                    if current_route_data and current_route_data.get("saved_route"):
                        save_route_snapshot(service_date_str, df, current_route_data)
                    st.rerun()

            with pay3:
                if st.button(
                    "❌ Unpaid",
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
                    df.at[master_idx, "WhatsAppSent"] = False
                    df.at[master_idx, "WhatsAppTime"] = ""

                    save_job(df.loc[master_idx])
                    st.session_state.master_df = df
                    current_route_data = st.session_state.get("route_data")
                    if current_route_data and current_route_data.get("saved_route"):
                        save_route_snapshot(service_date_str, df, current_route_data)
                    st.rerun()

            if payment in {"Bank Transfer", "Not Paid"}:
                whatsapp_sent = bool(row.get("WhatsAppSent", False))
                if phone:
                    st.link_button("💬 1. Open WhatsApp Message", whatsapp_url(phone, price), use_container_width=True)
                    if not whatsapp_sent:
                        if st.button("✅ 2. Confirm WhatsApp Sent", key=f"whatsapp_sent_{row['job_id']}", use_container_width=True):
                            master_idx = df.index[df["job_id"].astype(str) == str(row["job_id"])][0]
                            df.at[master_idx, "WhatsAppSent"] = True
                            df.at[master_idx, "WhatsAppTime"] = now_text()
                            st.session_state.master_df = df
                            current_route_data = st.session_state.get("route_data")
                            if current_route_data and current_route_data.get("saved_route"):
                                save_route_snapshot(service_date_str, df, current_route_data)
                            st.rerun()
                    else:
                        st.success("💬 WhatsApp confirmed sent — Complete is unlocked.")
                else:
                    st.warning("No phone number is stored for this customer, so WhatsApp cannot be sent.")

            current_notes = clean_val(row.get("Notes"))
            edit_key = f"edit_notes_{row['job_id']}"
            if not st.session_state.get(edit_key, False):
                st.markdown("**📝 Job notes 🔒**")
                st.caption(current_notes if current_notes else "No notes for this job.")
                if st.button(
                    "✏️ Edit Notes",
                    key=f"open_notes_{row['job_id']}",
                    use_container_width=True,
                ):
                    st.session_state[edit_key] = True
                    st.rerun()
            else:
                notes_value = st.text_area(
                    "📝 Edit job notes",
                    value=current_notes,
                    key=f"notes_{row['job_id']}",
                    placeholder="e.g. Full house, front + back, gate code, access details…",
                    height=70,
                )
                if st.button(
                    "💾 Save / Lock Notes 🔒",
                    key=f"save_notes_{row['job_id']}",
                    use_container_width=True,
                ):
                    master_idx = df.index[df["job_id"].astype(str) == str(row["job_id"])][0]
                    df.at[master_idx, "Notes"] = notes_value.strip()
                    save_job(df.loc[master_idx])
                    st.session_state.master_df = df
                    current_route_data = st.session_state.get("route_data")
                    if current_route_data and current_route_data.get("saved_route"):
                        save_route_snapshot(service_date_str, df, current_route_data)
                    st.session_state[edit_key] = False
                    st.rerun()

with st.container(border=True):
    st.write("### 🏁 FINISH — GRANTHAM DEPOT")
    st.write(DEPOT_FULL_ADDRESS)

    st.link_button(
        "🚗 Navigate Back to Depot",
        maps_url(DEPOT_FULL_ADDRESS),
        use_container_width=True,
        disabled=route_df.empty,
    )


# ============================================================
# COMPLETED / PAYMENT SUMMARY
# ============================================================

completed_df = df[
    df["Status"].astype(str).str.lower() == "completed"
].copy()

if not completed_df.empty and not driver_mode:
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


def save_report_to_snapshot(service_date, report_bytes, filename):
    snapshot = load_route_snapshot(service_date)
    if not snapshot:
        return False
    data = dict(snapshot.get("route_data") or {})
    data["report_b64"] = base64.b64encode(report_bytes).decode("ascii")
    data["report_filename"] = filename
    data["report_saved_at"] = now_text()
    url, _ = _supabase_config()
    headers = _supabase_headers("return=minimal")
    if not url or not headers:
        return False
    try:
        r = requests.patch(f"{url}/rest/v1/saved_routes", headers=headers, params={"route_date": f"eq.{service_date}"}, json={"route_data": data, "updated_at": datetime.now().astimezone().isoformat()}, timeout=20)
        return r.status_code in (200, 204)
    except requests.RequestException:
        return False


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
        "address_text",
        "Postcode",
        "Price",
        "Phone",
        "Status",
        "Payment",
        "PaymentTime",
        "CompletedTime",
        "Notes",
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

    report_ws = workbook.active
    report_ws.title = "Daily Report"

    header_fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    report_df = export_df.copy()
    if "address_text" in report_df.columns:
        report_df = report_df.rename(columns={"address_text": "Address"})
    if "Payment" in report_df.columns:
        report_df = report_df.rename(columns={"Payment": "Payment Method"})
    if "Status" in report_df.columns:
        report_df = report_df.rename(columns={"Status": "Completion Status"})

    visible_first = ["Address", "Postcode", "Price", "Payment Method", "Completion Status", "Notes", "PaymentTime", "CompletedTime", "Phone", "route_order"]
    first = [c for c in visible_first if c in report_df.columns]
    rest = [c for c in report_df.columns if c not in first and c != "job_id"]
    report_df = report_df[first + rest]

    for values in dataframe_to_rows(report_df, index=False, header=True):
        report_ws.append(values)
    for cell in report_ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    total_revenue = float(export_df["Price"].sum()) if not export_df.empty else 0
    completed_count = int(export_df["Status"].astype(str).str.lower().eq("completed").sum()) if not export_df.empty else 0
    cash_total = float(export_df.loc[export_df["Payment"].astype(str).str.lower() == "cash", "Price"].sum()) if not export_df.empty else 0
    bank_total = float(export_df.loc[export_df["Payment"].astype(str).str.lower() == "bank transfer", "Price"].sum()) if not export_df.empty else 0
    unpaid_total = float(export_df.loc[export_df["Payment"].astype(str).str.lower().isin(["waiting", "not paid"]), "Price"].sum()) if not export_df.empty else 0

    summary_start = report_ws.max_row + 3
    summary_rows = [
        ["DanCleanUK Daily Summary", ""], ["Route Date", service_date_str], ["Depot", DEPOT_FULL_ADDRESS],
        ["Total Jobs", len(export_df)], ["Completed Jobs", completed_count], ["Revenue", total_revenue],
        ["Cash Received", cash_total], ["Bank Transfer Received", bank_total], ["Outstanding / Unpaid", unpaid_total],
        ["Fuel Cost", route_data["fuel_cost"] if route_data else 0], ["Fuel Used (litres)", route_data["litres"] if route_data else 0],
        ["Driving Miles", route_data["miles"] if route_data else 0], ["Driving Time", format_duration(route_data["time"]) if route_data else "0m"],
        ["Tax Rate", f"{TAX_RATE * 100:.0f}%"], ["Estimated Take-Home", route_data["take_home"] if route_data else 0],
    ]
    for offset, values in enumerate(summary_rows):
        row_num = summary_start + offset
        report_ws.cell(row=row_num, column=1, value=values[0])
        report_ws.cell(row=row_num, column=2, value=values[1])
    report_ws.cell(row=summary_start, column=1).fill = header_fill
    report_ws.cell(row=summary_start, column=1).font = header_font

    for column_cells in report_ws.columns:
        max_length = 0
        letter = column_cells[0].column_letter
        for cell in column_cells:
            try:
                max_length = max(max_length, len(str(cell.value or "")))
            except Exception:
                pass
        report_ws.column_dimensions[letter].width = min(max(max_length + 2, 12), 50)

    workbook.save(output)
    report_bytes = output.getvalue()
    report_filename = f"DanCleanUK_{service_date_str}.xlsx"
    if st.sidebar.button("💾 Get Report / Save for Laptop", use_container_width=True):
        if save_report_to_snapshot(service_date_str, report_bytes, report_filename):
            st.sidebar.success("Report saved permanently — available from your laptop for this date.")
        else:
            st.sidebar.error("Could not save the report permanently. You can still download it on this device.")
    saved = load_route_snapshot(service_date_str)
    if saved and saved.get("report_b64"):
        try:
            st.sidebar.download_button("💻 Download Saved Report", base64.b64decode(saved["report_b64"]), saved.get("report_filename") or report_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        except Exception:
            pass
    st.sidebar.download_button("⬇️ Download Report on This Device", report_bytes, report_filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

st.sidebar.caption(f"DanCleanUK Route Optimizer v{APP_VERSION}")
