import io
import base64
import json
import hashlib
import math
import re
import sqlite3
import time
from datetime import datetime, date, timezone
from zoneinfo import ZoneInfo
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

APP_VERSION = "27.8.8.4.7-REPORT-SIDEBAR-CLEANUP"
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
            cleaning_plan TEXT NOT NULL DEFAULT '',
            next_cleaning_due TEXT,
            created_at TEXT NOT NULL
        )
        """
    )
    columns = {row[1] for row in conn.execute("PRAGMA table_info(jobs)").fetchall()}
    if "notes" not in columns:
        conn.execute("ALTER TABLE jobs ADD COLUMN notes TEXT NOT NULL DEFAULT ''")
    if "cleaning_plan" not in columns:
        conn.execute("ALTER TABLE jobs ADD COLUMN cleaning_plan TEXT NOT NULL DEFAULT ''")
    if "next_cleaning_due" not in columns:
        conn.execute("ALTER TABLE jobs ADD COLUMN next_cleaning_due TEXT")
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
            geo_query, notes, cleaning_plan, next_cleaning_due, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            notes=excluded.notes,
            cleaning_plan=excluded.cleaning_plan,
            next_cleaning_due=excluded.next_cleaning_due
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
            str(row.get("Cleaning Plan", row.get("cleaning_plan", "")) or ""),
            clean_optional(row.get("Next Cleaning Due", row.get("next_cleaning_due"))),
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
            "cleaning_plan": "Cleaning Plan",
            "next_cleaning_due": "Next Cleaning Due",
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
        "Cleaning Plan": "",
        "Next Cleaning Due": "",
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
    loaded["Cleaning Plan"] = loaded["Cleaning Plan"].fillna("")
    loaded["Next Cleaning Due"] = loaded["Next Cleaning Due"].fillna("")

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


def load_persistent_geocode(query, postcode):
    """Load the permanently pinned coordinate for one normalized address."""
    url, _ = _supabase_config()
    headers = _supabase_headers()
    if not url or not headers:
        return None
    address_key = cache_key_for(query, postcode)
    try:
        response = requests.get(
            f"{url}/rest/v1/geocode_registry",
            headers=headers,
            params={
                "address_key": f"eq.{address_key}",
                "select": "latitude,longitude,source",
                "limit": "1",
            },
            timeout=12,
        )
        if response.status_code != 200:
            return None
        rows = response.json() or []
        if not rows:
            return None
        row = rows[0]
        source = str(row.get("source") or "").strip().lower()
        # V27.8.4: legacy postcode-centroid rows are not address-level pins.
        # Ignore them so the normal exact/street geocoders get another chance.
        if source in {"postcode_anchor", "terminated_postcode_anchor"}:
            return None
        return (float(row["latitude"]), float(row["longitude"]))
    except (requests.RequestException, ValueError, TypeError, KeyError):
        return None


def _street_registry_postcode(postcode):
    """Canonical postcode used only by the permanent street registry."""
    compact = re.sub(r"[^A-Z0-9]", "", clean_val(postcode).upper())
    if 5 <= len(compact) <= 7:
        return f"{compact[:-3]} {compact[-3:]}"
    return normalise_postcode(postcode)


def _street_registry_query(street, postcode):
    """Stable synthetic key for a verified street + postcode routing point."""
    street_norm = re.sub(r"[^a-z0-9]+", " ", clean_val(street).lower()).strip()
    postcode_norm = _street_registry_postcode(postcode)
    if not street_norm or not postcode_norm:
        return ""
    return f"__street__:{street_norm}|{postcode_norm}"


def load_persistent_street_geocode(street, postcode):
    """Reuse a street-level point that was verified once and saved permanently."""
    registry_query = _street_registry_query(street, postcode)
    postcode_norm = _street_registry_postcode(postcode)
    if not registry_query or not postcode_norm:
        return None
    return load_persistent_geocode(registry_query, postcode_norm)


def save_persistent_street_geocode(street, postcode, coords, source="verified_street"):
    """Persist a verified street point without pretending it is a house pin."""
    registry_query = _street_registry_query(street, postcode)
    postcode_norm = _street_registry_postcode(postcode)
    if not registry_query or not postcode_norm:
        return False
    return save_persistent_geocode(registry_query, postcode_norm, coords, source)


def save_persistent_geocode(query, postcode, coords, source):
    """Pin a verified address coordinate in Supabase so reruns/devices agree."""
    if coords is None:
        return False
    url, _ = _supabase_config()
    headers = _supabase_headers("resolution=merge-duplicates,return=minimal")
    if not url or not headers:
        return False
    try:
        lat, lon = float(coords[0]), float(coords[1])
        payload = {
            "address_key": cache_key_for(query, postcode),
            "query_text": str(query or "").strip(),
            "postcode": normalise_postcode(postcode),
            "latitude": lat,
            "longitude": lon,
            "source": str(source or "verified"),
            "updated_at": datetime.now().astimezone().isoformat(),
        }
        response = requests.post(
            f"{url}/rest/v1/geocode_registry?on_conflict=address_key",
            headers=headers,
            json=payload,
            timeout=12,
        )
        return response.status_code in (200, 201, 204, 409)
    except (requests.RequestException, ValueError, TypeError):
        return False


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

    # V27.8 lock protection: a newly optimised route must never silently
    # replace an existing permanent route. Updates are allowed only when the
    # exact saved route has first been loaded (progress/payment/notes sync).
    existing_snapshot = load_route_snapshot(service_date)
    if (
        existing_snapshot is not None
        and not route_data.get("saved_route", False)
        and not st.session_state.get("replace_saved_route_allowed", False)
    ):
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
        "Cleaning Plan": "", "Next Cleaning Due": "",
        "WhatsAppSent": False, "WhatsAppTime": "",
        "MessageOpened": False, "MessageConfirmed": False, "MessageConfirmedTime": "",
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


def normalise_cleaning_plan(value):
    """Return a clean calendar-month plan such as '1 Month' or '3 Months'."""
    text = clean_val(value).strip()
    if not text:
        return ""
    match = re.fullmatch(r"(\d+)\s*(?:month|months|m)", text, flags=re.IGNORECASE)
    if not match:
        return ""
    months = int(match.group(1))
    if months < 1 or months > 24:
        return ""
    return f"{months} Month" if months == 1 else f"{months} Months"


def calculate_next_cleaning_due(service_date_value, cleaning_plan):
    """Calculate the next due date using calendar months, never fixed 30-day blocks."""
    plan = normalise_cleaning_plan(cleaning_plan)
    if not plan:
        return ""
    try:
        months = int(plan.split()[0])
        service = pd.Timestamp(str(service_date_value)).normalize()
        return (service + pd.DateOffset(months=months)).date().isoformat()
    except Exception:
        return ""


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
    st.session_state.pop("failed_jobs", None)
    # Leaving a reset/new-day screen for another date starts a clean session.
    st.session_state.pop("start_new_day_mode", None)
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

# Keep business identity internal; no pointless editable customer-message box in the sidebar.
BUSINESS_NAME = "DanCleanUK"

# V25.49 uses one consistent whole-day economic objective.
# There are deliberately no separate time/distance/nearby sliders.
# Those competing filters could override the route search itself.
DRIVING_TIME_VALUE_PER_HOUR = 6.0

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



def nominatim_nearby_street_geocode(street, postcode_anchor, headers):
    """Resolve a real street near the supplied UK postcode anchor.

    This is a transparent STREET-level fallback for genuine addresses when
    public map data has no house-number point. It never pretends the postcode
    centroid is the property and never invents an offset for a house number.
    """
    street = clean_val(street).strip()
    if not street or postcode_anchor is None:
        return None

    try:
        anchor_lat, anchor_lon = float(postcode_anchor[0]), float(postcode_anchor[1])
    except Exception:
        return None

    # Roughly a 2 km box around the official postcode point. This prevents a
    # same-named road elsewhere in Britain from being accepted.
    lat_pad = 0.018
    lon_pad = 0.030
    viewbox = f"{anchor_lon-lon_pad},{anchor_lat+lat_pad},{anchor_lon+lon_pad},{anchor_lat-lat_pad}"

    try:
        response = requests.get(
            "https://nominatim.openstreetmap.org/search",
            params={
                "q": street,
                "format": "jsonv2",
                "addressdetails": 1,
                "limit": 20,
                "countrycodes": "gb",
                "viewbox": viewbox,
                "bounded": 1,
            },
            headers=headers,
            timeout=20,
        )
        if response.status_code != 200:
            return None

        wanted = re.sub(r"[^a-z0-9]+", " ", street.lower()).strip()
        for item in response.json() or []:
            address = item.get("address") or {}
            returned = clean_val(
                address.get("road")
                or address.get("pedestrian")
                or address.get("residential")
                or item.get("display_name")
            ).lower()
            returned_norm = re.sub(r"[^a-z0-9]+", " ", returned).strip()
            if wanted not in returned_norm:
                continue
            try:
                lat, lon = float(item["lat"]), float(item["lon"])
            except Exception:
                continue
            if haversine_km(anchor_lat, anchor_lon, lat, lon) <= 2.5:
                return (lat, lon)
    except Exception:
        return None

    return None

def get_coords(query_string, postcode, allow_postcode_fallback=False):
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

    # The depot has its own protected start/finish geocoding path. Never let
    # a customer street-registry point replace the depot coordinate after a
    # reboot, even when the depot shares the same street and postcode.
    is_depot = (
        normalise_postcode(postcode) == normalise_postcode(DEPOT_POSTCODE)
        and str(query).strip().casefold() == str(DEPOT_FULL_ADDRESS).strip().casefold()
    )

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

    # V27.8.3: permanent address pinning. Once a verified address has been
    # resolved, every device/reboot uses the same coordinate instead of asking
    # public geocoders to choose again. The coordinate is still revalidated
    # under the current postcode safety rules before use.
    persistent = load_persistent_geocode(query, postcode)
    if persistent is not None:
        checked_persistent = safe_exact(persistent)
        if checked_persistent is not None:
            st.session_state.geocode_cache[key] = checked_persistent
            return checked_persistent

    # V27.8.8.3 PERMANENT STREET REGISTRY:
    # After checking for a permanent exact house pin, if this street + postcode was
    # already verified on an earlier run, reuse that verified STREET point.
    # This prevents a known customer street from depending on public geocoder
    # availability every morning. It remains explicitly street-level: no fake
    # house offset and no postcode-centre substitution.
    if expected_house and expected_street and postcode_anchor is not None and not is_depot:
        persistent_street = load_persistent_street_geocode(expected_street, postcode)
        checked_street = safe_exact(persistent_street)
        if checked_street is not None:
            st.session_state.geocode_cache[key] = checked_street
            approx = st.session_state.setdefault("approximate_geocodes", {})
            approx[key] = {
                "query": query,
                "postcode": postcode,
                "level": "street",
            }
            return checked_street

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
        save_persistent_geocode(query, postcode, exact_ors, "ors_exact")
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
            save_persistent_geocode(query, postcode, coords, "nominatim_exact")
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
        save_persistent_geocode(query, postcode, photon, "photon_exact")
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
                save_persistent_geocode(query, postcode, street_coords, "nominatim_street")
                save_persistent_street_geocode(
                    expected_street, postcode, street_coords, "nominatim_verified_street"
                )
                return street_coords
            time.sleep(0.35)

    # V27.8.7 UK STREET GEO:
    # Many valid UK residential addresses are absent as individual house points
    # from free public geocoders. Before rejecting a genuine house, resolve the
    # NAMED STREET inside a tight box around the official postcode coordinate.
    # This is explicitly street-level: it is not persisted as a house pin and
    # the UI warns when it is used.
    # Keep the depot distinct from nearby customer street clusters. The depot
    # is the fixed start/finish and may use its verified postcode anchor if its
    # individual building point is unavailable. Customer houses on the same
    # street may legitimately share one verified street-level routing point;
    # we do NOT invent house-number offsets.
    if expected_house and expected_street and postcode_anchor is not None and not allow_postcode_fallback and not is_depot:
        nearby_street = nominatim_nearby_street_geocode(
            expected_street, postcode_anchor, headers
        )
        nearby_street = safe_exact(nearby_street)
        if nearby_street is not None:
            st.session_state.geocode_cache[key] = nearby_street
            approx = st.session_state.setdefault("approximate_geocodes", {})
            approx[key] = {
                "query": query,
                "postcode": postcode,
                "level": "street",
            }
            save_persistent_street_geocode(
                expected_street, postcode, nearby_street, "nominatim_verified_street"
            )
            return nearby_street

    # V27.8.5 STRICT GEO:
    # If the imported record contains a house number + street, NEVER silently
    # turn that customer into a postcode centroid. The postcode remains a
    # validation anchor only. If exact/street geocoding could not resolve the
    # supplied address, leave the customer unlocated so the UI can flag it.
    # This prevents a plausible-looking but wrong postcode-centre coordinate
    # from entering the ORS matrix.
    # The depot is not an imported customer record. If public house/street
    # geocoders cannot resolve it, allow the verified live postcode point as
    # the routing start/finish fallback so STRICT GEO cannot block the whole
    # day before customer validation begins. Customer jobs remain strict.
    if expected_house and expected_street and not is_depot and not allow_postcode_fallback:
        st.session_state.geocode_cache.pop(key, None)
        return None

    # Postcode-only records (no usable house + street supplied) may still use
    # the official live postcode coordinate because there is no more precise
    # customer location in the imported data. This is a deliberate fallback,
    # not an address-level pin, so it is never written to geocode_registry.
    if postcode_anchor is not None:
        st.session_state.geocode_cache[key] = postcode_anchor
        return postcode_anchor

    # Likewise, a terminated postcode is only a last-resort coordinate for a
    # postcode-only record. It is never treated as a resolved house address and
    # is never persisted as one.
    if terminated_postcode_anchor is not None and (not expected_street or allow_postcode_fallback):
        # Synthetic benchmark mode may use the terminated-postcode anchor too.
        # Production customer runs remain strict because allow_postcode_fallback=False.
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





def _diagnostic_sha256(value):
    """Stable SHA-256 for diagnostic comparison only; never affects routing."""
    payload = json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _diagnostic_matrix_payload(matrix):
    """Convert ORS matrix values to plain floats without changing them."""
    return [[float(value) for value in row] for row in matrix]


def build_route_input_diagnostic(routing_df, locations, distances, durations):
    """Capture the exact inputs presented to the protected optimiser."""
    jobs = []
    for node_idx in range(len(routing_df)):
        row = routing_df.iloc[node_idx]
        lon, lat = locations[node_idx]
        jobs.append({
            "node": int(node_idx),
            "job_id": clean_val(row.get("job_id", "")),
            "address": clean_val(row.get("Address", row.get("address_text", ""))),
            "postcode": clean_val(row.get("Postcode", "")),
            "latitude": float(lat),
            "longitude": float(lon),
        })

    coord_payload = [[float(point[0]), float(point[1])] for point in locations]
    distance_payload = _diagnostic_matrix_payload(distances)
    duration_payload = _diagnostic_matrix_payload(durations)

    coordinate_fingerprint = _diagnostic_sha256(coord_payload)
    distance_fingerprint = _diagnostic_sha256(distance_payload)
    duration_fingerprint = _diagnostic_sha256(duration_payload)

    input_payload = {
        "jobs": jobs,
        "locations_lon_lat": coord_payload,
        "distance_matrix_sha256": distance_fingerprint,
        "duration_matrix_sha256": duration_fingerprint,
    }

    return {
        "app_version": APP_VERSION,
        "captured_at": datetime.now().astimezone().isoformat(),
        "jobs": jobs,
        "locations_lon_lat": coord_payload,
        "coordinate_sha256": coordinate_fingerprint,
        "distance_matrix_sha256": distance_fingerprint,
        "duration_matrix_sha256": duration_fingerprint,
        "optimizer_input_sha256": _diagnostic_sha256(input_payload),
    }


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

        # A structurally valid ORS matrix with no suspicious nodes is the
        # successful result. Do not fall through into the repair-failure path.
        if not suspects:
            return distances, durations, working_locations, repaired

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
            "Order": report_number,
            "Address": address,
            "Postcode": clean_val(r.get("Postcode")),
            "Price": float(r.get("Price", 0) or 0),
            "Cleaning Plan": normalise_cleaning_plan(r.get("Cleaning Plan")),
            "Phone": clean_val(r.get("Phone")),
            "Payment Method": payment_method,
            "Payment Status": payment_status,
            "Payment Time": report_time(r.get("PaymentTime")) if payment_status == "Paid" else "",
            "Message": message_value,
            "Sent Time": sent_time,
            "Service Date": report_date(r.get("service_date") or service_date_str),
            "Notes": clean_val(r.get("Notes")),
            "Next Cleaning Due": report_date(r.get("Next Cleaning Due")),
        })

    report_df = pd.DataFrame(rows)
    output = io.BytesIO()
    workbook = Workbook()
    report_ws = workbook.active
    report_ws.title = "Daily Report"

    header_fill = PatternFill(start_color="2F4F4F", end_color="2F4F4F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for values in dataframe_to_rows(report_df, index=False, header=True):
        report_ws.append(values)
    for cell in report_ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Business-friendly formatting.
    header_map = {cell.value: cell.column for cell in report_ws[1]}
    price_col = header_map.get("Price")
    if price_col:
        for row_no in range(2, report_ws.max_row + 1):
            report_ws.cell(row=row_no, column=price_col).number_format = '£0.00'

    widths = {
        "Order": 6, "Address": 30, "Postcode": 13, "Price": 11,
        "Cleaning Plan": 16, "Phone": 16,
        "Payment Method": 18, "Payment Status": 17, "Payment Time": 14,
        "Message": 12, "Sent Time": 12, "Service Date": 14,
        "Notes": 32, "Next Cleaning Due": 18,
    }
    for heading, width in widths.items():
        col = header_map.get(heading)
        if col:
            report_ws.column_dimensions[report_ws.cell(row=1, column=col).column_letter].width = width

    completed_revenue = float(report_source["Price"].sum())
    cash_received = float(report_source.loc[report_source["Payment"].astype(str).eq("Cash"), "Price"].sum())
    bank_received = float(report_source.loc[report_source["Payment"].astype(str).eq("Bank Transfer Paid"), "Price"].sum())
    outstanding = max(completed_revenue - cash_received - bank_received, 0.0)

    summary_start = report_ws.max_row + 3
    summary_rows = [
        ["DanCleanUK Daily Summary", ""],
        ["Date", report_date(service_date_str)],
        ["Completed Jobs", len(report_df)],
        ["Revenue", completed_revenue],
        ["Cash Received", cash_received],
        ["Bank Transfer Received", bank_received],
        ["Outstanding", outstanding],
    ]
    for offset, values in enumerate(summary_rows):
        row_num = summary_start + offset
        report_ws.cell(row=row_num, column=1, value=values[0])
        report_ws.cell(row=row_num, column=2, value=values[1])
    report_ws.cell(row=summary_start, column=1).fill = header_fill
    report_ws.cell(row=summary_start, column=1).font = header_font
    for row_num in range(summary_start + 3, summary_start + 7):
        report_ws.cell(row=row_num, column=2).number_format = '£0.00'
    report_ws.column_dimensions["A"].width = max(report_ws.column_dimensions["A"].width or 0, 32)
    report_ws.column_dimensions["B"].width = max(report_ws.column_dimensions["B"].width or 0, 20)
    report_ws.freeze_panes = "A2"
    report_ws.auto_filter.ref = f"A1:{report_ws.cell(row=report_ws.max_row if summary_start <= 1 else summary_start-3, column=report_ws.max_column).coordinate}"

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
            st.sidebar.download_button(
                "⬇️ Download Report",
                base64.b64decode(saved["report_b64"]),
                saved.get("report_filename") or report_filename,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception:
            pass

st.sidebar.caption(f"DanCleanUK Route Optimizer v{APP_VERSION}")
