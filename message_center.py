import re
from datetime import datetime, timedelta, timezone
from urllib.parse import quote
from zoneinfo import ZoneInfo

import requests
import streamlit as st


st.set_page_config(
    page_title="DanCleanUK Reminders",
    page_icon="📱",
    layout="centered",
)

st.markdown(
    """
    <style>
    .block-container {max-width:720px;padding-top:1.5rem;padding-bottom:4rem}
    div[data-testid="stButton"]>button,
    div[data-testid="stLinkButton"]>a {min-height:3.2rem;font-weight:700}
    html,body {overscroll-behavior-y:none}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_val(value):
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"none", "nan", "nat"} else text


def normalise_postcode(value):
    text = re.sub(r"\s+", "", clean_val(value).upper())
    if 5 <= len(text) <= 7:
        return f"{text[:-3]} {text[-3:]}"
    return clean_val(value).upper()


def normalise_phone(value):
    text = re.sub(r"[\s\-\(\)]", "", clean_val(value))
    if text.startswith("07"):
        text = "44" + text[1:]
    elif text.startswith("+44"):
        text = text[1:]
    return text


def sms_url(phone, message):
    number = normalise_phone(phone)
    if number and not number.startswith("+"):
        number = "+" + number
    return "sms:" + quote(number, safe="+") + "?body=" + quote(message)


def london_today():
    return datetime.now(ZoneInfo("Europe/London")).date()


# ============================================================
# SUPABASE — SAME DATABASE AS app.py
# ============================================================

def supabase_config():
    try:
        url = str(st.secrets["SUPABASE_URL"]).strip().rstrip("/")
        key = str(st.secrets["SUPABASE_SECRET_KEY"]).strip()
    except Exception:
        return None, None
    if not url.startswith("https://") or not key:
        return None, None
    return url, key


def headers(prefer=None):
    url, key = supabase_config()
    if not url or not key:
        return None
    result = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if prefer:
        result["Prefer"] = prefer
    return result


def connected():
    url, key = supabase_config()
    return bool(url and key)


# ============================================================
# TOMORROW / ACCESS REMINDERS
# ============================================================

def cleaning_message():
    return (
        "Hi, this is DanCleanUK 👋 Just a reminder that we’re due to clean your "
        "windows tomorrow. Please make sure we have access to the property. "
        "If tomorrow is not suitable, please reply to this message. "
        "Thank you, DanCleanUK."
    )


def load_cleaning_queue(target_date):
    url, _ = supabase_config()
    h = headers()
    if not url or not h:
        return []
    try:
        r = requests.get(
            f"{url}/rest/v1/customer_records",
            headers=h,
            params={
                "active": "eq.true",
                "reminder_queued_for": f"eq.{target_date}",
                "select": (
                    "id,address,postcode,phone,next_cleaning_due,"
                    "planned_service_date,reminder_sent_for,reminder_sent_at,"
                    "reminder_queued_for,reminder_queued_at"
                ),
                "order": "postcode.asc,address.asc",
            },
            timeout=20,
        )
        return (r.json() or []) if r.status_code == 200 else []
    except (requests.RequestException, ValueError, TypeError):
        return []


def mark_cleaning_sent(customer_id, target_date):
    """Record SENT and remove this customer from the prepared queue."""
    url, _ = supabase_config()
    h = headers("return=representation")
    if not url or not h or customer_id in (None, ""):
        return False

    now = datetime.now(timezone.utc).isoformat()
    try:
        r = requests.patch(
            f"{url}/rest/v1/customer_records",
            headers=h,
            params={"id": f"eq.{customer_id}", "select": "id"},
            json={
                "reminder_sent_for": str(target_date),
                "reminder_sent_at": now,
                "reminder_queued_for": None,
                "reminder_queued_at": None,
                "updated_at": now,
            },
            timeout=20,
        )
        return r.status_code == 200 and bool(r.json())
    except (requests.RequestException, ValueError, TypeError):
        return False


# ============================================================
# PAYMENT REMINDERS
# ============================================================

def load_saved_routes():
    url, _ = supabase_config()
    h = headers()
    if not url or not h:
        return []
    try:
        r = requests.get(
            f"{url}/rest/v1/saved_routes",
            headers=h,
            params={
                "select": "route_date,route_data",
                "order": "route_date.desc",
                "limit": "1000",
            },
            timeout=30,
        )
        return (r.json() or []) if r.status_code == 200 else []
    except (requests.RequestException, ValueError, TypeError):
        return []


def load_payment_queue():
    queued = []
    for snapshot in load_saved_routes():
        route_date = clean_val(snapshot.get("route_date"))
        route_data = snapshot.get("route_data") or {}
        for original in route_data.get("jobs_data", []) or []:
            job = dict(original)
            try:
                stage = int(job.get("PaymentReminderQueuedStage", 0) or 0)
            except Exception:
                stage = 0
            if stage not in (3, 7):
                continue
            if clean_val(job.get("Status")).lower() != "completed":
                continue
            if clean_val(job.get("Payment")) not in {"Bank Transfer", "Not Paid"}:
                continue
            job["_route_date"] = route_date
            job["_reminder_stage"] = stage
            queued.append(job)

    return sorted(
        queued,
        key=lambda x: (
            clean_val(x.get("_route_date")),
            clean_val(x.get("address_text")).lower(),
        ),
    )


def service_date(job):
    raw = clean_val(job.get("CompletedTime")) or clean_val(job.get("completed_time"))
    if raw:
        try:
            dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
            if dt.tzinfo is not None:
                dt = dt.astimezone(ZoneInfo("Europe/London"))
            return dt.date()
        except Exception:
            pass
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
            try:
                return datetime.strptime(raw, fmt).date()
            except Exception:
                pass

    try:
        return datetime.strptime(clean_val(job.get("_route_date"))[:10], "%Y-%m-%d").date()
    except Exception:
        return None


def payment_message(job, stage):
    try:
        amount = float(job.get("Price", 0) or 0)
    except Exception:
        amount = 0.0

    day = service_date(job)
    when = day.strftime("%d/%m/%Y") if day else clean_val(job.get("_route_date"))

    if int(stage) == 7:
        return (
            f"Hi, this is DanCleanUK. Just a second reminder that the £{amount:.2f} payment "
            f"for your window clean on {when} is still showing as outstanding. "
            "If you've already paid, please ignore this message. Otherwise, we'd appreciate "
            "it if you could arrange payment when possible. Thank you, DanCleanUK."
        )

    return (
        f"Hi, this is DanCleanUK 👋 Just a friendly reminder that the £{amount:.2f} payment "
        f"for your window clean on {when} is still showing as outstanding. "
        "If you've already made the payment, please ignore this message. Thank you! DanCleanUK"
    )


def load_route(route_date):
    url, _ = supabase_config()
    h = headers()
    if not url or not h:
        return None
    try:
        r = requests.get(
            f"{url}/rest/v1/saved_routes",
            headers=h,
            params={
                "route_date": f"eq.{route_date}",
                "select": "route_date,route_data",
                "limit": "1",
            },
            timeout=20,
        )
        rows = (r.json() or []) if r.status_code == 200 else []
        return rows[0] if rows else None
    except (requests.RequestException, ValueError, TypeError):
        return None


def mark_payment_sent(job):
    route_date = clean_val(job.get("_route_date"))
    job_id = job.get("job_id")
    try:
        stage = int(job.get("_reminder_stage", 0) or 0)
    except Exception:
        stage = 0

    snapshot = load_route(route_date)
    if not snapshot:
        return False

    route_data = dict(snapshot.get("route_data") or {})
    jobs = [dict(x) for x in (route_data.get("jobs_data") or [])]
    found = False

    for row in jobs:
        if str(row.get("job_id")) == str(job_id):
            sent_field = (
                "PaymentReminderDay7SentAt"
                if stage == 7
                else "PaymentReminderDay3SentAt"
            )
            row[sent_field] = datetime.now(timezone.utc).isoformat()
            row["PaymentReminderQueuedStage"] = 0
            row["PaymentReminderQueuedAt"] = ""
            found = True
            break

    if not found:
        return False

    route_data["jobs_data"] = jobs
    url, _ = supabase_config()
    h = headers("return=representation")
    if not url or not h:
        return False

    try:
        r = requests.patch(
            f"{url}/rest/v1/saved_routes",
            headers=h,
            params={"route_date": f"eq.{route_date}", "select": "route_date"},
            json={
                "route_data": route_data,
                "updated_at": datetime.now(timezone.utc).isoformat(),
            },
            timeout=30,
        )
        return r.status_code == 200 and bool(r.json())
    except (requests.RequestException, ValueError, TypeError):
        return False


# ============================================================
# SESSION HELPERS
# ============================================================

def reset_cleaning():
    for key in ("clean_queue", "clean_date", "clean_index", "clean_sent"):
        st.session_state.pop(key, None)


def reset_payments():
    for key in ("pay_queue", "pay_index", "pay_sent"):
        st.session_state.pop(key, None)


def home():
    st.session_state["screen"] = "home"


if "screen" not in st.session_state:
    st.session_state["screen"] = "home"


# ============================================================
# CONNECTION CHECK
# ============================================================

if not connected():
    st.title("📱 DanCleanUK — Phone Reminders")
    st.error("This Message Centre is not connected to Supabase yet.")
    st.info(
        "After we deploy this second Streamlit app, copy the same SUPABASE_URL "
        "and SUPABASE_SECRET_KEY into this app's Streamlit Secrets."
    )
    st.stop()


# ============================================================
# HOME
# ============================================================

if st.session_state["screen"] == "home":
    st.title("📱 DanCleanUK — Phone Reminders")
    st.caption("Choose which prepared messages you want to send.")

    if st.button(
        "🪟 TOMORROW'S CLEANING / ACCESS REMINDERS",
        type="primary",
        use_container_width=True,
    ):
        st.session_state["screen"] = "cleaning"
        st.rerun()

    st.caption(
        "Tomorrow's message includes the reminder to make sure we have access to the property."
    )

    if st.button(
        "💷 OUTSTANDING PAYMENT REMINDERS",
        use_container_width=True,
    ):
        st.session_state["screen"] = "payments"
        st.rerun()

    st.caption("Prepared Day 3 / Day 7 outstanding-payment reminders.")
    st.markdown("---")
    st.caption("Separate Message Centre · Driver app and route optimiser stay untouched.")
    st.stop()


# ============================================================
# CLEANING / ACCESS SCREEN
# ============================================================

if st.session_state["screen"] == "cleaning":
    if st.button("← MESSAGE CENTRE", use_container_width=True):
        home()
        st.rerun()

    st.title("🪟 Tomorrow / Access Reminders")
    st.caption("Send prepared reminders through this phone's normal Messages app.")

    tomorrow = london_today() + timedelta(days=1)
    target = tomorrow.isoformat()

    if st.session_state.get("clean_date") not in (None, target):
        reset_cleaning()

    queue = st.session_state.get("clean_queue")

    if not queue:
        valid = []
        for row in load_cleaning_queue(target):
            scheduled = clean_val(row.get("planned_service_date")) or clean_val(
                row.get("next_cleaning_due")
            )
            phone = clean_val(row.get("phone"))
            if scheduled != target or not phone:
                continue
            valid.append(
                {
                    "id": row.get("id"),
                    "address": clean_val(row.get("address")) or "Customer",
                    "postcode": normalise_postcode(row.get("postcode")),
                    "phone": phone,
                    "sent_for": clean_val(row.get("reminder_sent_for")),
                }
            )

        st.subheader(f"Tomorrow · {tomorrow.strftime('%d/%m/%Y')}")

        if not valid:
            st.info(
                "No reminders are prepared for tomorrow. "
                "Prepare them in Admin → Upcoming Work / Planner first."
            )
            if st.button("🔄 Refresh", use_container_width=True):
                st.rerun()
            st.stop()

        st.success(f"{len(valid)} customer reminder(s) ready.")

        with st.expander("Preview message"):
            st.write(cleaning_message())

        for row in valid:
            with st.container(border=True):
                st.markdown(f"**{row['address']}**")
                st.caption(f"{row['postcode']} · 📞 {row['phone']}")
                if row["sent_for"] == target:
                    st.warning("A reminder is already recorded for tomorrow.")

        if st.button(
            f"▶️ START SENDING {len(valid)} REMINDER(S)",
            type="primary",
            use_container_width=True,
        ):
            st.session_state["clean_queue"] = valid
            st.session_state["clean_date"] = target
            st.session_state["clean_index"] = 0
            st.session_state["clean_sent"] = 0
            st.rerun()
        st.stop()

    index = int(st.session_state.get("clean_index", 0))
    sent = int(st.session_state.get("clean_sent", 0))

    if index >= len(queue):
        st.success(f"✅ Finished. {sent} reminder(s) marked as sent.")
        remaining = load_cleaning_queue(target)
        if remaining:
            st.info(f"{len(remaining)} prepared reminder(s) remain, including any you skipped.")
        else:
            st.info("No prepared cleaning reminders remain.")
        if st.button("🔄 RELOAD REMAINING QUEUE", type="primary", use_container_width=True):
            reset_cleaning()
            st.rerun()
        st.stop()

    current = queue[index]
    st.progress((index + 1) / len(queue))
    st.caption(f"Customer {index + 1} of {len(queue)} · {sent} marked sent")
    st.markdown(f"### {current['address']}")
    st.write(current["postcode"])
    st.write(f"📞 {current['phone']}")

    with st.expander("Message", expanded=True):
        st.write(cleaning_message())

    st.link_button(
        "💬 OPEN SMS — MESSAGE READY",
        sms_url(current["phone"], cleaning_message()),
        type="primary",
        use_container_width=True,
    )
    st.caption("Send the SMS, return here, then press SENT — NEXT.")

    left, right = st.columns(2)
    if left.button("✅ SENT — NEXT", type="primary", use_container_width=True):
        if mark_cleaning_sent(current["id"], target):
            st.session_state["clean_sent"] = sent + 1
            st.session_state["clean_index"] = index + 1
            st.rerun()
        else:
            st.error("Could not save the sent status. Please try again before moving on.")

    if right.button("⏭️ SKIP", use_container_width=True):
        st.session_state["clean_index"] = index + 1
        st.rerun()

    if st.button("↩️ End / reload queue", use_container_width=True):
        reset_cleaning()
        st.rerun()

    st.stop()


# ============================================================
# PAYMENT SCREEN
# ============================================================

if st.session_state["screen"] == "payments":
    if st.button("← MESSAGE CENTRE", use_container_width=True):
        home()
        st.rerun()

    st.title("💷 Outstanding Payment Reminders")
    st.caption("Send prepared reminders through this phone's normal Messages app.")

    queue = st.session_state.get("pay_queue")

    if not queue:
        ready = load_payment_queue()

        if not ready:
            st.success("No payment reminders are prepared for this phone.")
            st.caption("Prepare them in Admin → Outstanding Payments first.")
            if st.button("🔄 Refresh", use_container_width=True):
                st.rerun()
            st.stop()

        st.success(f"{len(ready)} payment reminder(s) ready.")

        for row in ready:
            address = clean_val(row.get("Address")) or clean_val(row.get("address_text")) or "Customer"
            stage = int(row.get("_reminder_stage", 3) or 3)
            price = float(row.get("Price", 0) or 0)
            with st.container(border=True):
                st.markdown(f"**{address}**")
                st.caption(
                    f"{normalise_postcode(row.get('Postcode'))} · "
                    f"£{price:.2f} · Day {stage}"
                )

        if st.button(
            f"▶️ START SENDING {len(ready)} PAYMENT REMINDER(S)",
            type="primary",
            use_container_width=True,
        ):
            st.session_state["pay_queue"] = ready
            st.session_state["pay_index"] = 0
            st.session_state["pay_sent"] = 0
            st.rerun()
        st.stop()

    index = int(st.session_state.get("pay_index", 0))
    sent = int(st.session_state.get("pay_sent", 0))

    if index >= len(queue):
        st.success(f"✅ Finished. {sent} payment reminder(s) marked as sent.")
        remaining = load_payment_queue()
        if remaining:
            st.info(f"{len(remaining)} prepared reminder(s) remain, including any you skipped.")
        else:
            st.info("No prepared payment reminders remain.")
        if st.button("🔄 RELOAD REMAINING QUEUE", type="primary", use_container_width=True):
            reset_payments()
            st.rerun()
        st.stop()

    current = queue[index]
    address = clean_val(current.get("Address")) or clean_val(current.get("address_text")) or "Customer"
    stage = int(current.get("_reminder_stage", 3) or 3)
    price = float(current.get("Price", 0) or 0)
    message = payment_message(current, stage)

    st.progress((index + 1) / len(queue))
    st.caption(f"Customer {index + 1} of {len(queue)} · {sent} marked sent")
    st.markdown(f"### {address}")
    st.write(normalise_postcode(current.get("Postcode")))
    st.write(f"📞 {clean_val(current.get('Phone'))}")
    st.write(f"**£{price:.2f} outstanding · Day {stage} reminder**")

    with st.expander("Message", expanded=True):
        st.write(message)

    st.link_button(
        "💬 OPEN SMS — MESSAGE READY",
        sms_url(current.get("Phone"), message),
        type="primary",
        use_container_width=True,
    )
    st.caption("Send the SMS, return here, then press SENT — NEXT.")

    left, right = st.columns(2)
    if left.button("✅ SENT — NEXT", type="primary", use_container_width=True):
        if mark_payment_sent(current):
            st.session_state["pay_sent"] = sent + 1
            st.session_state["pay_index"] = index + 1
            st.rerun()
        else:
            st.error("Could not save the sent status. Please try again before moving on.")

    if right.button("⏭️ SKIP", use_container_width=True):
        st.session_state["pay_index"] = index + 1
        st.rerun()

    if st.button("↩️ End / reload queue", use_container_width=True):
        reset_payments()
        st.rerun()

    st.stop()
