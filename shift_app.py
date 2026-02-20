import streamlit as st
import pandas as pd
import re
from datetime import datetime, date, timedelta
import smtplib
from email.message import EmailMessage

# ============================================================
# Embedded Shift Calculator Logic (from your provided app)
# ============================================================
ANCHOR_DATE = date(2025, 7, 1)
ANCHOR_DAY_NUM = 2  # July 1, 2025 is Day 2

def get_day_number(target_dt: date) -> int:
    """Calculates the day number (1-4) for a given date."""
    delta_days = (target_dt - ANCHOR_DATE).days
    return (delta_days + ANCHOR_DAY_NUM - 1) % 4 + 1

def gold_shift_type(num: int, day_num: int) -> str:
    """Gold rules per your code."""
    if num == 1:
        return "Early"
    if num >= 6:
        return "Middle"
    if num in (3, 5):
        return "Early" if day_num in (1, 3) else "Middle"
    if num in (2, 4):
        return "Middle" if day_num in (1, 3) else "Early"
    raise ValueError(f"Invalid/Unhandled Gold number {num}")

def silver_shift_type(num: int) -> str:
    """Silver rules per your code."""
    if num == 1:
        return "Early"
    if num >= 2:
        return "Middle"
    raise ValueError(f"Invalid/Unhandled Silver number {num}")

def calc_shift_gold_silver(line: str, num: int, target_dt: date) -> str:
    """Returns 'Early' or 'Middle' for Gold/Silver."""
    day_num = get_day_number(target_dt)
    if line.lower() == "gold":
        return gold_shift_type(num, day_num)
    if line.lower() == "silver":
        return silver_shift_type(num)
    raise ValueError("Only Gold/Silver supported in this workflow.")

# ============================================================
# Parsing & Utilities
# ============================================================
# Gold requires a number; Silver number optional ("Silver AM" -> Silver 1)
GOLD_PATTERN = re.compile(r"^Gold\s+AM(?:\s+(\d+))?$", re.IGNORECASE)
SILVER_PATTERN = re.compile(r"^Silver\s+AM(?:\s+(\d+))?$", re.IGNORECASE)

def parse_event_title_with_reason(raw_title):
    """
    Accepts:
      - Gold AM <num>
      - Silver AM
      - Silver AM <num>

    Returns:
      (("Gold"/"Silver", num, "Gold 5"/"Silver 1"), None) on success
      (None, "reason") on failure
    """
    if raw_title is None or (isinstance(raw_title, float) and pd.isna(raw_title)):
        return None, "Missing event title"
    if not isinstance(raw_title, str):
        return None, "Event title is not text"

    s = raw_title.strip()
    if not s:
        return None, "Empty event title"

    has_gold = re.search(r"\bgold\b", s, re.IGNORECASE) is not None
    has_silver = re.search(r"\bsilver\b", s, re.IGNORECASE) is not None

    if not (has_gold or has_silver):
        return None, "Not a Gold/Silver event"

    if re.search(r"\bam\b", s, re.IGNORECASE) is None:
        return None, "Missing 'AM' (expected 'Gold AM #' or 'Silver AM [#]')"

    # GOLD
    if has_gold and re.match(r"^gold\b", s, re.IGNORECASE):
        mg = GOLD_PATTERN.match(s)
        if not mg:
            return None, "Gold format invalid (expected 'Gold AM <number>')"
        num_str = mg.group(1)
        if not num_str:
            return None, "Gold missing number (expected 'Gold AM <number>')"
        try:
            num = int(num_str)
        except ValueError:
            return None, f"Gold number is not an integer: '{num_str}'"
        if num <= 0:
            return None, "Gold number must be positive"
        return ("Gold", num, f"Gold {num}"), None

    # SILVER
    if has_silver and re.match(r"^silver\b", s, re.IGNORECASE):
        ms = SILVER_PATTERN.match(s)
        if not ms:
            return None, "Silver format invalid (expected 'Silver AM' or 'Silver AM <number>')"
        num_str = ms.group(1)
        if not num_str:
            num = 1  # default rule
        else:
            try:
                num = int(num_str)
            except ValueError:
                return None, f"Silver number is not an integer: '{num_str}'"
            if num <= 0:
                return None, "Silver number must be positive"
        return ("Silver", num, f"Silver {num}"), None

    return None, "Unrecognized title format"

def parse_any_date(x):
    """Robust date parsing from common exports. Returns date or None."""
    if pd.isna(x):
        return None
    s = str(x).strip()
    if not s:
        return None
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    if pd.isna(dt):
        return None
    return dt.date()

def shift_to_lower(shift: str) -> str:
    return shift.strip().lower()

def shift_to_start_time(shift_lower: str) -> str:
    if shift_lower == "early":
        return "06:45"
    if shift_lower == "middle":
        return "08:00"
    return "UNKNOWN"

def build_ics(df: pd.DataFrame, duration_minutes: int = 60) -> str:
    """Builds a basic Outlook-importable ICS file."""
    ics = "BEGIN:VCALENDAR\nVERSION:2.0\nCALSCALE:GREGORIAN\nPRODID:-//ShiftApp//EN\n"
    for _, r in df.iterrows():
        if r.get("start_time") == "UNKNOWN" or pd.isna(r.get("event_date_parsed")):
            continue
        start_dt = datetime.combine(
            r["event_date_parsed"],
            datetime.strptime(r["start_time"], "%H:%M").time()
        )
        end_dt = start_dt + timedelta(minutes=duration_minutes)

        summary = str(r.get("clean_event", "Shift"))
        # Basic escaping
        summary = summary.replace("\\", "\\\\").replace("\n", "\\n").replace(",", "\\,").replace(";", "\\;")

        ics += (
            "BEGIN:VEVENT\n"
            f"SUMMARY:{summary}\n"
            f"DTSTART:{start_dt.strftime('%Y%m%dT%H%M%S')}\n"
            f"DTEND:{end_dt.strftime('%Y%m%dT%H%M%S')}\n"
            "END:VEVENT\n"
        )
    ics += "END:VCALENDAR\n"
    return ics

def send_email_smtp(
    smtp_server: str,
    smtp_port: int,
    from_email: str,
    password: str,
    to_email: str,
    subject: str,
    body: str,
    filename: str,
    attachment_text: str,
):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email
    msg.set_content(body)
    msg.add_attachment(
        attachment_text.encode("utf-8"),
        maintype="text",
        subtype="calendar",
        filename=filename
    )
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(from_email, password)
        server.send_message(msg)

# ============================================================
# Streamlit App
# ============================================================
st.set_page_config(page_title="Gold/Silver Shift Scheduler", layout="centered")
st.title("Gold/Silver Shift Scheduler (CSV → Outlook + Email)")

st.markdown(
    """
Upload a **CSV** export (from Numbers/Outlook/etc.). The app will:
- process only titles like **Gold AM 5** / **Silver AM 2** / **Silver AM**
- treat **Silver AM** as **Silver AM 1**
- compute **early/middle** using the embedded calculator logic (no website calls)
- map to start times (**06:45** for early, **08:00** for middle)
- let you download a processed CSV and an Outlook **.ics**
- optionally email the **.ics**
"""
)

uploaded = st.file_uploader("Upload CSV", type=["csv"])

with st.expander("Optional settings"):
    event_duration_min = st.number_input(
        "Event duration (minutes) for Outlook export",
        min_value=15, max_value=24*60, value=60, step=15
    )

if uploaded:
    df_raw = pd.read_csv(uploaded)

    st.write("Detected columns:", list(df_raw.columns))

    # Column mapping (auto-detect, user can override)
    name_candidates = ["event_name", "Subject", "Title", "Event", "Name", "Summary"]
    date_candidates = ["event_date", "Date", "Start Date", "Start", "Event Date", "Start_Date", "StartDate"]

    default_name = next((c for c in name_candidates if c in df_raw.columns), df_raw.columns[0])
    default_date = next((c for c in date_candidates if c in df_raw.columns), df_raw.columns[min(1, len(df_raw.columns)-1)])

    st.subheader("Column mapping")
    name_col = st.selectbox("Event title column", options=list(df_raw.columns), index=list(df_raw.columns).index(default_name))
    date_col = st.selectbox("Event date column", options=list(df_raw.columns), index=list(df_raw.columns).index(default_date))

    df = df_raw.rename(columns={name_col: "event_name", date_col: "event_date"}).copy()

    # Parse titles with reasons
    parsed = df["event_name"].apply(parse_event_title_with_reason)
    parsed_tuple_series = parsed.apply(lambda x: x[0])
    reason_series = parsed.apply(lambda x: x[1])

    rejected_df = df.loc[parsed_tuple_series.isna(), ["event_name", "event_date"]].copy()
    rejected_df["reason"] = reason_series.loc[parsed_tuple_series.isna()].values

    df_ok = df.loc[parsed_tuple_series.notna()].copy()

    if df_ok.empty:
        st.warning("No rows matched the required Gold/Silver AM patterns.")
        if not rejected_df.empty:
            st.subheader("Rejected Events")
            st.caption("Fix titles and re-upload if needed.")
            st.dataframe(rejected_df, use_container_width=True)
        st.stop()

    # Unpack parsed tuples into columns
    parsed_ok = parsed_tuple_series.loc[parsed_tuple_series.notna()]
    df_ok[["line", "num", "clean_event"]] = pd.DataFrame(parsed_ok.tolist(), index=df_ok.index)

    # Parse dates (and show specific bad date rows)
    df_ok["event_date_parsed"] = df_ok["event_date"].apply(parse_any_date)
    bad_dates = df_ok["event_date_parsed"].isna()
    if bad_dates.any():
        bad_df = df_ok.loc[bad_dates, ["event_name", "event_date"]].copy()
        bad_df["reason"] = "Date could not be parsed"
        # Add to rejected list and remove from processing
        rejected_df = pd.concat([rejected_df, bad_df], ignore_index=True)
        df_ok = df_ok.loc[~bad_dates].copy()

    if df_ok.empty:
        st.warning("All candidate rows were rejected due to unparseable dates.")
        st.subheader("Rejected Events")
        st.dataframe(rejected_df, use_container_width=True)
        st.stop()

    # Calculate shift + start time
    shift_results = []
    for _, r in df_ok.iterrows():
        try:
            shift = calc_shift_gold_silver(str(r["line"]), int(r["num"]), r["event_date_parsed"])
            shift_results.append(shift_to_lower(shift))
        except Exception as e:
            shift_results.append(f"error: {e}")

    df_ok["shift_result"] = shift_results
    df_ok["start_time"] = df_ok["shift_result"].apply(shift_to_start_time)

    st.subheader("Processed Results")
    st.dataframe(
        df_ok[["event_name", "clean_event", "event_date", "event_date_parsed", "shift_result", "start_time"]],
        use_container_width=True
    )

    # Downloads
    st.subheader("Downloads")
    csv_bytes = df_ok.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download processed CSV",
        data=csv_bytes,
        file_name="events_with_shift_times.csv",
        mime="text/csv"
    )

    ics_text = build_ics(df_ok, duration_minutes=int(event_duration_min))
    st.download_button(
        "Download Outlook Calendar (.ics)",
        data=ics_text,
        file_name="shift_schedule.ics",
        mime="text/calendar"
    )

    # Rejected Events
    if not rejected_df.empty:
        st.subheader("Rejected Events")
        st.caption("These rows were ignored; see the reason column.")
        st.dataframe(rejected_df, use_container_width=True)

        rejected_csv = rejected_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download rejected rows (CSV)",
            data=rejected_csv,
            file_name="rejected_events.csv",
            mime="text/csv"
        )

    # Email
    st.subheader("Email the Outlook calendar (.ics)")

    col1, col2 = st.columns(2)
    with col1:
        to_email = st.text_input("Send to")
        smtp_server = st.text_input("SMTP server (e.g. smtp.gmail.com)")
    with col2:
        smtp_port = st.number_input("SMTP port", value=587, step=1)
        from_email = st.text_input("From email")

    password = st.text_input("Password / App Password", type="password")

    if st.button("Send email"):
        if not (to_email and smtp_server and from_email and password):
            st.error("Please fill out To, SMTP server, From email, and Password/App Password.")
        else:
            try:
                send_email_smtp(
                    smtp_server=smtp_server,
                    smtp_port=int(smtp_port),
                    from_email=from_email,
                    password=password,
                    to_email=to_email,
                    subject="Shift Schedule (Outlook .ics)",
                    body="Attached is your shift schedule calendar file.",
                    filename="shift_schedule.ics",
                    attachment_text=ics_text
                )
                st.success("Email sent!")
            except Exception as e:
                st.error(f"Email failed: {e}")
