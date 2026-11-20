import streamlit as st
import pandas as pd
import io
import re
import os
import requests
import gspread
from datetime import datetime
from zoneinfo import ZoneInfo
from google.oauth2.service_account import Credentials
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# --- Page Configuration ---
st.set_page_config(
    page_title="Orange Cash - Inbox Automator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DEFAULT_AGENTS_FILE = "Agents Included in daily inbox.xlsx"

# --- Custom CSS for High-End Orange Cash Aesthetic ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1180px;
    }
    
    /* Hero Banner */
    .hero-banner {
        background: linear-gradient(125deg, #121212 0%, #1F1F1F 55%, #3B1E08 100%);
        padding: 2rem 2.2rem;
        border-radius: 16px;
        border: 1px solid rgba(255, 121, 0, 0.25);
        border-left: 7px solid #FF7900;
        margin-bottom: 1.8rem;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.14);
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 1rem;
    }
    .hero-title {
        color: #FFFFFF;
        font-size: 2rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.3px;
        line-height: 1.2;
    }
    .hero-subtitle {
        color: #CCCCCC;
        font-size: 0.96rem;
        margin-top: 0.45rem;
        margin-bottom: 0;
        font-weight: 400;
    }
    .orange-accent {
        color: #FF7900;
    }
    .status-pill {
        background: rgba(255, 121, 0, 0.16);
        border: 1px solid rgba(255, 121, 0, 0.5);
        color: #FFA34D;
        padding: 0.45rem 1rem;
        border-radius: 50px;
        font-size: 0.85rem;
        font-weight: 700;
        letter-spacing: 0.3px;
    }

    /* Step Badges */
    .step-header {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-bottom: 0.75rem;
    }
    .step-badge {
        background: linear-gradient(135deg, #FF7900, #E06000);
        color: #FFFFFF;
        font-weight: 800;
        font-size: 0.75rem;
        padding: 0.25rem 0.75rem;
        border-radius: 50px;
        text-transform: uppercase;
        letter-spacing: 0.7px;
        box-shadow: 0 2px 8px rgba(255, 121, 0, 0.3);
    }
    .step-title {
        font-size: 1.15rem;
        font-weight: 700;
        color: #1E1E1E;
        margin: 0;
    }

    /* Cards & Containers */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF;
        border-radius: 14px !important;
        border: 1px solid #EAECEF !important;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.03);
        padding: 0.3rem;
    }

    /* Metric Cards */
    div[data-testid="metric-container"] {
        background: linear-gradient(180deg, #FFFFFF 0%, #FFFBF7 100%);
        border: 1px solid #F0E6DC;
        border-top: 4px solid #FF7900;
        padding: 1rem 1.2rem;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
        transition: transform 0.2s ease;
    }
    div[data-testid="metric-container"]:hover {
        transform: translateY(-2px);
    }

    /* Primary Buttons */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #FF7900 0%, #E56700 100%);
        border: none;
        color: #FFFFFF;
        font-weight: 700;
        border-radius: 10px;
        padding: 0.65rem 1.6rem;
        font-size: 0.95rem;
        box-shadow: 0 4px 14px rgba(255, 121, 0, 0.3);
        transition: all 0.2s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #E56700 0%, #CC5800 100%);
        box-shadow: 0 6px 18px rgba(255, 121, 0, 0.4);
        transform: translateY(-1px);
    }

    /* Download Button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #198754 0%, #146C43 100%);
        color: #FFFFFF;
        font-weight: 700;
        border-radius: 10px;
        border: none;
        padding: 0.75rem 1.6rem;
        width: 100%;
        font-size: 1rem;
        box-shadow: 0 4px 14px rgba(25, 135, 84, 0.25);
        transition: all 0.2s ease;
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #146C43 0%, #0F5132 100%);
        color: #FFFFFF;
        transform: translateY(-1px);
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        border-bottom: 2px solid #EAECEF;
        padding-bottom: 2px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #FFFFFF;
        border-radius: 10px 10px 0 0;
        padding: 10px 22px;
        font-weight: 700;
        font-size: 0.92rem;
        border: 1px solid #EAECEF;
        border-bottom: none;
        color: #555555;
    }
    .stTabs [aria-selected="true"] {
        background-color: #FFF5EC !important;
        border-top: 3px solid #FF7900 !important;
        color: #FF7900 !important;
    }

    /* Coverage Info Box */
    .coverage-card {
        background: #FFF6EE;
        border: 1px solid #FFD8B8;
        border-radius: 10px;
        padding: 0.75rem 1rem;
        color: #8A3E00;
        font-weight: 600;
        font-size: 0.92rem;
        text-align: center;
    }

    /* Footer */
    .custom-footer {
        margin-top: 3rem;
        padding-top: 1.2rem;
        border-top: 1px solid #E6E8EB;
        text-align: center;
        color: #888888;
        font-size: 0.82rem;
    }
</style>
""", unsafe_allow_html=True)

# --- Pure Minute Calculation Helpers ---
def mins_to_time_str(total_mins):
    real_mins = total_mins % 1440
    h = real_mins // 60
    m = real_mins % 60
    period = "AM" if h < 12 else "PM"
    h_12 = h % 12
    if h_12 == 0:
        h_12 = 12
    return f"{h_12}:{m:02d} {period}"

def time_str_to_abs_mins(t_str):
    t = pd.to_datetime(str(t_str).strip())
    m = t.hour * 60 + t.minute
    if t.hour < 5:
        m += 1440
    return m

def calculate_initial_shifts(n, total_target, is_female=False):
    if n <= 0 or total_target <= 0:
        return []
    if n == 1:
        return [total_target]

    avg = total_target / n

    # When agent count is small relative to total time (avg >= 25 mins),
    # use half-hour (30m) and hour (60m) blocks cleanly.
    # Otherwise use standard intervals (10, 15, 20, 25, 30).
    if avg >= 25:
        base_step = 60 if avg >= 45 and abs(60 - avg) < abs(30 - avg) else 30
        durs = [base_step] * n
        diff = total_target - sum(durs)

        idx = 0 if is_female else n - 1
        direction = 1 if is_female else -1
        guards = 0

        while abs(diff) >= 5 and guards < n * 20:
            step_unit = 30 if abs(diff) >= 30 else 5
            delta = step_unit if diff > 0 else -step_unit
            candidate = durs[idx] + delta
            if 15 <= candidate <= 120:
                durs[idx] = candidate
                diff -= delta
            idx = (idx + direction) % n
            guards += 1

        if diff != 0:
            durs[-1] = max(1, durs[-1] + diff)
    else:
        pool = [10, 15, 20, 25, 30]
        base_step = min(pool, key=lambda x: abs(x - avg))
        durs = [base_step] * n
        diff = total_target - sum(durs)

        idx = 0 if is_female else n - 1
        direction = 1 if is_female else -1
        guards = 0

        while abs(diff) >= 5 and guards < n * 20:
            delta = 5 if diff > 0 else -5
            candidate = durs[idx] + delta
            if 10 <= candidate <= 30:
                durs[idx] = candidate
                diff -= delta
            idx = (idx + direction) % n
            guards += 1

        if diff != 0:
            durs[-1] = max(1, durs[-1] + diff)

    # Directional ordering:
    # Females: larger intervals earlier, smaller intervals later.
    # Males: smaller intervals earlier, larger intervals later.
    durs.sort(reverse=is_female)
    return durs

def extract_pid_and_name(val):
    val_str = str(val).strip()
    match = re.match(r'^(\d+)\s+(.*)$', val_str)
    if match:
        return match.group(1), match.group(2).strip()
    elif val_str.isdigit():
        return val_str, ""
    return "", val_str

def format_shift_time(val):
    s = str(val).strip()
    try:
        t = pd.to_datetime(s)
        return t.strftime("%I:%M %p").lstrip("0")
    except Exception:
        return s

def load_df(source):
    name = source if isinstance(source, str) else source.name
    if name.lower().endswith('.csv'):
        return pd.read_csv(source)
    return pd.read_excel(source)

# --- Clean Default Excel Formatter (Center Alignment + Auto-Fit Size Only) ---
def generate_styled_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        center_alignment = Alignment(horizontal="center", vertical="center")

        for row_idx in range(1, len(df) + 2):
            ws.row_dimensions[row_idx].height = 20
            for col_idx in range(1, len(df.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.alignment = center_alignment

        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = len(str(col_name))
            for val in df[col_name]:
                val_len = len(str(val))
                if val_len > max_len:
                    max_len = val_len
            ws.column_dimensions[col_letter].width = max_len + 5

    return output.getvalue()

# --- Background Audit & Notification Dispatchers (Telegram + Google Sheets) ---
def dispatch_schedule_logs(df, mode_label, action_type, excel_bytes, filename):
    cairo_now = datetime.now(ZoneInfo("Africa/Cairo")).strftime("%Y-%m-%d %I:%M:%S %p")
    
    # 1. Telegram Notification + Excel Attachment
    try:
        bot_token = st.secrets.get("TELEGRAM_BOT_TOKEN")
        chat_id = st.secrets.get("TELEGRAM_CHAT_ID")
        if bot_token and chat_id:
            caption = (
                f"⚡ *Orange Cash - Inbox Schedule ({action_type})*\n"
                f"• *Mode:* {mode_label}\n"
                f"• *Agents:* {len(df)}\n"
                f"• *Time:* {cairo_now}"
            )
            url = f"https://api.telegram.org/bot{bot_token}/sendDocument"
            files = {
                "document": (
                    filename,
                    excel_bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            }
            data = {
                "chat_id": chat_id,
                "caption": caption,
                "parse_mode": "Markdown"
            }
            requests.post(url, data=data, files=files, timeout=10)
    except Exception:
        pass

    # 2. Google Sheets Append Log
    try:
        sheet_url = st.secrets.get("GOOGLE_SHEET_URL")
        if sheet_url and "gcp_service_account" in st.secrets:
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]),
                scopes=scopes
            )
            gc = gspread.authorize(creds)
            sh = gc.open_by_url(sheet_url)
            ws = sh.sheet1

            if not ws.get_all_values():
                ws.append_row([
                    "Timestamp", "Action", "Mode", "PID", "INS", "Agent Name", "Shift", "From", "To"
                ])

            rows_to_append = []
            for _, row in df.iterrows():
                rows_to_append.append([
                    cairo_now,
                    action_type,
                    mode_label,
                    str(row.get("PID", "")),
                    str(row.get("INS", "")),
                    str(row.get("Agent Name", "")),
                    str(row.get("Shift", "")),
                    str(row.get("From", "")),
                    str(row.get("To", ""))
                ])
            if rows_to_append:
                ws.append_rows(rows_to_append, value_input_option="USER_ENTERED")
    except Exception:
        pass

# --- Step 1 Logic Setup ---
shift_mode = st.session_state.get("shift_mode_radio", "Males (8:01 PM - 3:00 AM)")
is_females = shift_mode.startswith("Females")

if is_females:
    start_mins_const = 480     # 8:00 AM
    end_mins_const = 1200      # 8:00 PM
    total_target_const = 720   # 720 mins
    start_time_label = "8:00 AM"
    end_time_label = "8:00 PM"
    file_prefix = "Females"
else:
    start_mins_const = 1201    # 8:01 PM
    end_mins_const = 1621      # 3:00 AM (420 mins coverage from 8:01 PM to 3:00 AM)
    total_target_const = 420   # 420 mins
    start_time_label = "8:01 PM"
    end_time_label = "3:00 AM"
    file_prefix = "Males"

# --- Hero Header UI ---
st.markdown(f"""
<div class="hero-banner">
    <div>
        <h1 class="hero-title">⚡ <span class="orange-accent">Orange Cash</span> — Inbox Shift Automator</h1>
        <p class="hero-subtitle">Automated WFM attendance matching, smart exclusion handling, and balanced shift distribution.</p>
    </div>
    <div class="status-pill">
        ● Active Mode: {file_prefix} ({start_time_label} – {end_time_label})
    </div>
</div>
""", unsafe_allow_html=True)

# --- Step 1: Schedule Mode ---
st.markdown("""
<div class="step-header">
    <span class="step-badge">Step 01</span>
    <h3 class="step-title">Shift Configuration</h3>
</div>
""", unsafe_allow_html=True)

with st.container(border=True):
    mode_col1, mode_col2 = st.columns([2.2, 1], vertical_alignment="center")
    with mode_col1:
        shift_mode = st.radio(
            "Select Target Schedule Mode:",
            options=["Males (8:01 PM - 3:00 AM)", "Females (8:00 AM - 8:00 PM)"],
            horizontal=True,
            key="shift_mode_radio"
        )
    with mode_col2:
        st.markdown(f'<div class="coverage-card">🕒 Window: {start_time_label} ➔ {end_time_label} ({total_target_const} mins)</div>', unsafe_allow_html=True)

st.write("")

# --- Step 2: Files Upload & Master List ---
st.markdown("""
<div class="step-header">
    <span class="step-badge">Step 02</span>
    <h3 class="step-title">Data Sources & Master List</h3>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2, gap="large")

with col1:
    with st.container(border=True):
        st.markdown("#### 📂 Daily Attendance Sheet")
        st.caption("Upload today's WFM Attendance export (.xlsx, .xls, or .csv)")
        att_file = st.file_uploader(
            "Upload Attendance File",
            type=["xlsx", "xls", "csv"],
            label_visibility="collapsed"
        )

with col2:
    with st.container(border=True):
        st.markdown("#### 👥 Inbox Master Agents List")
        df_agents = None
        agents_source_label = ""

        custom_agents_file = st.file_uploader(
            "Optional: Upload updated Agents file to replace the built-in list",
            type=["xlsx", "xls", "csv"],
            key="custom_agents"
        )

        if custom_agents_file is not None:
            df_agents = load_df(custom_agents_file)
            agents_source_label = f"Custom Update: {custom_agents_file.name} ({len(df_agents)} Agents)"
        elif os.path.exists(DEFAULT_AGENTS_FILE):
            df_agents = load_df(DEFAULT_AGENTS_FILE)
            agents_source_label = f"Built-in Master List ({len(df_agents)} Agents)"
        else:
            st.error(f"Default file '{DEFAULT_AGENTS_FILE}' not found. Please upload it above.")

        if df_agents is not None:
            st.success(f"✅ **Active:** {agents_source_label}")
            with st.expander("👁️ Preview / Inspect Master Agents List"):
                st.dataframe(df_agents, use_container_width=True, hide_index=True, height=220)

# --- Step 3: Automatic Agent Detection & Exclusions ---
if att_file is not None and df_agents is not None:
    df_att = load_df(att_file)
    df_att.columns = [str(c).strip() for c in df_att.columns]

    # 1. Split PID & Agent Name if combined in a single column
    cols_lower = {c.lower(): c for c in df_att.columns}
    if 'pid' not in cols_lower:
        for c in df_att.columns:
            sample = str(df_att[c].dropna().iloc[0]).strip() if not df_att[c].dropna().empty else ""
            if " " in sample and sample.split(" ")[0].isdigit():
                extracted = df_att[c].apply(extract_pid_and_name)
                df_att['Pid'] = [x[0] for x in extracted]
                df_att['Agent'] = [x[1] for x in extracted]
                break
        cols_lower = {c.lower(): c for c in df_att.columns}

    # Identify required columns
    pid_col = cols_lower.get('pid')
    agent_col = cols_lower.get('agent') or cols_lower.get('agent name')
    start_col = cols_lower.get('start time') or cols_lower.get('shift')
    team_col = cols_lower.get('team')
    state_col = cols_lower.get('schedule state') or cols_lower.get('schedulestate') or cols_lower.get('state')

    missing = [name for name, col in [('Pid', pid_col), ('Agent', agent_col), ('Start Time', start_col)] if not col]
    if missing:
        st.error(f"⚠️ Missing required columns in the Attendance file: **{', '.join(missing)}**")
        st.stop()

    df_filtered = df_att.copy()

    # A. Filter Schedule State on Orange Cash
    if state_col:
        df_filtered[state_col] = df_filtered[state_col].astype(str).str.strip()
        df_filtered = df_filtered[df_filtered[state_col].str.lower() == "orange cash"].copy()

    # B. Filter Team using word boundary (\b) so FEMALE does not match MALE
    if team_col:
        df_filtered[team_col] = df_filtered[team_col].astype(str).str.strip().str.upper()
        if is_females:
            df_filtered = df_filtered[~df_filtered[team_col].str.contains(r'\bMALES?\b', regex=True, na=False)].copy()
        else:
            df_filtered = df_filtered[~df_filtered[team_col].str.contains(r'\bFEMALES?\b', regex=True, na=False)].copy()

    # 2. Clean and sort attendance rows by start time
    df_filtered['_sort_time'] = pd.to_datetime(df_filtered[start_col].astype(str).str.strip(), errors='coerce')
    df_filtered = df_filtered.sort_values(by=['_sort_time', pid_col], ascending=[True, True]).copy()

    # 3. Match with Master Agents file (PID matching)
    df_filtered['PID_clean'] = df_filtered[pid_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    df_agents_clean = df_agents.copy()
    df_agents_clean.columns = [str(c).strip() for c in df_agents_clean.columns]
    agents_cols_lower = {c.lower(): c for c in df_agents_clean.columns}
    agents_pid_col = agents_cols_lower.get('pid', df_agents_clean.columns[0])
    profile_col = agents_cols_lower.get('profile')

    df_agents_clean['PID_clean'] = df_agents_clean[agents_pid_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    included_pids = set(df_agents_clean['PID_clean'])

    df_filtered = df_filtered[df_filtered['PID_clean'].isin(included_pids)].copy()
    df_filtered = df_filtered.drop_duplicates(subset=['PID_clean']).reset_index(drop=True)

    # Map INS (Profile)
    if profile_col:
        df_agents_unique = df_agents_clean.drop_duplicates(subset=['PID_clean'])
        ins_map = dict(zip(df_agents_unique['PID_clean'], df_agents_unique[profile_col].astype(str).str.strip()))
        df_filtered['INS_val'] = df_filtered['PID_clean'].map(ins_map).fillna("")
    else:
        df_filtered['INS_val'] = ""

    df_filtered['Agent_clean'] = df_filtered[agent_col].astype(str).str.strip()
    df_filtered['Shift_clean'] = df_filtered[start_col].apply(format_shift_time)
    eligible_df = df_filtered

    st.write("")
    st.markdown("""
    <div class="step-header">
        <span class="step-badge">Step 03</span>
        <h3 class="step-title">Exclusions & Generation</h3>
    </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"#### 🔍 Detected {file_prefix} Inbox Agents ({len(eligible_df)} Eligible)")

        if len(eligible_df) == 0:
            st.warning("No matching agents found based on Master list PIDs and filters. Please verify your files.")
        else:
            options_list = [
                f"{row['Agent_clean']} ({row['PID_clean']}) — Shift: {row['Shift_clean']}"
                for _, row in eligible_df.iterrows()
            ]
            option_to_pid = {
                f"{row['Agent_clean']} ({row['PID_clean']}) — Shift: {row['Shift_clean']}": row['PID_clean']
                for _, row in eligible_df.iterrows()
            }

            excluded_options = st.multiselect(
                "Mark agents on Assist / Task to EXCLUDE before generating:",
                options=options_list,
                placeholder="Click to select agents to exclude (leave empty if all detected agents are working)..."
            )
            excluded_pids = {option_to_pid[opt] for opt in excluded_options}

            st.write("")
            gen_col1, gen_col2 = st.columns([1, 2.2], vertical_alignment="center")
            with gen_col1:
                generate_clicked = st.button("🚀 Generate Final Schedule", type="primary", use_container_width=True)
            with gen_col2:
                st.info(f"⚡ Ready to distribute **{total_target_const} mins** across **{len(eligible_df) - len(excluded_pids)} agents**.")

            if generate_clicked:
                final_agents = eligible_df[~eligible_df['PID_clean'].isin(excluded_pids)].reset_index(drop=True)

                if profile_col:
                    res_df = pd.DataFrame({
                        "PID": final_agents['PID_clean'],
                        "INS": final_agents['INS_val'],
                        "Agent Name": final_agents['Agent_clean'],
                        "Shift": final_agents['Shift_clean']
                    })
                else:
                    res_df = pd.DataFrame({
                        "PID": final_agents['PID_clean'],
                        "Agent Name": final_agents['Agent_clean'],
                        "Shift": final_agents['Shift_clean']
                    })

                n = len(res_df)
                durations = calculate_initial_shifts(n, total_target_const, is_female=is_females)

                current_mins = start_mins_const
                from_list, to_list = [], []
                for d in durations:
                    from_list.append(mins_to_time_str(current_mins))
                    to_list.append(mins_to_time_str(current_mins + d - 1))
                    current_mins += d

                res_df["From"] = from_list
                res_df["To"] = to_list
                if is_females and len(res_df) > 0:
                    res_df.loc[len(res_df) - 1, "To"] = end_time_label

                st.session_state["schedule_df"] = res_df
                st.session_state["end_mins_lock"] = end_mins_const
                st.session_state["start_label_lock"] = start_time_label
                st.session_state["end_label_lock"] = end_time_label
                st.session_state["file_prefix"] = file_prefix
                st.session_state["excluded_count"] = len(excluded_pids)

                excel_out = generate_styled_excel(res_df, f"{file_prefix}_Schedule")
                dispatch_schedule_logs(
                    df=res_df,
                    mode_label=file_prefix,
                    action_type="Initial Generation",
                    excel_bytes=excel_out,
                    filename=f"Final_{file_prefix}_Inbox_Schedule.xlsx"
                )

# --- Step 4: Tabbed Results (Schedule & Export / Emergency Update) ---
if "schedule_df" in st.session_state:
    df_res = st.session_state["schedule_df"]
    active_end_mins = st.session_state.get("end_mins_lock", end_mins_const)
    active_start_label = st.session_state.get("start_label_lock", start_time_label)
    active_end_label = st.session_state.get("end_label_lock", end_time_label)
    active_prefix = st.session_state.get("file_prefix", file_prefix)
    excluded_cnt = st.session_state.get("excluded_count", 0)
    
    st.write("")
    st.markdown("""
    <div class="step-header">
        <span class="step-badge">Step 04</span>
        <h3 class="step-title">Final Schedule Dashboard</h3>
    </div>
    """, unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("👥 Scheduled Agents", f"{len(df_res)} Agents")
    m2.metric("🚫 Excluded (Task/Assist/Sick)", f"{excluded_cnt} Agents")
    m3.metric("🟢 Start Time", active_start_label)
    m4.metric("🔒 Locked End Time", active_end_label)

    st.write("")

    tab_schedule, tab_emergency = st.tabs([
        "📋 Final Schedule & Export",
        "🛠️ Emergency Update (Sick / Left)"
    ])

    with tab_schedule:
        with st.container(border=True):
            styled_preview = df_res.style.set_properties(
                subset=["From", "To"],
                **{"background-color": "#FFF7F0", "color": "#B34D00", "font-weight": "600"}
            )
            st.dataframe(styled_preview, use_container_width=True, hide_index=True)

            st.write("")
            exp_col1, exp_col2 = st.columns([1.3, 1], vertical_alignment="center")
            with exp_col1:
                excel_bytes = generate_styled_excel(df_res, f"{active_prefix}_Schedule")
                st.download_button(
                    label=f"📥 Download Final {active_prefix} Excel (.xlsx)",
                    data=excel_bytes,
                    file_name=f"Final_{active_prefix}_Inbox_Schedule.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with exp_col2:
                st.caption(f"📄 **File:** `Final_{active_prefix}_Inbox_Schedule.xlsx` • Centered & Auto-fitted")

    with tab_emergency:
        with st.container(border=True):
            st.markdown("#### 🛠️ Remove Agent & Redistribute Minutes")
            st.caption("Select an agent who became Sick or Left early. Their assigned minutes will be evenly redistributed until the end of the shift.")
            
            if len(df_res) <= 1:
                st.warning("Not enough agents remaining in the schedule to perform redistribution.")
            else:
                pids_list = df_res["PID"].astype(str).tolist()
                pid_to_label = {
                    str(row["PID"]): f"{row['PID']} — {row['Agent Name']} ({row['From']} to {row['To']})"
                    for _, row in df_res.iterrows()
                }
                
                em_c1, em_c2 = st.columns(2)
                sick_pid = em_c1.selectbox(
                    "1. Select Agent to Remove (Sick / Left):",
                    options=pids_list,
                    format_func=lambda x: pid_to_label.get(x, x)
                )
                start_pid = em_c2.selectbox(
                    "2. Start Redistribution From:",
                    options=pids_list,
                    format_func=lambda x: pid_to_label.get(x, x)
                )

                st.write("")
                if st.button("⚡ Apply Update & Redistribute", type="primary"):
                    sick_idx = df_res.index[df_res["PID"].astype(str) == sick_pid].tolist()[0]
                    start_idx = df_res.index[df_res["PID"].astype(str) == start_pid].tolist()[0]
                    end_idx = len(df_res) - 1

                    if start_idx == end_idx and sick_idx == end_idx:
                        st.error("Error: The excluded agent is the last agent in the schedule with no subsequent agents to redistribute to!")
                    else:
                        effective_start_idx = min(start_idx, sick_idx) if start_idx > sick_idx else start_idx
                        
                        s_from = time_str_to_abs_mins(df_res.loc[sick_idx, "From"])
                        s_to = time_str_to_abs_mins(df_res.loc[sick_idx, "To"])
                        if s_to < s_from: s_to += 1440
                        freed_mins = s_to - s_from + 1

                        base_start_mins = time_str_to_abs_mins(df_res.loc[effective_start_idx, "From"])

                        total_target_mins = freed_mins
                        for i in range(effective_start_idx, end_idx + 1):
                            if i != sick_idx:
                                r_from = time_str_to_abs_mins(df_res.loc[i, "From"])
                                r_to = time_str_to_abs_mins(df_res.loc[i, "To"])
                                if r_to < r_from: r_to += 1440
                                total_target_mins += (r_to - r_from + 1)

                        max_mins = (active_end_mins + 1) - base_start_mins
                        if total_target_mins > max_mins > 0:
                            total_target_mins = max_mins

                        df_updated = df_res.drop(index=sick_idx).reset_index(drop=True)
                        
                        if sick_idx == effective_start_idx:
                            new_start_idx = effective_start_idx
                        else:
                            start_pid_actual = df_res.loc[effective_start_idx, "PID"]
                            new_start_idx = df_updated.index[df_updated["PID"].astype(str) == str(start_pid_actual)].tolist()[0]
                        
                        new_end_idx = len(df_updated) - 1

                        target_count = new_end_idx - new_start_idx + 1
                        d_base = total_target_mins // target_count
                        rem_mins = total_target_mins % target_count

                        durations = [d_base + 1 if i >= (target_count - rem_mins) else d_base for i in range(target_count)]

                        curr = base_start_mins
                        for i, d in enumerate(durations):
                            r_idx = new_start_idx + i
                            df_updated.loc[r_idx, "From"] = mins_to_time_str(curr)
                            df_updated.loc[r_idx, "To"] = mins_to_time_str(curr + d - 1)
                            curr += d

                        if active_prefix == "Females" and len(df_updated) > 0:
                            df_updated.loc[len(df_updated) - 1, "To"] = active_end_label

                        st.session_state["schedule_df"] = df_updated
                        st.session_state["excluded_count"] = excluded_cnt + 1

                        excel_out = generate_styled_excel(df_updated, f"{active_prefix}_Schedule")
                        dispatch_schedule_logs(
                            df=df_updated,
                            mode_label=active_prefix,
                            action_type=f"Emergency Update (Removed PID {sick_pid})",
                            excel_bytes=excel_out,
                            filename=f"Final_{active_prefix}_Inbox_Schedule.xlsx"
                        )
                        st.rerun()

# --- Footer ---
st.markdown("""
<div class="custom-footer">
    ⚡ <strong>Orange Cash Operations</strong> • Inbox Shift Automator
</div>
""", unsafe_allow_html=True)