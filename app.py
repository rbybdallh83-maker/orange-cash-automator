import streamlit as st
import pandas as pd
import io
import re
import os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- Page Configuration ---
st.set_page_config(
    page_title="Orange Cash - Inbox Automator",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

DEFAULT_AGENTS_FILE = "Agents Included in daily inbox.xlsx"

# --- Custom CSS for Modern Orange Cash UI ---
st.markdown("""
<style>
    /* Main container padding */
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1200px;
    }
    
    /* Hero Banner */
    .hero-banner {
        background: linear-gradient(135deg, #1E1E1E 0%, #2D2D2D 60%, #FF7900 140%);
        padding: 1.8rem 2rem;
        border-radius: 14px;
        border-left: 6px solid #FF7900;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.12);
    }
    .hero-title {
        color: #FFFFFF;
        font-size: 1.9rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: 0.3px;
    }
    .hero-subtitle {
        color: #E0E0E0;
        font-size: 0.98rem;
        margin-top: 0.4rem;
        margin-bottom: 0;
    }
    .orange-accent {
        color: #FF7900;
    }

    /* Section Headers */
    .step-badge {
        display: inline-block;
        background-color: #FF7900;
        color: white;
        font-weight: 700;
        font-size: 0.8rem;
        padding: 0.2rem 0.65rem;
        border-radius: 20px;
        margin-bottom: 0.4rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* Metric Cards Styling */
    div[data-testid="metric-container"] {
        background-color: rgba(255, 121, 0, 0.06);
        border: 1px solid rgba(255, 121, 0, 0.25);
        padding: 0.9rem 1.1rem;
        border-radius: 10px;
    }

    /* Primary Buttons */
    .stButton > button[kind="primary"] {
        background-color: #FF7900;
        border-color: #FF7900;
        color: white;
        font-weight: 700;
        border-radius: 8px;
        padding: 0.55rem 1.5rem;
        transition: all 0.2s ease;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #E06800;
        border-color: #E06800;
        transform: translateY(-1px);
    }

    /* Download Button */
    .stDownloadButton > button {
        background-color: #198754;
        color: white;
        font-weight: 700;
        border-radius: 8px;
        border: none;
        padding: 0.65rem 1.5rem;
        width: 100%;
    }
    .stDownloadButton > button:hover {
        background-color: #157347;
        color: white;
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

def calculate_initial_shifts(n, total_target):
    if n == 0:
        return []
    
    if 15 * n == total_target:
        durations = [15] * n
    elif 15 * n < total_target:
        durations = [15] * n
        diff = total_target - (15 * n)
        for i in range(n - 1, -1, -1):
            if diff >= 5:
                durations[i] = 20
                diff -= 5
            elif diff > 0:
                durations[i] += diff
                diff = 0
            if diff == 0:
                break
        while diff > 0:
            for i in range(n - 1, -1, -1):
                durations[i] += 1
                diff -= 1
                if diff == 0:
                    break
    else:
        durations = [10] * n
        diff = total_target - (10 * n)
        if diff >= 0:
            for i in range(n - 1, -1, -1):
                if diff >= 5:
                    durations[i] = 15
                    diff -= 5
                elif diff > 0:
                    durations[i] += diff
                    diff = 0
                if diff == 0:
                    break
        else:
            while diff < 0:
                for i in range(n - 1, -1, -1):
                    if durations[i] > 3:
                        durations[i] -= 1
                        diff += 1
                        if diff == 0:
                            break
    return durations

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

# --- Excel Formatter Matching Exact Reference Image ---
def generate_styled_excel(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.sheets[sheet_name]

        # Styling definitions matching the reference screenshot
        header_fill = PatternFill(start_color="C65911", end_color="C65911", fill_type="solid")
        data_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

        header_font = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        data_font = Font(name="Times New Roman", size=11, bold=True, color="000000")

        center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)

        # Double border style around all cells
        double_side = Side(style="double", color="808080")
        cell_border = Border(left=double_side, right=double_side, top=double_side, bottom=double_side)

        # Style Header Row (Row 1)
        ws.row_dimensions[1].height = 24
        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_alignment
            cell.border = cell_border

        # Style Data Rows (Row 2+)
        for row_idx in range(2, len(df) + 2):
            ws.row_dimensions[row_idx].height = 21
            for col_idx in range(1, len(df.columns) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.fill = data_fill
                cell.font = data_font
                cell.alignment = center_alignment
                cell.border = cell_border

        # Auto-fit column widths with generous padding so text never clips
        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = len(str(col_name))
            for val in df[col_name]:
                val_len = len(str(val))
                if val_len > max_len:
                    max_len = val_len
            ws.column_dimensions[col_letter].width = max(max_len + 6, 13)

    return output.getvalue()

# --- Header UI ---
st.markdown("""
<div class="hero-banner">
    <h1 class="hero-title">⚡ <span class="orange-accent">Orange Cash</span> — Inbox Shift Automator</h1>
    <p class="hero-subtitle">Smart shift distribution, automatic agent detection, and Outlook-ready formatted schedules.</p>
</div>
""", unsafe_allow_html=True)

# --- Step 1: Schedule Mode ---
st.markdown('<span class="step-badge">Step 1 • Shift Configuration</span>', unsafe_allow_html=True)
mode_col1, mode_col2 = st.columns([2, 1])

with mode_col1:
    shift_mode = st.radio(
        "Select Target Schedule Mode:",
        options=["Males (8:01 PM - 3:00 AM)", "Females (8:00 AM - 8:00 PM)"],
        horizontal=True
    )

is_females = shift_mode.startswith("Females")

if is_females:
    target_teams = [
        "CS - (Orange Cash) Females WFH",
        "CS - (Orange Cash) Females NC"
    ]
    start_mins_const = 480     # 8:00 AM
    end_mins_const = 1200      # 8:00 PM
    total_target_const = 720   # 720 mins
    start_time_label = "8:00 AM"
    end_time_label = "8:00 PM"
    file_prefix = "Females"
else:
    target_teams = [
        "CS - (Orange Cash) Males WFH",
        "CS - (Orange Cash) Males NC"
    ]
    start_mins_const = 1201    # 8:01 PM
    end_mins_const = 1620      # 3:00 AM (next day)
    total_target_const = 420   # 420 mins
    start_time_label = "8:01 PM"
    end_time_label = "3:00 AM"
    file_prefix = "Males"

with mode_col2:
    st.info(f"🕒 **Coverage Window:** {start_time_label} ➔ {end_time_label}")

st.markdown("---")

# --- Step 2: Files Upload & Master List ---
st.markdown('<span class="step-badge">Step 2 • Data Sources</span>', unsafe_allow_html=True)
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

        if "custom_agents" in st.session_state and st.session_state["custom_agents"] is not None:
            pass

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
    team_col = cols_lower.get('team')
    state_col = cols_lower.get('schedule state')
    pid_col = cols_lower.get('pid')
    agent_col = cols_lower.get('agent') or cols_lower.get('agent name')
    start_col = cols_lower.get('start time') or cols_lower.get('shift')

    missing = [name for name, col in [('Team', team_col), ('Schedule State', state_col), ('Pid', pid_col), ('Agent', agent_col), ('Start Time', start_col)] if not col]
    if missing:
        st.error(f"⚠️ Missing required columns in the Attendance file: **{', '.join(missing)}**")
        st.stop()

    # 2. Filter Team column
    df_att[team_col] = df_att[team_col].astype(str).str.strip()
    df_filtered = df_att[df_att[team_col].isin(target_teams)].copy()

    # 3. Filter Schedule State column
    df_filtered[state_col] = df_filtered[state_col].astype(str).str.strip()
    df_filtered = df_filtered[df_filtered[state_col] == "Orange Cash"].copy()

    # 4. Sort Start Time ascending
    df_filtered['_sort_time'] = pd.to_datetime(df_filtered[start_col].astype(str).str.strip(), errors='coerce')
    df_filtered = df_filtered.sort_values(by=['_sort_time', pid_col], ascending=[True, True]).copy()

    # 5. Match with Master Agents file & fetch Profile (INS) if available
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

    st.markdown("---")
    st.markdown('<span class="step-badge">Step 3 • Exclusions & Generation</span>', unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown(f"#### 🔍 Detected {file_prefix} Inbox Agents ({len(eligible_df)} Eligible)")

        if len(eligible_df) == 0:
            st.warning("No matching agents found for the selected mode and Attendance file. Please check the selected mode (Males/Females).")
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
                placeholder="Choose agents to exclude (leave empty if all detected agents are working)..."
            )
            excluded_pids = {option_to_pid[opt] for opt in excluded_options}

            gen_col1, gen_col2 = st.columns([1, 2])
            with gen_col1:
                generate_clicked = st.button("🚀 Generate Final Schedule", type="primary", use_container_width=True)
            with gen_col2:
                st.caption(f"Ready to distribute **{total_target_const} minutes** across **{len(eligible_df) - len(excluded_pids)} agents**.")

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

                # Calculate and distribute shifts
                n = len(res_df)
                durations = calculate_initial_shifts(n, total_target_const)

                current_mins = start_mins_const
                from_list, to_list = [], []
                for d in durations:
                    from_list.append(mins_to_time_str(current_mins))
                    to_list.append(mins_to_time_str(current_mins + d - 1))
                    current_mins += d

                res_df["From"] = from_list
                res_df["To"] = to_list
                st.session_state["schedule_df"] = res_df
                st.session_state["end_mins_lock"] = end_mins_const
                st.session_state["start_label_lock"] = start_time_label
                st.session_state["end_label_lock"] = end_time_label
                st.session_state["file_prefix"] = file_prefix
                st.session_state["excluded_count"] = len(excluded_pids)

# --- Step 4: Final Schedule Display, Metrics, Emergency Update & Export ---
if "schedule_df" in st.session_state:
    df_res = st.session_state["schedule_df"]
    active_end_mins = st.session_state.get("end_mins_lock", end_mins_const)
    active_start_label = st.session_state.get("start_label_lock", start_time_label)
    active_end_label = st.session_state.get("end_label_lock", end_time_label)
    active_prefix = st.session_state.get("file_prefix", file_prefix)
    excluded_cnt = st.session_state.get("excluded_count", 0)
    
    st.markdown("---")
    st.markdown('<span class="step-badge">Step 4 • Final Output & Export</span>', unsafe_allow_html=True)
    st.subheader(f"📋 Final {active_prefix} Inbox Schedule")

    # Summary Metrics Row
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("👥 Scheduled Agents", f"{len(df_res)} Agents")
    m2.metric("🚫 Excluded (Task/Assist)", f"{excluded_cnt} Agents")
    m3.metric("🟢 Start Time", active_start_label)
    m4.metric("🔒 Locked End Time", active_end_label)

    st.write("")
    st.dataframe(df_res, use_container_width=True, hide_index=True)

    # Action Area: Download & Emergency Update side-by-side
    action_col1, action_col2 = st.columns([1, 1], gap="large")

    with action_col1:
        st.markdown("#### 📥 Export Formatted Excel")
        st.caption("Formatted with Times New Roman, Bold 11pt, centered cells, double borders, and auto-fit widths.")
        excel_bytes = generate_styled_excel(df_res, f"{active_prefix}_Schedule")
        st.download_button(
            label=f"📥 Download Final {active_prefix} Schedule (.xlsx)",
            data=excel_bytes,
            file_name=f"Final_{active_prefix}_Inbox_Schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with action_col2:
        if len(df_res) > 1:
            st.markdown("#### 🛠️ Emergency Redistribution")
            st.caption("Remove an agent (Sick / Left early) and automatically redistribute their minutes until the end.")
            with st.expander("Open Emergency Update Controls"):
                pids_list = df_res["PID"].astype(str).tolist()
                pid_to_label = {
                    str(row["PID"]): f"{row['PID']} - {row['Agent Name']}"
                    for _, row in df_res.iterrows()
                }
                
                sick_pid = st.selectbox(
                    "1. Select Agent to Remove (Sick / Left):",
                    options=pids_list,
                    format_func=lambda x: pid_to_label.get(x, x)
                )
                start_pid = st.selectbox(
                    "2. Start Redistribution From:",
                    options=pids_list,
                    format_func=lambda x: pid_to_label.get(x, x)
                )

                if st.button("⚡ Apply Update & Redistribute", type="primary", use_container_width=True):
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

                        # Lock end time at 3:00 AM (Males) or 8:00 PM (Females)
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

                        st.session_state["schedule_df"] = df_updated
                        st.session_state["excluded_count"] = excluded_cnt + 1
                        st.rerun()