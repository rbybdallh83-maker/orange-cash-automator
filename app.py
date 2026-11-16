import streamlit as st
import pandas as pd
import io
import re
import os

st.set_page_config(page_title="Orange Cash - Inbox Automator", layout="wide")

DEFAULT_AGENTS_FILE = "Agents Included in daily inbox.xlsx"

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

# --- App UI ---
st.title("⚡ Orange Cash - Inbox Shift Automator")

shift_mode = st.radio(
    "Select Schedule Mode:",
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
    total_target_const = 720   # 720 mins (8:00 AM to 8:00 PM)
    end_time_label = "8:00 PM"
    file_prefix = "Females"
else:
    target_teams = [
        "CS - (Orange Cash) Males WFH",
        "CS - (Orange Cash) Males NC"
    ]
    start_mins_const = 1201    # 8:01 PM
    end_mins_const = 1620      # 3:00 AM (next day)
    total_target_const = 420   # 420 mins (8:01 PM to 3:00 AM)
    end_time_label = "3:00 AM"
    file_prefix = "Males"

st.markdown(f"Upload the Attendance file to automatically detect eligible Inbox agents, mark any **Assist / Task** exclusions, and lock the schedule until **{end_time_label}**.")

col1, col2 = st.columns(2)

with col1:
    att_file = st.file_uploader("1️⃣ Upload Attendance File", type=["xlsx", "xls", "csv"])

with col2:
    st.markdown("**2️⃣ Agents Included in Daily Inbox (Master List)**")
    custom_agents_file = st.file_uploader(
        "Optional: Upload updated Agents file to override default",
        type=["xlsx", "xls", "csv"],
        key="custom_agents"
    )
    
    df_agents = None
    agents_source_label = ""
    if custom_agents_file is not None:
        df_agents = load_df(custom_agents_file)
        agents_source_label = f"Uploaded Update ({custom_agents_file.name})"
    elif os.path.exists(DEFAULT_AGENTS_FILE):
        df_agents = load_df(DEFAULT_AGENTS_FILE)
        agents_source_label = f"Default Built-in File ({len(df_agents)} Agents)"
    else:
        st.error(f"Default file '{DEFAULT_AGENTS_FILE}' not found. Please upload it above.")

    if df_agents is not None:
        st.caption(f"✅ Active Master List: **{agents_source_label}**")
        with st.expander("👁️ View / Inspect Master Agents List"):
            st.dataframe(df_agents, use_container_width=True, hide_index=True, height=250)

# --- Automatic Detection upon Uploading Attendance File ---
eligible_df = None
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
        st.error(f"Missing required columns in the Attendance file: {', '.join(missing)}")
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

    # --- 3. Interactive Selection for Assist / Task Exclusions ---
    st.markdown("---")
    st.subheader(f"3️⃣ Detected Eligible {file_prefix} Inbox Agents ({len(eligible_df)} Found)")

    if len(eligible_df) == 0:
        st.warning("No matching agents found for the selected mode and Attendance file.")
    else:
        options_list = [
            f"{row['Agent_clean']} ({row['PID_clean']}) - Shift: {row['Shift_clean']}"
            for _, row in eligible_df.iterrows()
        ]
        option_to_pid = {
            f"{row['Agent_clean']} ({row['PID_clean']}) - Shift: {row['Shift_clean']}": row['PID_clean']
            for _, row in eligible_df.iterrows()
        }

        excluded_options = st.multiselect(
            "Select agents on Assist / Task to EXCLUDE from the schedule:",
            options=options_list,
            placeholder="Click here to choose agents to exclude (leave empty if none)..."
        )
        excluded_pids = {option_to_pid[opt] for opt in excluded_options}

        if st.button("🚀 Generate Final Schedule", type="primary"):
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

            res_df["from"] = from_list
            res_df["to"] = to_list
            st.session_state["schedule_df"] = res_df
            st.session_state["end_mins_lock"] = end_mins_const
            st.session_state["file_prefix"] = file_prefix

# --- Display Table, Emergency Update & Export ---
if "schedule_df" in st.session_state:
    df_res = st.session_state["schedule_df"]
    active_end_mins = st.session_state.get("end_mins_lock", end_mins_const)
    active_prefix = st.session_state.get("file_prefix", file_prefix)
    
    st.markdown("---")
    st.subheader(f"📋 Final {active_prefix} Schedule ({len(df_res)} Agents)")
    st.dataframe(df_res, use_container_width=True, hide_index=True)

    if len(df_res) > 1:
        with st.expander("🛠️ Emergency Update (Remove Sick / Left Agent & Redistribute Until End of Schedule)"):
            c1, c2 = st.columns(2)
            pids_list = df_res["PID"].astype(str).tolist()
            
            sick_pid = c1.selectbox("Select Excluded PID (Sick / Left):", pids_list)
            start_pid = c2.selectbox("Start Range From PID:", pids_list)

            if st.button("⚡ Apply Update & Redistribute"):
                sick_idx = df_res.index[df_res["PID"].astype(str) == sick_pid].tolist()[0]
                start_idx = df_res.index[df_res["PID"].astype(str) == start_pid].tolist()[0]
                end_idx = len(df_res) - 1  # Automatically ends at the last agent

                if start_idx == end_idx and sick_idx == end_idx:
                    st.error("Error: The excluded agent is the last agent in the schedule with no subsequent agents to redistribute to!")
                else:
                    effective_start_idx = min(start_idx, sick_idx) if start_idx > sick_idx else start_idx
                    
                    s_from = time_str_to_abs_mins(df_res.loc[sick_idx, "from"])
                    s_to = time_str_to_abs_mins(df_res.loc[sick_idx, "to"])
                    if s_to < s_from: s_to += 1440
                    freed_mins = s_to - s_from + 1

                    base_start_mins = time_str_to_abs_mins(df_res.loc[effective_start_idx, "from"])

                    total_target_mins = freed_mins
                    for i in range(effective_start_idx, end_idx + 1):
                        if i != sick_idx:
                            r_from = time_str_to_abs_mins(df_res.loc[i, "from"])
                            r_to = time_str_to_abs_mins(df_res.loc[i, "to"])
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
                        df_updated.loc[r_idx, "from"] = mins_to_time_str(curr)
                        df_updated.loc[r_idx, "to"] = mins_to_time_str(curr + d - 1)
                        curr += d

                    st.session_state["schedule_df"] = df_updated
                    st.rerun()

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_res.to_excel(writer, index=False, sheet_name=f'{active_prefix}_Schedule')
        
        st.download_button(
            label=f"📥 Download Final {active_prefix} Excel (Ready for Outlook)",
            data=output.getvalue(),
            file_name=f"Final_{active_prefix}_Inbox_Schedule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )