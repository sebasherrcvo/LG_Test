import os
import sys
import subprocess
import re
import pandas as pd
import plotly.express as px
from datetime import datetime, time, timedelta

# --- AUTO-LAUNCHER LOGIC ---
try:
    import streamlit as st
    from streamlit.web import cli as stcli
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "streamlit", "pandas", "plotly"])
    import streamlit as st
    from streamlit.web import cli as stcli

def is_running_streamlit():
    return "streamlit" in sys.modules and st.runtime.exists()

if __name__ == "__main__":
    if not is_running_streamlit():
        sys.argv = ["streamlit", "run", sys.argv[0], "--server.maxUploadSize", "5000"]
        sys.exit(stcli.main())

# --- HELPER FUNCTIONS ---
def extract_numeric_suffix(text):
    """
    Specifically extracts the station number (e.g., extracts 13 from '_S13').
    If no S-number is found, it falls back to the first number found.
    """
    s_match = re.search(r'_S(\d+)', str(text))
    if s_match:
        return int(s_match.group(1))
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else 999

def sort_by_station_number(station_list):
    """Sorts the station list based on the S-suffix number."""
    return sorted(station_list, key=extract_numeric_suffix)

@st.cache_data(show_spinner="Optimizing Data Access...")
def load_data(file):
    dtype_map = {'mainprogram_name1': 'category', 'stepprogram_name1': 'category', 'station_name1': 'category'}
    cols = ['mainprogram_name1', 'stepprogram_name1', 'station_name1', 'total_cycle_time_secs1', 
            'cycle_end_utc1', 'step_start_utc1', 'cycle_number1']
    df = pd.read_csv(file, usecols=cols, dtype=dtype_map)
    df['step_start_utc1'] = pd.to_datetime(df['step_start_utc1']) - pd.Timedelta(hours=7)
    df['sv_tag'] = df['station_name1'].apply(lambda x: re.search(r'SV\d+', str(x)).group(0) if re.search(r'SV\d+', str(x)) else "Other")
    return df

def main():
    st.set_page_config(page_title="SV Analysis Pro", layout="wide")
    st.markdown("""<style>.stDeployButton { display: none !important; } footer { visibility: hidden; }</style>""", unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload CSV (Up to 5GB)", type="csv")

    if uploaded_file:
        df = load_data(uploaded_file)

        # --- SIDEBAR: GLOBAL FILTERS ---
        st.sidebar.header("Global Filters")
        progs = sorted(df['mainprogram_name1'].unique())
        selected_program = st.sidebar.selectbox("Main Program", progs)
        
        min_date, max_date = df['step_start_utc1'].min().date(), df['step_start_utc1'].max().date()
        selected_dates = st.sidebar.date_input("Date Range", value=(min_date, max_date))
        hour_range = st.sidebar.slider("Hour Range", value=(time(0, 0), time(23, 59)), format="HH:mm")

        all_svs = sorted(df['sv_tag'].unique(), key=lambda x: int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 999)
        select_all_sv = st.sidebar.checkbox("Select All SVs", value=True)
        selected_svs = all_svs if select_all_sv else st.sidebar.multiselect("Choose SVs", all_svs, default=all_svs)

        goal_time = st.sidebar.number_input("Goal (s)", value=120)
        time_filter = st.sidebar.slider("Cycle Time Noise Filter (s)", 0, 500, (70, 300))
        use_unique = st.sidebar.toggle("Enforce Unique Cycle IDs", value=True)

        # --- DATA PROCESSING PIPELINE ---
        if len(selected_dates) == 2:
            start_date, end_date = selected_dates
        else:
            start_date = end_date = selected_dates[0]

        mask = (df['mainprogram_name1'] == selected_program) & \
               (df['sv_tag'].isin(selected_svs)) & \
               (df['step_start_utc1'].dt.date >= start_date) & \
               (df['step_start_utc1'].dt.date <= end_date) & \
               (df['step_start_utc1'].dt.time >= hour_range[0]) & \
               (df['step_start_utc1'].dt.time <= hour_range[1])
        
        df_pipe = df[mask].copy()

        if use_unique:
            df_pipe = df_pipe.drop_duplicates(subset=['cycle_number1'])

        df_pipe = df_pipe[(df_pipe['total_cycle_time_secs1'] >= time_filter[0]) & 
                          (df_pipe['total_cycle_time_secs1'] <= time_filter[1])]

        # --- STATION MANAGER ---
        st.subheader("Station Visibility Manager")
        all_found_stations = df_pipe['station_name1'].unique().tolist()
        if 'ignored_stations' not in st.session_state: st.session_state.ignored_stations = set()

        f_col1, f_col2 = st.columns([3, 1])
        sel_sv_manager = f_col1.multiselect("Filter Manager by SV Group", options=all_svs)
        if f_col2.button("Reset Ignore List", use_container_width=True):
            st.session_state.ignored_stations = set()
            st.rerun()
        
        col_m1, col_m2 = st.columns(2)
        def apply_manager_filters(target_list):
            return [s for s in target_list if any(sv in str(s) for sv in sel_sv_manager)] if sel_sv_manager else target_list

        with col_m1:
            st.write("✅ **Considered (Active)**")
            consider_filtered = sort_by_station_number(apply_manager_filters([s for s in all_found_stations if s not in st.session_state.ignored_stations]))
            to_ignore = st.multiselect("Stations to Ignore:", consider_filtered)
            if st.button("Ignore ➔", disabled=not to_ignore):
                st.session_state.ignored_stations.update(to_ignore); st.rerun()

        with col_m2:
            st.write("❌ **Ignored**")
            ignored_filtered = sort_by_station_number(apply_manager_filters(list(st.session_state.ignored_stations)))
            to_consider = st.multiselect("Stations to Re-enable:", ignored_filtered)
            if st.button("⬅ Consider", disabled=not to_consider):
                for s in to_consider: st.session_state.ignored_stations.discard(s)
                st.rerun()

        # --- FINAL DATASET & SUMMARY ---
        df_final = df_pipe[~df_pipe['station_name1'].isin(st.session_state.ignored_stations)].copy()

        if not df_final.empty:
            summary = df_final.groupby(['station_name1', 'sv_tag'], observed=True)['total_cycle_time_secs1'].agg(['mean', 'count']).reset_index()
            summary['sort_key'] = summary['station_name1'].apply(extract_numeric_suffix)
            summary = summary.sort_values('sort_key')

            # --- METRICS ---
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Samples", len(df_final))
            m2.metric("Mean CT", f"{df_final['total_cycle_time_secs1'].mean():.1f}s")
            
            bottleneck_ct = summary['mean'].max()
            uph = 3600 / bottleneck_ct if bottleneck_ct > 0 else 0
            m3.metric("Estimated UPH", f"{uph:.1f}")
            m4.metric("Bottleneck CT", f"{bottleneck_ct:.1f}s")

            # Chart with Count in Hover and S01-S16 Sorting
            st.plotly_chart(px.bar(summary, x='station_name1', y='mean', color='sv_tag', 
                hover_data={'station_name1': False, 'mean': ':.2f', 'count': True},
                labels={'count': 'Sample Size', 'mean': 'Avg Cycle Time (s)'},
                template="plotly_dark", text_auto='.1f', title="Cycle Time by Station").add_hline(y=goal_time, line_color="#00FF00"), use_container_width=True)

            # --- DRILL DOWN ---
            st.markdown("---")
            st.subheader("Station Drill Down")
            curr_stations = sort_by_station_number(df_final['station_name1'].unique())
            sel_st = st.selectbox("Select Station to Analyze", curr_stations)
            
            detail = df_final[df_final['station_name1'] == sel_st]
            st_mean = detail['total_cycle_time_secs1'].mean()
            
            st.plotly_chart(px.line(detail, x='step_start_utc1', y='total_cycle_time_secs1', markers=True, 
                hover_data={'total_cycle_time_secs1': ':.2f', 'cycle_number1': True},
                template="plotly_dark", title=f"Timeline: {sel_st}").add_hline(y=goal_time, line_color="#00FF00"), use_container_width=True)
            
            st.plotly_chart(px.histogram(detail, x='total_cycle_time_secs1', nbins=30, 
                template="plotly_dark", title=f"Mean Distribution: {sel_st}").add_vline(x=st_mean, line_dash="dash", line_color="orange"), use_container_width=True)
        else:
            st.warning("No data matches the current criteria.")

if __name__ == "__main__":
    main()