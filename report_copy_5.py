import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
from datetime import datetime, time

# --- CONFIG ---
st.set_page_config(page_title="Cycle Time Analytics", layout="wide")

# Custom CSS to hide default Streamlit elements
st.markdown("""
    <style>
    .stDeployButton { display: none !important; } 
    footer { visibility: hidden; }
    </style>
    """, unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---
def extract_numeric_suffix(text):
    s_match = re.search(r'_S(\d+)', str(text))
    if s_match: return int(s_match.group(1))
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else 999

def sort_by_station_number(station_list):
    return sorted(station_list, key=extract_numeric_suffix)

@st.cache_data(show_spinner="Unpacking Parquet Data...")
def load_data(file):
    # Using read_parquet uses significantly less RAM than CSV
    df = pd.read_parquet(file)
    
    if not pd.api.types.is_datetime64_any_dtype(df['step_start_utc1']):
        df['step_start_utc1'] = pd.to_datetime(df['step_start_utc1'])
    
    # Adjust for local timezone (Example: -7 hours)
    df['step_start_utc1'] = df['step_start_utc1'] - pd.Timedelta(hours=7)
    
    # SV Tag extraction logic
    df['sv_tag'] = df['station_name1'].apply(
        lambda x: re.search(r'SV\d+', str(x)).group(0) if re.search(r'SV\d+', str(x)) else "Other"
    )
    return df

def main():
    st.title("Station Cycle Time Analyzer")
    
    # --- THIS IS THE FIX FOR YOUR IMAGE ---
    # Changing type to "parquet" allows the browser to see those files
    uploaded_file = st.file_uploader("Upload Data (Parquet Format)", type=["parquet"])

    if uploaded_file:
        df = load_data(uploaded_file)

        # --- SIDEBAR FILTERS ---
        st.sidebar.header("Global Filters")
        progs = sorted(df['mainprogram_name1'].unique())
        selected_program = st.sidebar.selectbox("Main Program", progs)
        
        min_date, max_date = df['step_start_utc1'].min().date(), df['step_start_utc1'].max().date()
        selected_dates = st.sidebar.date_input("Date Range", value=(min_date, max_date))
        hour_range = st.sidebar.slider("Hour Range", value=(time(0, 0), time(23, 59)), format="HH:mm")

        goal_time = st.sidebar.number_input("Goal (s)", value=120)
        time_filter = st.sidebar.slider("Noise Filter (s)", 0, 1000, (70, 300))

        # --- DATA FILTERING ---
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
        else:
            start_date = end_date = selected_dates

        mask = (df['mainprogram_name1'] == selected_program) & \
               (df['step_start_utc1'].dt.date >= start_date) & \
               (df['step_start_utc1'].dt.date <= end_date) & \
               (df['step_start_utc1'].dt.time >= hour_range[0]) & \
               (df['step_start_utc1'].dt.time <= hour_range[1])
        
        df_filtered = df[mask].copy()
        df_filtered = df_filtered[(df_filtered['total_cycle_time_secs1'] >= time_filter[0]) & 
                                  (df_filtered['total_cycle_time_secs1'] <= time_filter[1])]

        # --- VISIBILITY MANAGER ---
        if 'ignored_stations' not in st.session_state: 
            st.session_state.ignored_stations = set()

        all_stations = df_filtered['station_name1'].unique().tolist()
        active_list = sort_by_station_number([s for s in all_stations if s not in st.session_state.ignored_stations])
        
        st.subheader("Station Visibility Manager")
        to_hide = st.multiselect("Select stations to hide:", active_list)
        if st.button("Hide Selected"):
            st.session_state.ignored_stations.update(to_hide)
            st.rerun()

        # --- FINAL METRICS & CHART ---
        df_final = df_filtered[~df_filtered['station_name1'].isin(st.session_state.ignored_stations)].copy()

        if not df_final.empty:
            summary = df_final.groupby(['station_name1', 'sv_tag'], observed=True)['total_cycle_time_secs1'].agg(['median', 'count']).reset_index()
            summary['sort_key'] = summary['station_name1'].apply(extract_numeric_suffix)
            summary = summary.sort_values('sort_key')

            raw_bottleneck = summary['median'].max()
            bottleneck_buffered = raw_bottleneck * 1.15
            uph = 3600 / bottleneck_buffered if bottleneck_buffered > 0 else 0
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Samples", f"{len(df_final):,}")
            m2.metric("Est. UPH (+15%)", f"{uph:.1f}")
            m3.metric("Bottleneck (+15%)", f"{bottleneck_buffered:.1f}s")

            fig = px.bar(summary, x='station_name1', y='median', color='sv_tag', text_auto='.1f', template="plotly_dark")
            fig.add_hline(y=goal_time, line_color="green", annotation_text="Goal")
            fig.add_hline(y=bottleneck_buffered, line_dash="dash", line_color="orange")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("No data matches selected filters.")

if __name__ == "__main__":
    main()
