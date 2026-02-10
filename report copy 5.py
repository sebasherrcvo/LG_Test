import os
import sys
import subprocess
import re
import io
import pandas as pd
import plotly.express as px
from datetime import datetime, time

# --- AUTO-LAUNCHER & INSTALLER ---
try:
    import streamlit as st
    from streamlit.web import cli as stcli
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "streamlit", "pandas", "plotly", "xlsxwriter"])
    import streamlit as st
    from streamlit.web import cli as stcli

def is_running_streamlit():
    return "streamlit" in sys.modules and st.runtime.exists()

# --- HELPER FUNCTIONS ---
def extract_numeric_suffix(text):
    """Extracts station number for logical sorting (e.g., S06 -> 6)."""
    s_match = re.search(r'_S(\d+)', str(text))
    if s_match:
        return int(s_match.group(1))
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else 999

def sort_by_station_number(station_list):
    return sorted(station_list, key=extract_numeric_suffix)

@st.cache_data(show_spinner="Loading data (Processing 5GB file)...")
def load_data(file):
    # Optimized dtypes for large file handling
    dtype_map = {
        'mainprogram_name1': 'category', 
        'stepprogram_name1': 'category', 
        'station_name1': 'category',
        'cycle_number1': 'int64',
        'total_cycle_time_secs1': 'float32'
    }
    cols = ['mainprogram_name1', 'stepprogram_name1', 'station_name1', 
            'total_cycle_time_secs1', 'cycle_end_utc1', 'step_start_utc1', 'cycle_number1']
    
    df = pd.read_csv(file, usecols=cols, dtype=dtype_map)
    # Timezone adjustment (Example: UTC to Local)
    df['step_start_utc1'] = pd.to_datetime(df['step_start_utc1']) - pd.Timedelta(hours=7)
    # Extract SV Unit (e.g., SV5)
    df['sv_tag'] = df['station_name1'].apply(
        lambda x: re.search(r'SV\d+', str(x)).group(0) if re.search(r'SV\d+', str(x)) else "Other"
    )
    return df

def convert_df_to_excel(df_final, summary_df):
    """Generates an Excel file in memory with Summary and Cleaned Raw Data."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, index=False, sheet_name='Summary_Stats')
        # Excel limit safety check
        if len(df_final) <= 1048500:
            df_final.to_excel(writer, index=False, sheet_name='Cleaned_Raw_Data')
        else:
            df_final.iloc[:1048500].to_excel(writer, index=False, sheet_name='Raw_Data_Truncated')
    return output.getvalue()

def main():
    st.set_page_config(page_title="Cycle Time Analytics", layout="wide")
    st.markdown("""<style>.stDeployButton { display: none !important; } footer { visibility: hidden; }</style>""", unsafe_allow_html=True)

    st.title("Station Cycle Time Analyzer")
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
        time_filter = st.sidebar.slider("Noise Filter (s)", 0, 1000, (70, 300))
        use_unique = st.sidebar.toggle("Enforce Unique Cycle IDs", value=True)

        # --- DATA PROCESSING ---
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
        else:
            start_date = end_date = selected_dates if not isinstance(selected_dates, (tuple, list)) else selected_dates[0]

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

        # --- STATION VISIBILITY MANAGER ---
        st.subheader("Station Visibility Manager")
        if 'ignored_stations' not in st.session_state: 
            st.session_state.ignored_stations = set()

        all_found_stations = df_pipe['station_name1'].unique().tolist()
        col_m1, col_m2 = st.columns(2)
        
        with col_m1:
            st.write("✅ **Active Stations**")
            active_list = sort_by_station_number([s for s in all_found_stations if s not in st.session_state.ignored_stations])
            to_ignore = st.multiselect("Select stations to hide:", active_list)
            if st.button("Hide Selected ➔"):
                st.session_state.ignored_stations.update(to_ignore)
                st.rerun()

        with col_m2:
            st.write("❌ **Hidden Stations**")
            hidden_list = sort_by_station_number(list(st.session_state.ignored_stations))
            to_active = st.multiselect("Select stations to show:", hidden_list)
            if st.button("⬅ Show Selected"):
                for s in to_active: st.session_state.ignored_stations.discard(s)
                st.rerun()
        
        if st.button("Reset All Hidden Stations"):
            st.session_state.ignored_stations = set()
            st.rerun()

        # --- FINAL DATASET & SUMMARY ---
        df_final = df_pipe[~df_pipe['station_name1'].isin(st.session_state.ignored_stations)].copy()

        if not df_final.empty:
            # Aggregation by Median
            summary = df_final.groupby(['station_name1', 'sv_tag'], observed=True)['total_cycle_time_secs1'].agg(['median', 'count']).reset_index()
            summary['sort_key'] = summary['station_name1'].apply(extract_numeric_suffix)
            summary = summary.sort_values('sort_key')

            # --- EXPORT ---
            excel_file = convert_df_to_excel(df_final, summary)
            st.sidebar.download_button(label="📥 Download Excel Report", data=excel_file, 
                                       file_name=f"CT_Report_{datetime.now().strftime('%H%M%S')}.xlsx", 
                                       use_container_width=True)

            # --- METRICS (WITH 15% BUFFER) ---
            m1, m2, m3 = st.columns(3)
            m1.metric("Total Samples", f"{len(df_final):,}")
            
            raw_bottleneck = summary['median'].max()
            # Adding 15% increase to the final bottleneck
            bottleneck_ct = raw_bottleneck * 1.15
            # Calculating UPH based on the buffered bottleneck
            uph = 3600 / bottleneck_ct if bottleneck_ct > 0 else 0
            
            m2.metric("Est. UPH (+15% Buffer)", f"{uph:.1f}")
            m3.metric("Bottleneck Median CT (+15%)", f"{bottleneck_ct:.1f}s", delta=f"Raw: {raw_bottleneck:.1f}s")

            # --- BAR CHART ---
            fig_bar = px.bar(
                summary, 
                x='station_name1', 
                y='median', 
                color='sv_tag', 
                hover_data={'station_name1': True, 'median': ':.2f', 'count': True},
                labels={'station_name1': 'Station', 'count': 'Samples', 'median': 'Median CT (s)'},
                template="plotly_dark", 
                text_auto='.1f', 
                title="Median Cycle Time by Station"
            )
            fig_bar.add_hline(y=goal_time, line_color="#00FF00", annotation_text="Goal")
            # Visualizing the buffered bottleneck line
            fig_bar.add_hline(y=bottleneck_ct, line_dash="dash", line_color="orange", annotation_text="Buffered Bottleneck")
            st.plotly_chart(fig_bar, use_container_width=True)

            # --- DRILL DOWN ---
            st.markdown("---")
            sel_st = st.selectbox("Analyze Timeline for:", sort_by_station_number(df_final['station_name1'].unique()))
            detail = df_final[df_final['station_name1'] == sel_st]
            st.plotly_chart(px.line(detail, x='step_start_utc1', y='total_cycle_time_secs1', markers=True, 
                template="plotly_dark", title=f"Timeline: {sel_st}"), use_container_width=True)

        else:
            st.warning("No data matches the current criteria.")

if __name__ == "__main__":
    if not is_running_streamlit():
        # Auto-configure server for large file uploads
        sys.argv = ["streamlit", "run", sys.argv[0], "--server.maxUploadSize", "5000"]
        sys.exit(stcli.main())
    else:
        main()