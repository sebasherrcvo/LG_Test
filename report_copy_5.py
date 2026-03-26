import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
from datetime import datetime, time

# --- CONFIG ---
st.set_page_config(page_title="Cycle Time Analytics", layout="wide")

# Hide Streamlit's default UI elements
st.markdown("""
    <style>
    .stDeployButton { display: none !important; } 
    footer { visibility: hidden; }
    </style>
    """, unsafe_allow_html=True)

# --- HELPER FUNCTIONS ---
def extract_numeric_suffix(text):
    """Extracts station number for logical sorting (e.g., S06 -> 6)."""
    s_match = re.search(r'S(\d+)', str(text))
    if s_match: return int(s_match.group(1))
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else 999

def sort_by_station_number(station_list):
    return sorted(station_list, key=extract_numeric_suffix)

@st.cache_data(show_spinner="Loading Parquet Data...")
def load_data(file):
    df = pd.read_parquet(file)
    if not pd.api.types.is_datetime64_any_dtype(df['step_start_utc1']):
        df['step_start_utc1'] = pd.to_datetime(df['step_start_utc1'])
    
    # Timezone adjustment (UTC to Local)
    df['step_start_utc1'] = df['step_start_utc1'] - pd.Timedelta(hours=7)
    
    # Extract Unit/Line (e.g., SV5)
    df['sv_tag'] = df['station_name1'].apply(
        lambda x: re.search(r'SV\d+', str(x)).group(0) if re.search(r'SV\d+', str(x)) else "Other"
    )
    
    # Extract Base Station (e.g., S02) for cross-line aggregation
    df['base_station'] = df['station_name1'].apply(
        lambda x: re.search(r'S\d+', str(x)).group(0) if re.search(r'S\d+', str(x)) else "Unknown"
    )
    
    return df

def convert_df_to_excel(df_final, summary_df, cross_summary):
    """Generates an Excel download with multiple sheets."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, index=False, sheet_name='By_Line_Station')
        cross_summary.to_excel(writer, index=False, sheet_name='Cross_Line_Aggregated')
        df_final.iloc[:1000000].to_excel(writer, index=False, sheet_name='Cleaned_Raw_Data')
    return output.getvalue()

def main():
    st.title("Station Cycle Time Analyzer")
    st.info("Upload your compressed .parquet file to begin analysis.")
    
    uploaded_file = st.file_uploader("Upload Parquet File", type=["parquet"])

    if uploaded_file:
        df = load_data(uploaded_file)

        # --- SIDEBAR: GLOBAL FILTERS ---
        st.sidebar.header("Global Filters")
        
        # 1. Main Program Filter
        progs = sorted(df['mainprogram_name1'].unique())
        selected_program = st.sidebar.selectbox("Main Program", progs)
        
        # 2. Line Filter (SV Tags) - RE-ENABLED
        available_lines = sorted(df['sv_tag'].unique())
        selected_lines = st.sidebar.multiselect(
            "Select Lines (SV tags)", 
            options=available_lines, 
            default=available_lines
        )
        
        # 3. Date Range Filter
        min_date, max_date = df['step_start_utc1'].min().date(), df['step_start_utc1'].max().date()
        selected_dates = st.sidebar.date_input("Date Range", value=(min_date, max_date))
        
        # 4. Other Operational Filters
        hour_range = st.sidebar.slider("Hour Range", value=(time(0, 0), time(23, 59)), format="HH:mm")
        goal_time = st.sidebar.number_input("Goal (seconds)", value=120)
        time_filter = st.sidebar.slider("Cycle Time Range Filter (s)", 0, 1000, (70, 500))

        # --- DATA FILTERING LOGIC ---
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
        else:
            start_date = end_date = selected_dates

        # Apply the logic including the Line selection
        mask = (df['mainprogram_name1'] == selected_program) & \
               (df['sv_tag'].isin(selected_lines)) & \
               (df['step_start_utc1'].dt.date >= start_date) & \
               (df['step_start_utc1'].dt.date <= end_date) & \
               (df['step_start_utc1'].dt.time >= hour_range[0]) & \
               (df['step_start_utc1'].dt.time <= hour_range[1])
        
        df_filtered = df[mask].copy()
        
        # Secondary filter for specific CT ranges
        df_filtered = df_filtered[(df_filtered['total_cycle_time_secs1'] >= time_filter[0]) & 
                                  (df_filtered['total_cycle_time_secs1'] <= time_filter[1])]

        # --- STATION VISIBILITY MANAGER ---
        if 'ignored_stations' not in st.session_state: 
            st.session_state.ignored_stations = set()

        all_stations = df_filtered['station_name1'].unique().tolist()
        st.subheader("Station Visibility Manager")
        active_list = sort_by_station_number([s for s in all_stations if s not in st.session_state.ignored_stations])
        to_hide = st.multiselect("Select stations to hide from charts:", active_list)
        
        col_btn1, col_btn2 = st.columns([1, 5])
        if col_btn1.button("Hide Selected"):
            st.session_state.ignored_stations.update(to_hide)
            st.rerun()
        if col_btn2.button("Reset All Stations"):
            st.session_state.ignored_stations = set()
            st.rerun()

        # --- CALCULATE FINAL DATASET ---
        df_final = df_filtered[~df_filtered['station_name1'].isin(st.session_state.ignored_stations)].copy()

        if not df_final.empty:
            # 1. Summary by specific Station Name (e.g., SV1_S02)
            summary = df_final.groupby(['station_name1', 'sv_tag'], observed=True)['total_cycle_time_secs1'].agg(['median', 'count']).reset_index()
            summary['sort_key'] = summary['station_name1'].apply(extract_numeric_suffix)
            summary = summary.sort_values('sort_key')

            # --- BOTTLENECK CALCULATIONS ---
            raw_bottleneck = summary['median'].max()
            bottleneck_buffered = raw_bottleneck * 1.15
            uph = 3600 / bottleneck_buffered if bottleneck_buffered > 0 else 0
            
            # --- DISPLAY METRICS ---
            m1, m2, m3 = st.columns(3)
            m1.metric("Samples Count", f"{len(df_final):,}")
            m2.metric("Est. UPH (+15% Buffer)", f"{uph:.1f}")
            m3.metric("Bottleneck CT (+15%)", f"{bottleneck_buffered:.1f}s")

            # --- VISUALIZATION: BY LINE ---
            fig_bar = px.bar(
                summary, 
                x='station_name1', 
                y='median', 
                color='sv_tag', 
                text_auto='.1f', 
                title="Median Cycle Time per Station (By Line)",
                template="plotly_dark",
                labels={'median': 'Median CT (s)', 'station_name1': 'Station'}
            )
            fig_bar.add_hline(y=goal_time, line_color="green", annotation_text="Goal")
            fig_bar.add_hline(y=bottleneck_buffered, line_dash="dash", line_color="orange", annotation_text="Buffered Bottleneck")
            st.plotly_chart(fig_bar, use_container_width=True)

            # --- CROSS-LINE AGGREGATION ---
            st.markdown("---")
            st.subheader("Station Cycle Time Across All Selected Lines")
            
            # Grouping by the base station ID (e.g., S01, S02)
            cross_line_summary = df_final.groupby('base_station')['total_cycle_time_secs1'].agg(
                Median_CT='median',
                Average_CT='mean',
                Std_Dev='std',
                Sample_Size='count'
            ).reset_index()

            # Sort logically
            cross_line_summary['sort_key'] = cross_line_summary['base_station'].apply(extract_numeric_suffix)
            cross_line_summary = cross_line_summary.sort_values('sort_key').drop(columns=['sort_key'])

            # Visual: Cross-Line Bar Graph
            fig_cross = px.bar(
                cross_line_summary,
                x='base_station',
                y='Median_CT',
                text_auto='.1f',
                title="Aggregated Median Cycle Time per Station ID",
                template="plotly_dark",
                labels={'Median_CT': 'Global Median CT (s)', 'base_station': 'Station ID'}
            )
            fig_cross.update_traces(marker_color='steelblue')
            fig_cross.add_hline(y=goal_time, line_color="green", annotation_text="Goal")
            st.plotly_chart(fig_cross, use_container_width=True)

            # Table: Cross-Line Data
            st.dataframe(
                cross_line_summary.style.format({
                    'Median_CT': '{:.2f}s',
                    'Average_CT': '{:.2f}s',
                    'Std_Dev': '{:.2f}',
                    'Sample_Size': '{:,}'
                }),
                use_container_width=True
            )

            # --- EXCEL EXPORT ---
            excel_file = convert_df_to_excel(df_final, summary, cross_line_summary)
            st.download_button(
                label="Download Excel Report", 
                data=excel_file, 
                file_name=f"CT_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No data matches your current filters. Adjust the range or program.")

if __name__ == "__main__":
    main()
