import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
import gc 
from datetime import datetime, time

# --- CONFIG ---
st.set_page_config(page_title="Cycle Time Analytics", layout="wide")

# --- MEMORY-OPTIMIZED HELPER FUNCTIONS ---
def extract_numeric_suffix(text):
    s_match = re.search(r'_S(\d+)', str(text))
    if s_match: return int(s_match.group(1))
    match = re.search(r'(\d+)', str(text))
    return int(match.group(1)) if match else 999

def sort_by_station_number(station_list):
    return sorted(station_list, key=extract_numeric_suffix)

@st.cache_data(show_spinner="Unpacking Parquet Data...")
def load_data(file):
    df = pd.read_parquet(file)
    
    # Enforce Unique Cycle IDs
    unique_cols = ['mainprogram_name1', 'station_name1', 'cycle_number1']
    df = df.drop_duplicates(subset=unique_cols, keep='last')
    
    # Memory Optimization: Categorical types reduce RAM usage by up to 90%
    for col in ['mainprogram_name1', 'stepprogram_name1', 'station_name1']:
        if col in df.columns:
            df[col] = df[col].astype('category')
    
    if not pd.api.types.is_datetime64_any_dtype(df['step_start_utc1']):
        df['step_start_utc1'] = pd.to_datetime(df['step_start_utc1'])
    
    df['step_start_utc1'] = df['step_start_utc1'] - pd.Timedelta(hours=7)
    
    # Extract SV Tag
    df['sv_tag'] = df['station_name1'].astype(str).apply(
        lambda x: re.search(r'SV\d+', x).group(0) if re.search(r'SV\d+', x) else "Other"
    ).astype('category')
    
    return df

def convert_df_to_excel(df_final, summary_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, index=False, sheet_name='Summary_Stats')
        df_final.iloc[:1000000].to_excel(writer, index=False, sheet_name='Cleaned_Raw_Data')
    return output.getvalue()

def main():
    st.title("Station Cycle Time Analyzer")
    
    # FIX: Changed type to "parquet" to allow file selection and upload
    uploaded_file = st.file_uploader("Upload Data (Parquet Format)", type=["parquet"])

    if uploaded_file:
        df = load_data(uploaded_file)

        # --- SIDEBAR FILTERS ---
        st.sidebar.header("Global Filters")
        
        # 1. Main Program
        progs = sorted(df['mainprogram_name1'].unique())
        selected_program = st.sidebar.selectbox("Main Program", progs)
        
        # 2. SV Selection (New Multi-select)
        all_svs = sorted(df['sv_tag'].unique())
        selected_svs = st.sidebar.multiselect("Select SVs to Analyze", all_svs, default=all_svs)
        
        # 3. Date & Time
        min_date, max_date = df['step_start_utc1'].min().date(), df['step_start_utc1'].max().date()
        selected_dates = st.sidebar.date_input("Date Range", value=(min_date, max_date))
        hour_range = st.sidebar.slider("Hour Range", value=(time(0, 0), time(23, 59)), format="HH:mm")

        # 4. Noise Filter (Slider + Manual Type-in)
        st.sidebar.subheader("Noise Filter (Seconds)")
        c_min, c_max = st.sidebar.columns(2)
        n_min = c_min.number_input("Min", value=70)
        n_max = c_max.number_input("Max", value=300)
        
        # Slider synced to manual inputs for intuitive control
        noise_range = st.sidebar.slider("Fine Range Adjustment", 0, 1000, (int(n_min), int(n_max)))

        goal_time = st.sidebar.number_input("Goal (s)", value=120)

        # --- DATA FILTERING ---
        if isinstance(selected_dates, (tuple, list)) and len(selected_dates) == 2:
            start_date, end_date = selected_dates
        else:
            start_date = end_date = selected_dates

        mask = (df['mainprogram_name1'] == selected_program) & \
               (df['sv_tag'].isin(selected_svs)) & \
               (df['step_start_utc1'].dt.date >= start_date) & \
               (df['step_start_utc1'].dt.date <= end_date) & \
               (df['step_start_utc1'].dt.time >= hour_range[0]) & \
               (df['step_start_utc1'].dt.time <= hour_range[1])
        
        df_filtered = df[mask].copy()
        df_filtered = df_filtered[(df_filtered['total_cycle_time_secs1'] >= noise_range[0]) & 
                                  (df_filtered['total_cycle_time_secs1'] <= noise_range[1])]

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

        if st.button("Reset Visibility"):
            st.session_state.ignored_stations = set()
            st.rerun()

        # Final Dataset (Unique cycles only)
        df_final = df_filtered[~df_filtered['station_name1'].isin(st.session_state.ignored_stations)].copy()

        if not df_final.empty:
            df_final['station_name1'] = df_final['station_name1'].cat.remove_unused_categories()
            
            summary = df_final.groupby(['station_name1', 'sv_tag'], observed=True)['total_cycle_time_secs1'].agg(['median', 'count']).reset_index()
            summary['sort_key'] = summary['station_name1'].apply(extract_numeric_suffix)
            summary = summary.sort_values('sort_key')

            total_samples = int(summary['count'].sum()) 
            raw_bottleneck = summary['median'].max()
            bottleneck_buffered = raw_bottleneck * 1.15
            uph = 3600 / bottleneck_buffered if bottleneck_buffered > 0 else 0
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Unique Cycles (Filtered)", f"{total_samples:,}")
            m2.metric("Est. UPH (+15%)", f"{uph:.1f}")
            m3.metric("Bottleneck (+15%)", f"{bottleneck_buffered:.1f}s")

            # --- PLOTLY WITH SAMPLES IN HOVER ---
            fig = px.bar(
                summary, 
                x='station_name1', 
                y='median', 
                color='sv_tag', 
                text_auto='.1f', 
                template="plotly_dark",
                custom_data=['count'] # Pass count to hover data
            )

            # Fix: Explicitly map the sample count into the hover tooltip
            fig.update_traces(
                hovertemplate="<b>Station:</b> %{x}<br><b>Median CT:</b> %{y:.1f}s<br><b>Samples Used:</b> %{custom_data[0]}"
            )

            fig.add_hline(y=goal_time, line_color="green", annotation_text="Goal")
            fig.add_hline(y=bottleneck_buffered, line_dash="dash", line_color="orange")
            st.plotly_chart(fig, use_container_width=True)

            excel_file = convert_df_to_excel(df_final, summary)
            st.download_button(label="📥 Download Excel Report", data=excel_file, file_name="Report.xlsx")
        else:
            st.warning("No data matches selected filters.")

        # CLEAN UP RAM
        del df_filtered
        del df_final
        gc.collect() 

if __name__ == "__main__":
    main()
