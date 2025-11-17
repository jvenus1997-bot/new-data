import streamlit as st
import pandas as pd
import os

# Page config
st.set_page_config(page_title="Excel Data Viewer", layout="wide")
st.title("📊 Excel Data Viewer with Streamlit")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read Excel file
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        # Sidebar: Select sheet
        selected_sheet = st.sidebar.selectbox("Select Sheet", sheet_names)
        
        # Read selected sheet
        df = xl.parse(selected_sheet)
        
        st.success(f"Loaded sheet: **{selected_sheet}** | Rows: {df.shape[0]}, Columns: {df.shape[1]}")
        
        # --- Display Raw Data ---
        with st.expander("View Raw Data", expanded=False):
            st.dataframe(df, use_container_width=True)
        
        # --- Search & Filter ---
        st.subheader("🔍 Search & Filter")
        col1, col2 = st.columns(2)
        
        with col1:
            search_col = st.selectbox("Search in Column", options=["None"] + list(df.columns))
            if search_col != "None":
                search_term = st.text_input(f"Search in {search_col}")
                if search_term:
                    df_filtered = df[df[search_col].astype(str).str.contains(search_term, case=False, na=False)]
                else:
                    df_filtered = df.copy()
            else:
                df_filtered = df.copy()
        
        with col2:
            # Optional: Filter by column value
            filter_col = st.selectbox("Filter by Column", options=["None"] + list(df.columns))
            if filter_col != "None":
                unique_vals = df[filter_col].dropna().unique()
                selected_val = st.multiselect(f"Select {filter_col} values", options=unique_vals)
                if selected_val:
                    df_filtered = df_filtered[df_filtered[filter_col].isin(selected_val)]
        
        # Show filtered data
        if 'df_filtered' in locals():
            st.dataframe(df_filtered, use_container_width=True)
            
            # --- Download filtered data ---
            csv = df_filtered.to_csv(index=False).encode()
            st.download_button(
                label="📥 Download Filtered Data as CSV",
                data=csv,
                file_name=f"filtered_{selected_sheet}.csv",
                mime="text/csv"
            )
        
        # --- Summary Stats ---
        st.subheader("📈 Summary Statistics")
        st.write(df.describe(include='all'))
        
        # --- Column Info ---
        with st.expander("Column Info"):
            st.write(df.dtypes.reset_index().rename(columns={"index": "Column", 0: "Data Type"}))
    
    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.info("👆 Please upload an Excel (.xlsx) file to get started.")
    st.markdown("""
    ### Sample Use:
    1. Upload an Excel file
    2. Select a sheet
    3. Search, filter, and explore data
    4. Download results
    """)
