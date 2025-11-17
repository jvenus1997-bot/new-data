import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Page config
st.set_page_config(page_title="Procurement Spending Dashboard 2023-24", layout="wide")
st.title("📊 Procurement Spending Dashboard (FY 2023-24)")

# --- File Upload (for flexibility; assumes file is named as attached) ---
uploaded_file = st.file_uploader(r"C:\Users\QlikSense\Downloads\ME2M Spent April 2023-March 2024.xlsx", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Read Excel
        df = pd.read_excel(uploaded_file, sheet_name="Sheet1", skiprows=0)  # Headers in row 1 (0-indexed)
        
        # Clean data
        df.columns = df.columns.str.strip().str.replace('\n', ' ').str.upper()  # Normalize headers
        numeric_cols = ['PO QUANTITY', 'RATE INR', 'NET SPEND INR CR.']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        df = df.dropna(subset=['NET SPEND INR CR.'])  # Drop invalid spends
        
        st.success(f"Loaded data: {df.shape[0]} rows, {df.shape[1]} columns")
        
        # --- Sidebar Filters ---
        st.sidebar.header("Filters")
        selected_periods = st.sidebar.multiselect("Select Periods", options=sorted(df['PERIOD'].unique()), default=df['PERIOD'].unique())
        selected_categories = st.sidebar.multiselect("Select Categories", options=sorted(df['CATEGORY'].unique()), default=df['CATEGORY'].unique())
        selected_vendors = st.sidebar.multiselect("Select Vendors (Top 20 by Spend)", 
                                                  options=df.groupby('VENDOR')['NET SPEND INR CR.'].sum().nlargest(20).index)
        min_spend, max_spend = st.sidebar.slider("Net Spend Range (INR Cr.)", 
                                                 float(df['NET SPEND INR CR.'].min()), 
                                                 float(df['NET SPEND INR CR.'].max()), 
                                                 (0.0, float(df['NET SPEND INR CR.'].max())))
        
        # Apply filters
        df_filtered = df[
            (df['PERIOD'].isin(selected_periods)) &
            (df['CATEGORY'].isin(selected_categories)) &
            (df['NET SPEND INR CR.'].between(min_spend, max_spend))
        ]
        if selected_vendors:
            df_filtered = df_filtered[df_filtered['VENDOR'].isin(selected_vendors)]
        
        # --- Overview Metrics ---
        st.header("Overview")
        col1, col2, col3, col4 = st.columns(4)
        total_spend = df_filtered['NET SPEND INR CR.'].sum()
        num_pos = df_filtered['PO_NUMBER_SP'].nunique()
        num_vendors = df_filtered['VENDOR'].nunique()
        avg_spend = total_spend / num_pos if num_pos > 0 else 0
        
        col1.metric("Total Net Spend (INR Cr.)", f"{total_spend:,.2f}")
        col2.metric("Number of POs", f"{num_pos:,}")
        col3.metric("Number of Vendors", f"{num_vendors:,}")
        col4.metric("Avg Spend per PO (INR Cr.)", f"{avg_spend:,.2f}")
        
        # --- Charts ---
        st.header("Visualizations")
        
        # Spend by Period (Monthly Trend)
        spend_by_period = df_filtered.groupby('PERIOD')['NET SPEND INR CR.'].sum().reset_index()
        fig_period = px.line(spend_by_period, x='PERIOD', y='NET SPEND INR CR.', 
                             title="Spend Trend by Period", markers=True)
        st.plotly_chart(fig_period, use_container_width=True)
        
        # Spend by Category (Pie Chart)
        spend_by_category = df_filtered.groupby('CATEGORY')['NET SPEND INR CR.'].sum().reset_index()
        fig_category = px.pie(spend_by_category, values='NET SPEND INR CR.', names='CATEGORY', 
                              title="Spend Distribution by Category")
        st.plotly_chart(fig_category, use_container_width=True)
        
        # Top 10 Vendors by Spend (Bar Chart)
        top_vendors = df_filtered.groupby('VENDOR')['NET SPEND INR CR.'].sum().nlargest(10).reset_index()
        fig_vendors = px.bar(top_vendors, x='NET SPEND INR CR.', y='VENDOR', 
                             title="Top 10 Vendors by Spend", orientation='h')
        st.plotly_chart(fig_vendors, use_container_width=True)
        
        # Top 10 Materials by Spend (Bar Chart)
        top_materials = df_filtered.groupby('MAT.DESC')['NET SPEND INR CR.'].sum().nlargest(10).reset_index()
        fig_materials = px.bar(top_materials, x='NET SPEND INR CR.', y='MAT.DESC', 
                               title="Top 10 Materials by Spend", orientation='h')
        st.plotly_chart(fig_materials, use_container_width=True)
        
        # --- Data Table ---
        st.header("Detailed Data")
        st.dataframe(df_filtered.style.format({
            'PO QUANTITY': '{:,.2f}',
            'RATE INR': '{:,.2f}',
            'NET SPEND INR CR.': '{:,.2f}'
        }), use_container_width=True)
        
        # Download filtered data
        csv = df_filtered.to_csv(index=False).encode()
        st.download_button(
            label="📥 Download Filtered Data as CSV",
            data=csv,
            file_name="filtered_spending_data.csv",
            mime="text/csv"
        )
    
    except Exception as e:
        st.error(f"Error loading file: {e}")

else:
    st.info("👆 Upload the Excel file to start analyzing. Expected sheet: 'Sheet1' with procurement data.")
    st.markdown("""
    ### Dashboard Features:
    - **Overview Metrics**: Total spend, POs, vendors.
    - **Filters**: Period, category, vendor, spend range.
    - **Charts**: Trends, distributions, top vendors/materials.
    - **Interactive Table**: Sort, filter, download.
    """)
