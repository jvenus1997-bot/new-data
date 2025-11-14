import streamlit as st
import pandas as pd

st.title("Excel File Loader in Streamlit")

# Use st.file_uploader to allow users to upload an Excel file
uploaded_file = st.file_uploader("C:/venus/Open GRN.xlsx", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Read the Excel file into a pandas DataFrame
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.success("File successfully uploaded and read!")
        
        # Display the DataFrame in the Streamlit app
        st.write("### Data Preview")
        st.dataframe(df)
        
    except Exception as e:
        st.error(f"An error occurred: {e}")
        st.write("Please ensure the file is a valid Excel format (xlsx or xls) and not password-protected.")

else:
    st.info("Please upload an Excel file to proceed.")
