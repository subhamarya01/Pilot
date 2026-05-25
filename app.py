import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="AdoptionIQ Phase 1",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AdoptionIQ Phase 1")
st.success("App is running. Now testing Excel upload.")

st.subheader("Upload Excel File")

uploaded_file = st.file_uploader(
    "Upload one Excel file",
    type=["xlsx", "xlsm", "xls"]
)

if uploaded_file is not None:
    st.write("File uploaded:", uploaded_file.name)

    try:
        excel_file = pd.ExcelFile(uploaded_file)

        st.success("Excel file read successfully!")

        st.write("### Sheets found")
        st.write(excel_file.sheet_names)

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
else:
    st.info("Please upload one Excel file.")
