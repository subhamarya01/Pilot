import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="AdoptionIQ Phase 1",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AdoptionIQ Phase 1")
st.write("Excel Reader + KPI Engine + Dashboard")

st.subheader("Step 1: Upload Excel File")

uploaded_file = st.file_uploader(
    "Upload one Excel file",
    type=["xlsx", "xlsm", "xls"]
)

if uploaded_file is not None:
    try:
        excel_file = pd.ExcelFile(uploaded_file)

        st.success("Excel file read successfully!")

        st.write("### Sheets found")
        st.write(excel_file.sheet_names)

        sheet_name = st.selectbox(
            "Select sheet to preview",
            excel_file.sheet_names
        )

        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        st.write("### File Summary")
        st.write("Rows:", len(df))
        st.write("Columns:", len(df.columns))

        st.write("### Columns found")
        st.write(list(df.columns))

        st.write("### Data Preview")
        st.dataframe(df.head(20), use_container_width=True)

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
else:
    st.info("Please upload one Excel file to test.")
