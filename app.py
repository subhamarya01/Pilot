import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="AdoptionIQ Phase 1",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AdoptionIQ Phase 1")
st.success("App is running. Excel reading is working.")


def detect_file_type(columns):
    text = " ".join([str(c).lower() for c in columns])

    if any(word in text for word in ["license", "licenses", "seat", "seats", "entitlement"]):
        return "License Data"

    if any(word in text for word in ["active", "login", "usage", "feature", "session"]):
        return "Product Usage Data"

    if any(word in text for word in ["point", "points", "credit", "credits", "token", "tokens"]):
        return "AI Points Data"

    if any(word in text for word in ["training", "learner", "course", "attended", "completed", "nominated"]):
        return "Training Data"

    return "Unknown"


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

        sheet_name = st.selectbox(
            "Select sheet",
            excel_file.sheet_names
        )

        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

        detected_type = detect_file_type(df.columns)

        st.subheader("Detected File Type")
        st.info(detected_type)

        st.subheader("File Summary")

        col1, col2 = st.columns(2)
        col1.metric("Rows", len(df))
        col2.metric("Columns", len(df.columns))

        st.subheader("Columns Found")
        st.write(list(df.columns))

        st.subheader("Data Preview")
        st.dataframe(df.head(20), use_container_width=True)

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
else:
    st.info("Please upload one Excel file.")
