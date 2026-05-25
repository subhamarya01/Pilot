import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="AdoptionIQ Phase 1",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AdoptionIQ Phase 1")
st.success("App is running. Excel reading + file detection + first KPI working.")


# ------------------------------------------------------------
# Helper: Detect file type from column names
# ------------------------------------------------------------
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


# ------------------------------------------------------------
# Helper: Find column by possible names
# ------------------------------------------------------------
def find_column(df, possible_names):
    for col in df.columns:
        col_clean = str(col).strip().lower()

        for name in possible_names:
            if col_clean == name.strip().lower():
                return col

    return None


# ------------------------------------------------------------
# Upload section
# ------------------------------------------------------------
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

        # Remove fully empty rows and columns
        df = df.dropna(how="all").dropna(axis=1, how="all")

        detected_type = detect_file_type(df.columns)

        # ------------------------------------------------------------
        # Detected file type
        # ------------------------------------------------------------
        st.subheader("Detected File Type")
        st.info(detected_type)

        # ------------------------------------------------------------
        # File summary
        # ------------------------------------------------------------
        st.subheader("File Summary")

        col1, col2 = st.columns(2)
        col1.metric("Rows", len(df))
        col2.metric("Columns", len(df.columns))

        # ------------------------------------------------------------
        # Columns found
        # ------------------------------------------------------------
        st.subheader("Columns Found")
        st.write(list(df.columns))

        # ------------------------------------------------------------
        # License KPI
        # ------------------------------------------------------------
        if detected_type == "License Data":
            st.subheader("License KPI")

            license_col = find_column(
                df,
                [
                    "Licenses Sold",
                    "License Sold",
                    "Seats Sold",
                    "Seat Sold",
                    "Purchased Licenses",
                    "Purchased Seats",
                    "Entitled Users",
                    "Entitlement",
                    "License Qty",
                    "License Quantity",
                    "Sold Licenses"
                ]
            )

            if license_col:
                df[license_col] = pd.to_numeric(df[license_col], errors="coerce").fillna(0)
                total_licenses = df[license_col].sum()

                st.metric("Total Licenses Sold", f"{total_licenses:,.0f}")
                st.caption(f"Calculated from column: {license_col}")
            else:
                st.warning(
                    "License column not found. Please use a column like "
                    "'Licenses Sold', 'Seats Sold', or 'License Quantity'."
                )

        # ------------------------------------------------------------
        # Usage KPI
        # ------------------------------------------------------------
        if detected_type == "Product Usage Data":
            st.subheader("Usage KPI")

            active_user_col = find_column(
                df,
                [
                    "Active Users",
                    "Active User",
                    "Active Seats",
                    "Activated Users",
                    "Monthly Active Users",
                    "MAU",
                    "Used Licenses"
                ]
            )

            if active_user_col:
                df[active_user_col] = pd.to_numeric(df[active_user_col], errors="coerce").fillna(0)
                total_active_users = df[active_user_col].sum()

                st.metric("Total Active Users", f"{total_active_users:,.0f}")
                st.caption(f"Calculated from column: {active_user_col}")
            else:
                st.warning(
                    "Active user column not found. Please use a column like "
                    "'Active Users', 'Active Seats', or 'Monthly Active Users'."
                )

        # ------------------------------------------------------------
        # AI Points KPI
        # ------------------------------------------------------------
        if detected_type == "AI Points Data":
            st.subheader("AI Points KPI")

            allocated_col = find_column(
                df,
                [
                    "AI Points Allocated",
                    "Points Allocated",
                    "Credits Allocated",
                    "Tokens Allocated",
                    "AI Credit Assigned",
                    "AI Credits Assigned",
                    "Allocated Credits",
                    "Entitlement Points"
                ]
            )

            consumed_col = find_column(
                df,
                [
                    "AI Points Consumed",
                    "Points Consumed",
                    "Credits Used",
                    "Tokens Consumed",
                    "AI Credit Used",
                    "AI Credits Used",
                    "Points Used",
                    "Consumed Credits"
                ]
            )

            if allocated_col:
                df[allocated_col] = pd.to_numeric(df[allocated_col], errors="coerce").fillna(0)
                total_allocated = df[allocated_col].sum()
                st.metric("AI Points Allocated", f"{total_allocated:,.0f}")
                st.caption(f"Allocated calculated from column: {allocated_col}")
            else:
                total_allocated = 0
                st.warning("AI points allocated column not found.")

            if consumed_col:
                df[consumed_col] = pd.to_numeric(df[consumed_col], errors="coerce").fillna(0)
                total_consumed = df[consumed_col].sum()
                st.metric("AI Points Consumed", f"{total_consumed:,.0f}")
                st.caption(f"Consumed calculated from column: {consumed_col}")
            else:
                total_consumed = 0
                st.warning("AI points consumed column not found.")

            if allocated_col and consumed_col and total_allocated > 0:
                utilization = round((total_consumed / total_allocated) * 100, 1)
                st.metric("AI Points Utilization %", f"{utilization}%")

        # ------------------------------------------------------------
        # Training KPI
        # ------------------------------------------------------------
        if detected_type == "Training Data":
            st.subheader("Training KPI")

            nominated_col = find_column(
                df,
                [
                    "Users Nominated",
                    "Nominated Users",
                    "Registered Users",
                    "Enrolled Users",
                    "Assigned Learners"
                ]
            )

            completed_col = find_column(
                df,
                [
                    "Users Completed",
                    "Completed Users",
                    "Learners Completed",
                    "Training Completed",
                    "Completion Count",
                    "Completed Learners"
                ]
            )

            if nominated_col:
                df[nominated_col] = pd.to_numeric(df[nominated_col], errors="coerce").fillna(0)
                total_nominated = df[nominated_col].sum()
                st.metric("Users Nominated", f"{total_nominated:,.0f}")
                st.caption(f"Nominated calculated from column: {nominated_col}")
            else:
                total_nominated = 0
                st.warning("Users nominated column not found.")

            if completed_col:
                df[completed_col] = pd.to_numeric(df[completed_col], errors="coerce").fillna(0)
                total_completed = df[completed_col].sum()
                st.metric("Users Completed", f"{total_completed:,.0f}")
                st.caption(f"Completed calculated from column: {completed_col}")
            else:
                total_completed = 0
                st.warning("Users completed column not found.")

            if nominated_col and completed_col and total_nominated > 0:
                completion_pct = round((total_completed / total_nominated) * 100, 1)
                st.metric("Training Completion %", f"{completion_pct}%")

        # ------------------------------------------------------------
        # Data Preview
        # ------------------------------------------------------------
        st.subheader("Data Preview")
        st.dataframe(df.head(20), use_container_width=True)

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")

else:
    st.info("Please upload one Excel file.")
