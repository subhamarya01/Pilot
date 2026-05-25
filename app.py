import io
import re
from typing import List

import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# Page Setup
# ============================================================
st.set_page_config(
    page_title="AdoptionIQ Phase 1",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AdoptionIQ Phase 1")
st.caption("Excel Reader + KPI Engine + Customer Adoption Dashboard")


# ============================================================
# Helper Functions
# ============================================================
def clean_text(value):
    value = str(value).strip().lower()
    value = re.sub(r"[_\-]+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value


def find_column(df: pd.DataFrame, possible_names: List[str]):
    possible_clean = [clean_text(x) for x in possible_names]

    for col in df.columns:
        if clean_text(col) in possible_clean:
            return col

    return None


def safe_numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


def detect_file_type(columns):
    text = " ".join([clean_text(c) for c in columns])

    scores = {
        "License Data": 0,
        "Product Usage Data": 0,
        "AI Points Data": 0,
        "Training Data": 0,
        "Unknown": 0
    }

    if any(word in text for word in ["license", "licenses", "seat", "seats", "entitlement"]):
        scores["License Data"] += 4

    if any(word in text for word in ["active", "login", "usage", "feature", "session", "mau"]):
        scores["Product Usage Data"] += 4

    if any(word in text for word in ["point", "points", "credit", "credits", "token", "tokens"]):
        scores["AI Points Data"] += 5

    if any(word in text for word in ["training", "learner", "course", "attended", "completed", "nominated"]):
        scores["Training Data"] += 5

    best_type = max(scores, key=scores.get)

    if scores[best_type] == 0:
        return "Unknown"

    return best_type


def get_customer_columns(df):
    customer_id_col = find_column(
        df,
        [
            "Customer ID", "Customer Id", "Customer Number", "Customer No",
            "Account ID", "Account Id", "Account Number", "Client ID"
        ]
    )

    customer_name_col = find_column(
        df,
        [
            "Customer Name", "Customer", "Account Name", "Account",
            "Client Name", "Client", "Company Name", "Company"
        ]
    )

    return customer_id_col, customer_name_col


def create_customer_key(df):
    customer_id_col, customer_name_col = get_customer_columns(df)

    if customer_id_col:
        df["customer_key"] = df[customer_id_col].astype(str).str.strip()
    elif customer_name_col:
        df["customer_key"] = df[customer_name_col].astype(str).str.strip()
    else:
        df["customer_key"] = "UNKNOWN"

    if customer_name_col:
        df["customer_name_std"] = df[customer_name_col].astype(str).str.strip()
    else:
        df["customer_name_std"] = df["customer_key"]

    return df


def classify_status(score):
    if score >= 80:
        return "Healthy"
    elif score >= 60:
        return "Watch"
    elif score >= 40:
        return "Adoption Risk"
    else:
        return "High Risk"


def generate_insight(row):
    issues = []

    if row["licenses_sold"] > 0 and row["license_utilization_pct"] < 40:
        issues.append("low license utilization")

    if row["ai_points_allocated"] > 0 and row["ai_points_utilization_pct"] < 25:
        issues.append("low AI points consumption")

    if row["users_nominated"] > 0 and row["training_completion_pct"] < 40:
        issues.append("low training completion")

    if row["training_completion_pct"] >= 70 and row["license_utilization_pct"] < 40:
        issues.append("training-to-usage gap")

    if row["license_utilization_pct"] >= 80 and row["ai_points_utilization_pct"] >= 80:
        return "Strong adoption. Possible expansion / upsell opportunity."

    if issues:
        return "Attention needed due to " + ", ".join(issues) + "."

    return "Moderate adoption. Continue monitoring."


def aggregate_metric(df, metric_col, output_col):
    if df.empty or metric_col is None:
        return pd.DataFrame(columns=["customer_key", "customer_name_std", output_col])

    df = create_customer_key(df.copy())
    df[metric_col] = safe_numeric(df[metric_col])

    result = (
        df.groupby(["customer_key", "customer_name_std"], dropna=False)[metric_col]
        .sum()
        .reset_index()
        .rename(columns={metric_col: output_col})
    )

    return result


# ============================================================
# Upload Section
# ============================================================
st.sidebar.header("Upload Excel Files")

uploaded_files = st.sidebar.file_uploader(
    "Upload one or more Excel files",
    type=["xlsx", "xlsm", "xls"],
    accept_multiple_files=True
)

st.sidebar.info(
    "Upload License, Usage, AI Points and Training Excel files together."
)


if not uploaded_files:
    st.info("Please upload Excel files from the left sidebar.")
    st.stop()


# ============================================================
# Storage Lists
# ============================================================
file_log = []

license_frames = []
usage_frames = []
ai_points_frames = []
training_frames = []
unknown_frames = []


# ============================================================
# Read All Files and Sheets
# ============================================================
st.subheader("1. File Reading and Detection")

for uploaded_file in uploaded_files:
    try:
        excel_file = pd.ExcelFile(uploaded_file)

        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)

            df = df.dropna(how="all").dropna(axis=1, how="all")

            if df.empty:
                file_log.append({
                    "File Name": uploaded_file.name,
                    "Sheet Name": sheet_name,
                    "Detected Type": "Empty Sheet",
                    "Rows": 0,
                    "Columns": 0,
                    "Status": "Skipped"
                })
                continue

            detected_type = detect_file_type(df.columns)

            file_log.append({
                "File Name": uploaded_file.name,
                "Sheet Name": sheet_name,
                "Detected Type": detected_type,
                "Rows": len(df),
                "Columns": len(df.columns),
                "Status": "Success"
            })

            if detected_type == "License Data":
                license_frames.append(df)
            elif detected_type == "Product Usage Data":
                usage_frames.append(df)
            elif detected_type == "AI Points Data":
                ai_points_frames.append(df)
            elif detected_type == "Training Data":
                training_frames.append(df)
            else:
                unknown_frames.append(df)

    except Exception as e:
        file_log.append({
            "File Name": uploaded_file.name,
            "Sheet Name": "-",
            "Detected Type": "Error",
            "Rows": 0,
            "Columns": 0,
            "Status": str(e)
        })


file_log_df = pd.DataFrame(file_log)
st.dataframe(file_log_df, use_container_width=True)


# ============================================================
# Combine Frames
# ============================================================
license_df = pd.concat(license_frames, ignore_index=True, sort=False) if license_frames else pd.DataFrame()
usage_df = pd.concat(usage_frames, ignore_index=True, sort=False) if usage_frames else pd.DataFrame()
ai_points_df = pd.concat(ai_points_frames, ignore_index=True, sort=False) if ai_points_frames else pd.DataFrame()
training_df = pd.concat(training_frames, ignore_index=True, sort=False) if training_frames else pd.DataFrame()


with st.expander("Preview Detected Raw Data"):
    tab1, tab2, tab3, tab4 = st.tabs(
        ["License Data", "Usage Data", "AI Points Data", "Training Data"]
    )

    with tab1:
        st.dataframe(license_df.head(30), use_container_width=True)
    with tab2:
        st.dataframe(usage_df.head(30), use_container_width=True)
    with tab3:
        st.dataframe(ai_points_df.head(30), use_container_width=True)
    with tab4:
        st.dataframe(training_df.head(30), use_container_width=True)


# ============================================================
# Find Important Columns
# ============================================================
license_col = find_column(
    license_df,
    [
        "Licenses Sold", "License Sold", "Seats Sold", "Seat Sold",
        "Purchased Licenses", "Purchased Seats", "Entitled Users",
        "Entitlement", "License Qty", "License Quantity", "Sold Licenses"
    ]
) if not license_df.empty else None

active_users_col = find_column(
    usage_df,
    [
        "Active Users", "Active User", "Active Seats", "Activated Users",
        "Monthly Active Users", "MAU", "Used Licenses"
    ]
) if not usage_df.empty else None

login_count_col = find_column(
    usage_df,
    [
        "Login Count", "Logins", "Number of Logins", "Sessions", "Visits"
    ]
) if not usage_df.empty else None

feature_usage_col = find_column(
    usage_df,
    [
        "Feature Usage Count", "Feature Usage", "Usage Count",
        "Transactions", "Activities", "Events"
    ]
) if not usage_df.empty else None

ai_allocated_col = find_column(
    ai_points_df,
    [
        "AI Points Allocated", "Points Allocated", "Credits Allocated",
        "Tokens Allocated", "AI Credit Assigned", "AI Credits Assigned",
        "Allocated Credits", "Entitlement Points"
    ]
) if not ai_points_df.empty else None

ai_consumed_col = find_column(
    ai_points_df,
    [
        "AI Points Consumed", "Points Consumed", "Credits Used",
        "Tokens Consumed", "AI Credit Used", "AI Credits Used",
        "Points Used", "Consumed Credits"
    ]
) if not ai_points_df.empty else None

users_nominated_col = find_column(
    training_df,
    [
        "Users Nominated", "Nominated Users", "Registered Users",
        "Enrolled Users", "Assigned Learners"
    ]
) if not training_df.empty else None

users_attended_col = find_column(
    training_df,
    [
        "Users Attended", "Attended Users", "Participants",
        "Attendance", "Learners Attended"
    ]
) if not training_df.empty else None

users_completed_col = find_column(
    training_df,
    [
        "Users Completed", "Completed Users", "Learners Completed",
        "Training Completed", "Completion Count", "Completed Learners"
    ]
) if not training_df.empty else None


# ============================================================
# Column Detection Summary
# ============================================================
st.subheader("2. Column Detection Summary")

column_detection = pd.DataFrame([
    {"Area": "License", "Required Metric": "Licenses Sold", "Detected Column": license_col},
    {"Area": "Usage", "Required Metric": "Active Users", "Detected Column": active_users_col},
    {"Area": "Usage", "Required Metric": "Login Count", "Detected Column": login_count_col},
    {"Area": "Usage", "Required Metric": "Feature Usage Count", "Detected Column": feature_usage_col},
    {"Area": "AI Points", "Required Metric": "AI Points Allocated", "Detected Column": ai_allocated_col},
    {"Area": "AI Points", "Required Metric": "AI Points Consumed", "Detected Column": ai_consumed_col},
    {"Area": "Training", "Required Metric": "Users Nominated", "Detected Column": users_nominated_col},
    {"Area": "Training", "Required Metric": "Users Attended", "Detected Column": users_attended_col},
    {"Area": "Training", "Required Metric": "Users Completed", "Detected Column": users_completed_col},
])

st.dataframe(column_detection, use_container_width=True)


# ============================================================
# Aggregate Customer KPIs
# ============================================================
license_agg = aggregate_metric(license_df, license_col, "licenses_sold")
usage_active_agg = aggregate_metric(usage_df, active_users_col, "active_users")
usage_login_agg = aggregate_metric(usage_df, login_count_col, "login_count")
usage_feature_agg = aggregate_metric(usage_df, feature_usage_col, "feature_usage_count")
ai_allocated_agg = aggregate_metric(ai_points_df, ai_allocated_col, "ai_points_allocated")
ai_consumed_agg = aggregate_metric(ai_points_df, ai_consumed_col, "ai_points_consumed")
training_nom_agg = aggregate_metric(training_df, users_nominated_col, "users_nominated")
training_att_agg = aggregate_metric(training_df, users_attended_col, "users_attended")
training_comp_agg = aggregate_metric(training_df, users_completed_col, "users_completed")


summary_parts = [
    license_agg,
    usage_active_agg,
    usage_login_agg,
    usage_feature_agg,
    ai_allocated_agg,
    ai_consumed_agg,
    training_nom_agg,
    training_att_agg,
    training_comp_agg,
]

customer_summary = pd.DataFrame()

for part in summary_parts:
    if part.empty:
        continue

    if customer_summary.empty:
        customer_summary = part.copy()
    else:
        customer_summary = customer_summary.merge(
            part,
            on=["customer_key", "customer_name_std"],
            how="outer"
        )


if customer_summary.empty:
    st.error("No customer-level summary could be created. Please check customer columns in uploaded Excel files.")
    st.stop()


# Fill missing numeric columns
required_numeric_cols = [
    "licenses_sold",
    "active_users",
    "login_count",
    "feature_usage_count",
    "ai_points_allocated",
    "ai_points_consumed",
    "users_nominated",
    "users_attended",
    "users_completed"
]

for col in required_numeric_cols:
    if col not in customer_summary.columns:
        customer_summary[col] = 0
    customer_summary[col] = safe_numeric(customer_summary[col])


# ============================================================
# Calculate KPIs
# ============================================================
customer_summary["license_utilization_pct"] = customer_summary.apply(
    lambda r: round((r["active_users"] / r["licenses_sold"]) * 100, 1)
    if r["licenses_sold"] > 0 else 0,
    axis=1
)

customer_summary["ai_points_utilization_pct"] = customer_summary.apply(
    lambda r: round((r["ai_points_consumed"] / r["ai_points_allocated"]) * 100, 1)
    if r["ai_points_allocated"] > 0 else 0,
    axis=1
)

customer_summary["training_completion_pct"] = customer_summary.apply(
    lambda r: round((r["users_completed"] / r["users_nominated"]) * 100, 1)
    if r["users_nominated"] > 0 else 0,
    axis=1
)

customer_summary["usage_activity_score"] = customer_summary.apply(
    lambda r: 100 if r["login_count"] > 0 or r["feature_usage_count"] > 0 else 0,
    axis=1
)

customer_summary["adoption_score"] = (
    customer_summary["license_utilization_pct"] * 0.35
    + customer_summary["ai_points_utilization_pct"] * 0.35
    + customer_summary["training_completion_pct"] * 0.20
    + customer_summary["usage_activity_score"] * 0.10
).round(1)

customer_summary["adoption_score"] = customer_summary["adoption_score"].clip(upper=100)

customer_summary["status"] = customer_summary["adoption_score"].apply(classify_status)
customer_summary["insight"] = customer_summary.apply(generate_insight, axis=1)


# ============================================================
# Dashboard KPIs
# ============================================================
st.subheader("3. Management Dashboard")

total_customers = len(customer_summary)
total_licenses = customer_summary["licenses_sold"].sum()
total_active_users = customer_summary["active_users"].sum()
license_util_total = round((total_active_users / total_licenses) * 100, 1) if total_licenses > 0 else 0

total_ai_allocated = customer_summary["ai_points_allocated"].sum()
total_ai_consumed = customer_summary["ai_points_consumed"].sum()
ai_util_total = round((total_ai_consumed / total_ai_allocated) * 100, 1) if total_ai_allocated > 0 else 0

total_users_completed = customer_summary["users_completed"].sum()

risk_customers = customer_summary[
    customer_summary["status"].isin(["High Risk", "Adoption Risk"])
].shape[0]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Customers", f"{total_customers:,.0f}")
c2.metric("Total Licenses Sold", f"{total_licenses:,.0f}")
c3.metric("Total Active Users", f"{total_active_users:,.0f}")
c4.metric("License Utilization", f"{license_util_total}%")

c5, c6, c7, c8 = st.columns(4)
c5.metric("AI Points Allocated", f"{total_ai_allocated:,.0f}")
c6.metric("AI Points Consumed", f"{total_ai_consumed:,.0f}")
c7.metric("AI Points Utilization", f"{ai_util_total}%")
c8.metric("Risk Customers", f"{risk_customers:,.0f}")


# ============================================================
# Charts
# ============================================================
st.subheader("4. Charts")

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    status_count = customer_summary["status"].value_counts().reset_index()
    status_count.columns = ["Status", "Count"]

    fig = px.bar(
        status_count,
        x="Status",
        y="Count",
        title="Customer Health Distribution"
    )
    st.plotly_chart(fig, use_container_width=True)

with chart_col2:
    top_license = customer_summary.sort_values("licenses_sold", ascending=False).head(10)

    fig = px.bar(
        top_license,
        x="customer_name_std",
        y="licenses_sold",
        title="Top 10 Customers by Licenses Sold"
    )
    st.plotly_chart(fig, use_container_width=True)


chart_col3, chart_col4 = st.columns(2)

with chart_col3:
    low_usage = customer_summary[
        customer_summary["licenses_sold"] > 0
    ].sort_values("license_utilization_pct", ascending=True).head(10)

    fig = px.bar(
        low_usage,
        x="customer_name_std",
        y="license_utilization_pct",
        title="Lowest 10 Customers by License Utilization %"
    )
    st.plotly_chart(fig, use_container_width=True)

with chart_col4:
    ai_chart = customer_summary.sort_values("ai_points_allocated", ascending=False).head(10)

    fig = px.bar(
        ai_chart,
        x="customer_name_std",
        y=["ai_points_allocated", "ai_points_consumed"],
        title="AI Points Allocated vs Consumed"
    )
    st.plotly_chart(fig, use_container_width=True)


training_chart = customer_summary.sort_values("users_completed", ascending=False).head(15)

fig = px.scatter(
    training_chart,
    x="users_completed",
    y="active_users",
    size="licenses_sold",
    hover_name="customer_name_std",
    title="Training Completed vs Active Users"
)
st.plotly_chart(fig, use_container_width=True)


# ============================================================
# Customer Summary Table
# ============================================================
st.subheader("5. Customer Adoption Summary")

display_cols = [
    "customer_key",
    "customer_name_std",
    "licenses_sold",
    "active_users",
    "license_utilization_pct",
    "ai_points_allocated",
    "ai_points_consumed",
    "ai_points_utilization_pct",
    "users_nominated",
    "users_attended",
    "users_completed",
    "training_completion_pct",
    "adoption_score",
    "status",
    "insight"
]

customer_summary_display = customer_summary[display_cols].sort_values(
    "adoption_score",
    ascending=True
)

st.dataframe(customer_summary_display, use_container_width=True)


# ============================================================
# Rule-Based Executive Summary
# ============================================================
st.subheader("6. Rule-Based Executive Summary")

healthy = customer_summary[customer_summary["status"] == "Healthy"].shape[0]
watch = customer_summary[customer_summary["status"] == "Watch"].shape[0]
adoption_risk = customer_summary[customer_summary["status"] == "Adoption Risk"].shape[0]
high_risk = customer_summary[customer_summary["status"] == "High Risk"].shape[0]

summary_text = f"""
Overall, {total_customers} customers were analyzed.

Healthy customers: {healthy}  
Watch customers: {watch}  
Adoption Risk customers: {adoption_risk}  
High Risk customers: {high_risk}  

Overall license utilization is {license_util_total}%.  
Overall AI points utilization is {ai_util_total}%.  
Total users who completed training: {total_users_completed:,.0f}.  

Priority should be given to customers with high licenses sold but low active usage, and customers with low AI points consumption.
"""

st.info(summary_text)


# ============================================================
# Download Output
# ============================================================
st.subheader("7. Download Final Report Excel")

output = io.BytesIO()

with pd.ExcelWriter(output, engine="openpyxl") as writer:
    customer_summary_display.to_excel(writer, index=False, sheet_name="Customer Summary")
    file_log_df.to_excel(writer, index=False, sheet_name="File Processing Log")
    column_detection.to_excel(writer, index=False, sheet_name="Column Detection")

processed_file = output.getvalue()

st.download_button(
    label="Download Customer Adoption Summary",
    data=processed_file,
    file_name="customer_adoption_summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
