import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import time


# Function to highlight expiry dates
def highlight_expiry_dates(row, date_column):
    today = datetime.today()
    target_date = today + timedelta(days=30)

    if pd.isnull(row[date_column]):
        return [''] * len(row)

    expiry_date = row[date_column]

    if expiry_date < today:
        return ['background-color: red'] * len(row)
    elif today <= expiry_date <= target_date:
        return ['background-color: orange'] * len(row)
    else:
        return [''] * len(row)


# Function to filter students with visa expiry in around 1 month
def get_students_with_visa_expiry_soon(data, date_column, days=30):
    try:
        # Calculate today's date and the target date range
        today = datetime.today()
        target_date = today + timedelta(days=days)

        # Filter students whose visa expiry date is within the next `days`
        expiring_soon = data[(data[date_column] >= today)
                             & (data[date_column] <= target_date)].copy()

        # Calculate days remaining for each student
        expiring_soon["Days Remaining"] = (expiring_soon[date_column] - today).dt.days

        # Sort by days remaining
        expiring_soon = expiring_soon.sort_values(by="Days Remaining", ascending=True)

        return expiring_soon
    except Exception as e:
        st.error(f"Error processing the data: {e}")
        return pd.DataFrame()


# Function to filter students whose visa has already expired
def get_students_with_expired_visa(data, date_column):
    try:
        today = datetime.today()

        # Filter students whose visa expiry date is in the past
        expired = data[data[date_column] < today].copy()

        # Add a "Days Overdue" column
        expired["Days Overdue"] = (today - expired[date_column]).dt.days

        # Sort by overdue days in descending order
        expired = expired.sort_values(by="Days Overdue", ascending=False)

        return expired
    except Exception as e:
        st.error(f"Error processing the data: {e}")
        return pd.DataFrame()


# Streamlit App
# Add CSS styling for the background
st.markdown(
    """
    <style>
    body {
        background-color: #f0f8ff;
    }
    footer {visibility: hidden;}
    </style>
    """,
    unsafe_allow_html=True,
)

# Add app header image
st.image("SPH_Rastar_image_landscape-1-removebg-preview.png", use_container_width=True)  # Replace with your image file name

# Add title and subheading
st.markdown("<h1 style='text-align: center; color: blue;'>SPH's Visa Expiry Alerts</h1>", unsafe_allow_html=True)
st.markdown("### 🚨 Quickly check visa expiration dates and act on them!")

# Sidebar content
st.sidebar.title("Navigation")
st.sidebar.info("Use this app to identify visa expiry alerts for students.")
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

# Main logic
if uploaded_file:
    try:
        # Read and preview uploaded file
        df = pd.read_excel(uploaded_file)
        st.write("Uploaded File Preview:")
        st.dataframe(df, height=300)

        # User input for the date column
        date_column = st.sidebar.selectbox("Select the Visa Expiry Date Column", df.columns)

        # Filter dates using sidebar
        st.sidebar.markdown("### Filter by Expiry Date")
        start_date = st.sidebar.date_input("Start Date", datetime.today())
        end_date = st.sidebar.date_input("End Date", datetime.today() + timedelta(days=30))

        # Process data
        with st.spinner("Processing your data..."):
            time.sleep(1)  # Simulate processing time
            df[date_column] = pd.to_datetime(df[date_column], format="%d-%m-%Y", errors="coerce")

        # Get students with visa expiring soon
        expiring_students = get_students_with_visa_expiry_soon(df, date_column)

        # Get students with expired visa
        expired_students = get_students_with_expired_visa(df, date_column)

        # Filtered data within date range
        filtered_df = df[(df[date_column] >= pd.Timestamp(start_date)) & (df[date_column] <= pd.Timestamp(end_date))]

        # Display metrics
        st.markdown("## 📊 Key Metrics")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Records", len(df))
        col2.metric("Expiring Soon", len(expiring_students))
        col3.metric("Already Expired", len(expired_students))

        # Display filtered results
        st.markdown("## 📅 Filtered Results by Date Range")
        st.dataframe(filtered_df)

        # Display expiring soon students
        if not expiring_students.empty:
            st.markdown("## 🟠 Students with Visa Expiring Soon (Sorted by Days Remaining):")
            st.dataframe(expiring_students)
        else:
            st.info("No students have a visa expiring in the next month.")

        # Display expired students
        if not expired_students.empty:
            st.markdown("## 🔴 Students Whose Visa Has Already Expired:")
            st.dataframe(expired_students)
        else:
            st.info("No students have expired visas.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")

# Footer
st.markdown("<hr><p style='text-align: center;'>© 2025 SPH Team</p>", unsafe_allow_html=True)
