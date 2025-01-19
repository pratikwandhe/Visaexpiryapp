import streamlit as st
import pandas as pd
from datetime import datetime, timedelta


# Function to style rows based on visa expiry status
def highlight_rows(row, date_column):
    today = datetime.today()
    target_date = today + timedelta(days=30)

    if pd.isnull(row[date_column]):
        return [''] * len(row)

    expiry_date = row[date_column]
    if expiry_date < today:
        return ['background-color: red; color: white;'] * len(row)
    elif today <= expiry_date <= target_date:
        return ['background-color: orange; color: black;'] * len(row)
    else:
        return [''] * len(row)


# Function to send email (Stubbed for now)
def send_email(recipient_email, student_name, days_remaining):
    return f"Email sent to {student_name} ({recipient_email}) with {days_remaining} days remaining."


# Streamlit App
st.set_page_config(page_title="SPH Visa Alerts", page_icon="🛂", layout="wide")

# Header
st.markdown(
    """
    <style>
    body {
        background-color: #f0f8ff;
    }
    footer {visibility: hidden;}
    .small-button button {
        padding: 2px 5px;
        font-size: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.image("SPH_Rastar_image_landscape-1-removebg-preview.png", use_container_width=True)
st.markdown("<h1 style='text-align: center; color: blue;'>SPH's Visa Expiry Alerts</h1>", unsafe_allow_html=True)

# Sidebar
st.sidebar.title("Navigation")
st.sidebar.info("Use this app to identify visa expiry alerts for students.")
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read uploaded file
        df = pd.read_excel(uploaded_file)

        # Sidebar dropdowns for column selection
        date_column = st.sidebar.selectbox("Select the Visa Expiry Date Column", df.columns)
        email_column = st.sidebar.selectbox("Select the Email Column", df.columns)

        # Convert date column to datetime
        df[date_column] = pd.to_datetime(df[date_column], format="%d-%m-%Y", errors="coerce")

        # Calculate days remaining
        today = datetime.today()
        df["Days Remaining"] = (df[date_column] - today).dt.days

        # Split data into expiring soon and expired
        expiring_students = df[(df["Days Remaining"] > 0) & (df["Days Remaining"] <= 30)]
        expired_students = df[df["Days Remaining"] < 0]

        # Highlight rows based on visa expiry
        styled_df = df.style.apply(lambda row: highlight_rows(row, date_column), axis=1)

        # Metrics
        st.markdown("## 📊 Key Metrics")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Records", len(df))
        col2.metric("Expiring Soon", len(expiring_students))
        col3.metric("Already Expired", len(expired_students))

        # Display expiring students in a styled table with buttons
        st.markdown("## 🟠 Students with Visa Expiring Soon:")
        if not expiring_students.empty:
            for i, row in expiring_students.iterrows():
                student_name = row["Student Name"]
                days_remaining = int(row["Days Remaining"])
                recipient_email = row[email_column]

                col1, col2, col3 = st.columns([5, 2, 2])
                col1.write(f"**{student_name}** - {days_remaining} days remaining")
                if col2.button("Preview Email", key=f"preview_{i}"):
                    st.info(
                        f"""
                        **Preview Email:**
                        Subject: Visa Renewal Notification

                        Dear {student_name},

                        Your visa is expiring in {days_remaining} days. Please gather the necessary documents and be ready for the visa renewal process.

                        Thanks and regards,
                        Study Palace Hub Team
                        """
                    )
                if col3.button("Send Email", key=f"send_{i}"):
                    result = send_email(recipient_email, student_name, days_remaining)
                    st.success(result)

        else:
            st.info("No students have a visa expiring in the next 30 days.")

        # Display expired students in a styled table (no buttons)
        st.markdown("## 🔴 Students Whose Visa Has Already Expired:")
        if not expired_students.empty:
            expired_students_table = expired_students[["Student Name", date_column, "Days Remaining"]]
            st.dataframe(expired_students_table.style.apply(
                lambda row: ['background-color: red; color: white;'] * len(row), axis=1
            ))
        else:
            st.info("No students have expired visas.")

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")
