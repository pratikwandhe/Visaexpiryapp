import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText
import time


# Function to highlight rows based on visa expiry status
def style_rows(row, date_column):
    today = datetime.today()
    target_date = today + timedelta(days=30)

    if pd.isnull(row[date_column]):
        return ''
    expiry_date = row[date_column]
    if expiry_date < today:
        return 'background-color: red; color: white;'
    elif today <= expiry_date <= target_date:
        return 'background-color: orange; color: black;'
    else:
        return ''


# Function to send email
def send_email(recipient_email, student_name, days_remaining):
    sender_email = "your_email@gmail.com"  # Replace with your Gmail address
    sender_password = "your_app_password"  # Replace with your Gmail App Password

    subject = "Visa Renewal Notification"
    body = f"""Dear {student_name},

Your visa is expiring in {days_remaining} days. Please gather the necessary documents and be ready for the visa renewal process.

Thanks and regards,
Study Palace Hub Team
"""

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = recipient_email

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()  # Secure the connection
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, msg.as_string())
        return True
    except Exception as e:
        st.error(f"Error sending email to {recipient_email}: {e}")
        return False


# Streamlit App
st.set_page_config(page_title="SPH Visa Alerts", page_icon="🛂", layout="wide")

# Header with improved styling
st.markdown(
    """
    <div style="background-color: #1E90FF; padding: 10px; border-radius: 10px; text-align: center;">
        <h1 style="color: white;">SPH's Visa Expiry Alerts</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("### 🚨 Quickly check visa expiration dates and notify students!")

# Sidebar
st.sidebar.title("Navigation")
st.sidebar.info("Use this app to identify visa expiry alerts for students.")
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read uploaded file
        df = pd.read_excel(uploaded_file)

        # Select columns for date and email
        date_column = st.sidebar.selectbox("Select the Visa Expiry Date Column", df.columns)
        email_column = st.sidebar.selectbox("Select the Email Column", df.columns)

        # Convert date column to datetime
        df[date_column] = pd.to_datetime(df[date_column], format="%d-%m-%Y", errors="coerce")

        # Calculate days remaining
        today = datetime.today()
        df["Days Remaining"] = (df[date_column] - today).dt.days

        # Sort rows by days remaining
        df = df.sort_values(by="Days Remaining")

        # Display the table with email preview and send buttons
        st.markdown("### 🟠 Visa Expiry Details")
        for index, row in df.iterrows():
            row_style = style_rows(row, date_column)
            student_name = row["Student Name"]
            expiry_date = row[date_column]
            days_remaining = row["Days Remaining"]
            recipient_email = row[email_column]

            col1, col2, col3, col4 = st.columns([2, 2, 2, 2])
            col1.markdown(f"<div style='{row_style}'>{student_name}</div>", unsafe_allow_html=True)
            col2.markdown(f"<div style='{row_style}'>{expiry_date.strftime('%d-%m-%Y') if not pd.isnull(expiry_date) else 'N/A'}</div>", unsafe_allow_html=True)
            col3.markdown(f"<div style='{row_style}'>{days_remaining if not pd.isnull(days_remaining) else 'N/A'} days</div>", unsafe_allow_html=True)

            if col4.button("Send Email", key=f"send_{index}"):
                if not pd.isnull(days_remaining) and days_remaining > 0:
                    success = send_email(recipient_email, student_name, days_remaining)
                    if success:
                        st.success(f"Email sent to {student_name} ({recipient_email}).")
                else:
                    st.warning(f"Cannot send email to {student_name} ({recipient_email}). Invalid or expired visa.")

            if col4.button("Preview Email", key=f"preview_{index}"):
                st.info(
                    f"""
                    **Preview Email for {student_name}:**
                    Subject: Visa Renewal Notification
                    
                    Dear {student_name},

                    Your visa is expiring in {days_remaining} days. Please gather the necessary documents and be ready for the visa renewal process.

                    Thanks and regards,
                    Study Palace Hub Team
                    """
                )

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
else:
    st.info("Please upload an Excel file to proceed.")
