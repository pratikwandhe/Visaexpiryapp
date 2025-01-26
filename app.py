import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText


# Function to send email
def send_email(recipient_email, student_name, days_remaining, context):
    sender_email = "pratikwandhe9095@gmail.com"  # Replace with your Gmail address
    sender_password = "fixx dnwn jpin bwix"  # Replace with your Gmail App Password

    subject = f"{context} Renewal Notification"
    body = f"""Dear {student_name},

Your {context.lower()} is expiring in {days_remaining} days. Please gather the necessary documents and be ready for the renewal process.

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
        return str(e)  # Return error message


# Functions for processing visa and registration expiry
def get_students_with_expiry_soon(data, date_column, days=30):
    try:
        today = datetime.today()
        target_date = today + timedelta(days=days)
        expiring_soon = data[(data[date_column] >= today) & (data[date_column] <= target_date)].copy()
        expiring_soon["Days Remaining"] = (expiring_soon[date_column] - today).dt.days
        expiring_soon = expiring_soon.sort_values(by="Days Remaining", ascending=True)
        return expiring_soon
    except Exception as e:
        st.error(f"Error processing the data: {e}")
        return pd.DataFrame()


def get_students_with_expired(data, date_column):
    try:
        today = datetime.today()
        expired = data[data[date_column] < today].copy()
        expired["Days Overdue"] = (today - expired[date_column]).dt.days
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
    .small-button button {
        padding: 5px 10px;
        font-size: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Add app header image
st.image("SPH_Rastar_image_landscape-1-removebg-preview.png", use_container_width=True)  # Replace with your image file name

# Add title and subheading
st.markdown("<h1 style='text-align: center; color: blue;'>SPH's Expiry Alerts</h1>", unsafe_allow_html=True)
st.markdown("### 🚨 Quickly check visa and registration expiration dates and act on them!")

# Sidebar content
st.sidebar.title("Navigation")
st.sidebar.info("Use this app to identify visa and registration expiry alerts for students.")
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

# Main logic
if uploaded_file:
    try:
        # Read and preview uploaded file
        df = pd.read_excel(uploaded_file)
        st.write("Uploaded File Preview:")
        st.dataframe(df, height=300)

        # User input for the date columns and email column
        visa_date_column = st.sidebar.selectbox("Select the Visa Expiry Date Column", df.columns)
        registration_date_column = st.sidebar.selectbox("Select the Registration Expiry Date Column", df.columns)
        email_column = st.sidebar.selectbox("Select the Email Column", df.columns)

        # Convert date columns to datetime
        df[visa_date_column] = pd.to_datetime(df[visa_date_column], format="%d-%m-%Y", errors="coerce")
        df[registration_date_column] = pd.to_datetime(df[registration_date_column], format="%d-%m-%Y", errors="coerce")

        # Get students with visa expiry
        visa_expiring_students = get_students_with_expiry_soon(df, visa_date_column)
        visa_expired_students = get_students_with_expired(df, visa_date_column)

        # Get students with registration expiry
        registration_expiring_students = get_students_with_expiry_soon(df, registration_date_column)
        registration_expired_students = get_students_with_expired(df, registration_date_column)

        # Display metrics
        st.markdown("## 📊 Key Metrics")
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("Total Records", len(df))
        col2.metric("Visa Expiring Soon", len(visa_expiring_students))
        col3.metric("Visa Already Expired", len(visa_expired_students))
        col4.metric("Reg. Expiring Soon", len(registration_expiring_students))
        col5.metric("Reg. Already Expired", len(registration_expired_students))

        # Display visa expiry details
        st.markdown("## 🟠 Students with Visa Expiring Soon:")
        if not visa_expiring_students.empty:
            for i, row in visa_expiring_students.iterrows():
                student_name = row["Student Name"]
                days_remaining = int(row["Days Remaining"])
                recipient_email = row[email_column]

                col1, col2, col3 = st.columns([4, 1, 1])
                col1.write(f"**{student_name}** - {days_remaining} days remaining")
                if col2.button("Preview Visa Email", key=f"visa_preview_{i}"):
                    st.info(
                        f"""
                        **Preview Email:**
                        Subject: Visa Renewal Notification

                        Dear {student_name},

                        Your visa is expiring in {days_remaining} days. Please gather the necessary documents and be ready for the renewal process.

                        Thanks and regards,
                        Study Palace Hub Team
                        """
                    )
                if col3.button("Send Visa Email", key=f"visa_send_{i}"):
                    result = send_email(recipient_email, student_name, days_remaining, "Visa")
                    if result is True:
                        st.success(f"Visa email sent to {student_name} ({recipient_email}).")
                    else:
                        st.error(f"Error sending visa email to {recipient_email}: {result}")
        else:
            st.info("No students have a visa expiring in the next month.")

        # Display registration expiry details
        st.markdown("## 🟠 Students with Registration Expiring Soon:")
        if not registration_expiring_students.empty:
            for i, row in registration_expiring_students.iterrows():
                student_name = row["Student Name"]
                days_remaining = int(row["Days Remaining"])
                recipient_email = row[email_column]

                col1, col2, col3 = st.columns([4, 1, 1])
                col1.write(f"**{student_name}** - {days_remaining} days remaining")
                if col2.button("Preview Reg Email", key=f"reg_preview_{i}"):
                    st.info(
                        f"""
                        **Preview Email:**
                        Subject: Registration Renewal Notification

                        Dear {student_name},

                        Your registration is expiring in {days_remaining} days. Please gather the necessary documents and be ready for the renewal process.

                        Thanks and regards,
                        Study Palace Hub Team
                        """
                    )
                if col3.button("Send Reg Email", key=f"reg_send_{i}"):
                    result = send_email(recipient_email, student_name, days_remaining, "Registration")
                    if result is True:
                        st.success(f"Registration email sent to {student_name} ({recipient_email}).")
                    else:
                        st.error(f"Error sending registration email to {recipient_email}: {result}")
        else:
            st.info("No students have a registration expiring in the next month.")

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")

# Footer
st.markdown("<hr><p style='text-align: center;'>© 2025 SPH Team</p>", unsafe_allow_html=True)
