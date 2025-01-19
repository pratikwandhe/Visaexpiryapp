import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText


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
        return str(e)  # Return error message


# Streamlit App
st.markdown(
    """
    <style>
    body {
        background-color: #f0f8ff;
    }
    footer {visibility: hidden;}
    .dataframe-container {
        margin: auto;
        width: 90%;
        border: 1px solid #ddd;
        border-radius: 8px;
        padding: 10px;
        background-color: #f9f9f9;
    }
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

if uploaded_file:
    try:
        # Read and preview uploaded file
        df = pd.read_excel(uploaded_file)
        st.write("Uploaded File Preview:")
        st.dataframe(df, height=300)

        # User input for the date column
        date_column = st.sidebar.selectbox("Select the Visa Expiry Date Column", df.columns)
        email_column = st.sidebar.selectbox("Select the Email Column", df.columns)

        # Convert the date column to datetime format
        df[date_column] = pd.to_datetime(df[date_column], format="%d-%m-%Y", errors="coerce")

        # Calculate days remaining
        today = datetime.today()
        df["Days Remaining"] = (df[date_column] - today).dt.days

        # Separate expiring soon and expired students
        expiring_students = df[(df["Days Remaining"] > 0) & (df["Days Remaining"] <= 30)]
        expired_students = df[df["Days Remaining"] < 0]

        # Display metrics
        st.markdown("## 📊 Key Metrics")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Records", len(df))
        col2.metric("Expiring Soon", len(expiring_students))
        col3.metric("Already Expired", len(expired_students))

        # Display expiring soon students in a table
        if not expiring_students.empty:
            st.markdown("## 🟠 Students with Visa Expiring Soon:")
            expiring_students = expiring_students.reset_index(drop=True)
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
                    if result is True:
                        st.success(f"Email sent to {student_name} ({recipient_email}).")
                    else:
                        st.error(f"Error sending email to {recipient_email}: {result}")

        else:
            st.info("No students have a visa expiring in the next 30 days.")

        # Display expired students
        if not expired_students.empty:
            st.markdown("## 🔴 Students Whose Visa Has Already Expired:")
            expired_students = expired_students.reset_index(drop=True)
            for i, row in expired_students.iterrows():
                student_name = row["Student Name"]
                recipient_email = row[email_column]

                col1, col2 = st.columns([5, 2])
                col1.write(f"**{student_name}** - Visa expired")
                if col2.button("Send Reminder Email", key=f"send_reminder_{i}"):
                    result = send_email(recipient_email, student_name, 0)
                    if result is True:
                        st.success(f"Reminder email sent to {student_name} ({recipient_email}).")
                    else:
                        st.error(f"Error sending email to {recipient_email}: {result}")

    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")

# Footer
st.markdown("<hr><p style='text-align: center;'>© 2025 SPH Team</p>", unsafe_allow_html=True)
