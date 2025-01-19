import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email.mime.text import MIMEText


# Function to send email
def send_email(recipient_email, student_name, days_remaining):
    sender_email = "pratikwandhe9095@gmail.com"  # Replace with your Gmail address
    sender_password = "fixx dnwn jpin bwix"  # Replace with your Gmail App Password

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

        # User input for the date column and email column
        date_column = st.sidebar.selectbox("Select the Visa Expiry Date Column", df.columns)
        email_column = st.sidebar.selectbox("Select the Email Column", df.columns)

        # Process data
        df[date_column] = pd.to_datetime(df[date_column], format="%d-%m-%Y", errors="coerce")

        # Get students with visa expiring soon
        expiring_students = get_students_with_visa_expiry_soon(df, date_column)

        # Get students with expired visa
        expired_students = get_students_with_expired_visa(df, date_column)

        # Display metrics
        st.markdown("## 📊 Key Metrics")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Records", len(df))
        col2.metric("Expiring Soon", len(expiring_students))
        col3.metric("Already Expired", len(expired_students))

        # Display expiring soon students
        if not expiring_students.empty:
            st.markdown("## 🟠 Students with Visa Expiring Soon (Sorted by Days Remaining):")
            for i, row in expiring_students.iterrows():
                student_name = row["Student Name"]
                days_remaining = int(row["Days Remaining"])
                recipient_email = row[email_column]

                col1, col2, col3 = st.columns([4, 1, 1])
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
