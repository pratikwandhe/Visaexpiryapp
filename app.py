import streamlit as st
import pandas as pd
from datetime import datetime, timedelta


# Function to filter students with visa expiry in around 1 month and highlight rows
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
        # Convert date column to datetime (assuming dd-mm-yyyy format)
        data[date_column] = pd.to_datetime(data[date_column],
                                           format='%d-%m-%Y',
                                           errors='coerce')

        # Calculate today's date and the target date range
        today = datetime.today()
        target_date = today + timedelta(days=days)

        # Filter students whose visa expiry date is within the next `days`
        expiring_soon = data[(data[date_column] >= today)
                             & (data[date_column] <= target_date)]
        return expiring_soon
    except Exception as e:
        st.error(f"Error processing the data: {e}")
        return pd.DataFrame()


# Streamlit App
st.title("Visa Expiry Checker")

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)

        # Display the first few rows of the uploaded file with scrollable view
        st.write("Uploaded File Preview:")
        st.dataframe(df, height=400)

        # User input for the date column
        date_column = st.selectbox("Select the Visa Expiry Date Column",
                                   df.columns)

        # Check button
        if st.button("Check"):
            # Convert date column to datetime
            df[date_column] = pd.to_datetime(df[date_column],
                                             format='%d-%m-%Y',
                                             errors='coerce')

            # Highlight rows based on expiry dates
            styled_df = df.style.apply(
                lambda row: highlight_expiry_dates(row, date_column), axis=1)

            # Display the styled dataframe
            st.write("Highlighted Visa Expiry Dates:")
            st.dataframe(styled_df, height=400)

            # Get students with visa expiring soon
            expiring_students = get_students_with_visa_expiry_soon(
                df, date_column)

            if not expiring_students.empty:
                st.success("Students with visa expiring soon:")
                st.dataframe(expiring_students)
            else:
                st.info("No students have a visa expiring in the next month.")
    except Exception as e:
        st.error(f"An error occurred: {e}")
else:
    st.info("Please upload an Excel file to proceed.")
