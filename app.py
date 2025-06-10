import streamlit as st
from cash_up_core import run_automation, EXCEL_FILE_PATH # Import the function and Excel path
import os # For checking if file exists for download

st.set_page_config(page_title="Cash Up Automator", layout="centered")

st.title("💰 Cash Up Data Automation")
st.markdown("Click the button below to fetch the latest cash up summary email and add it to your Excel sheet.")

# --- User Inputs (from Streamlit secrets) ---
# It's good practice to display these for verification, but handle them securely.
try:
    GMAIL_USER = st.secrets["GMAIL_USER"]
    GMAIL_APP_PASSWORD = st.secrets["GMAIL_APP_PASSWORD"]
    # CASH_UP_SENDER_EMAIL is no longer loaded or used here
    EMAIL_SUBJECT = st.secrets["EMAIL_SUBJECT"]
except KeyError as e:
    st.error(f"Error loading secrets: {e}. Please ensure your `.streamlit/secrets.toml` file is correctly configured with GMAIL_USER, GMAIL_APP_PASSWORD, and EMAIL_SUBJECT.")
    st.stop() # Stop the app if secrets are missing

# Button to trigger the automation
if st.button("Process Latest Cash Up Email"):
    st.info("Processing... Please wait.")

    # Run the automation logic
    status, messages = run_automation()

    # Display messages
    if status == "Success":
        st.success("Automation Completed Successfully!")
    elif status == "No New Emails": # Updated status from core.py
        st.warning("No new cash up summary emails found matching your criteria.")
    else: # Error or Login Failed
        st.error(f"Automation encountered an issue: {status}")

    for msg in messages:
        st.write(msg) # Display each message from the core script

    # Option to download the Excel file (optional, for local testing)
    try:
        if os.path.exists(EXCEL_FILE_PATH):
            with open(EXCEL_FILE_PATH, "rb") as file:
                btn = st.download_button(
                        label="Download Updated Excel File",
                        data=file,
                        file_name=os.path.basename(EXCEL_FILE_PATH),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.warning(f"Excel file '{os.path.basename(EXCEL_FILE_PATH)}' not found. It might not have been created yet.")
    except Exception as e:
        st.error(f"Could not prepare Excel file for download: {e}")

st.markdown("---")
st.write("Ensure your `secrets.toml` has the correct Gmail credentials and email subject filter.")
st.write("This app processes only *unseen* emails matching the subject. If you want to re-process an email, you'll need to manually mark it as unread in Gmail.")