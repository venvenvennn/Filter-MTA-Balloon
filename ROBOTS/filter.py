import os
import pandas as pd
import streamlit as st
from werkzeug.utils import secure_filename
from zipfile import ZipFile
import msoffcrypto
import io

# Configure allowed extensions
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

# Helper function to check file extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Function to decrypt an Excel file
def decrypt_excel(file, password):
    try:
        # Load the encrypted file
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)  # Provide the password here

        # Decrypt the file and load into a BytesIO object
        decrypted_file = io.BytesIO()
        office_file.decrypt(decrypted_file)

        # Move the file pointer to the beginning
        decrypted_file.seek(0)
        return decrypted_file
    except Exception as e:
        st.error(f"Failed to decrypt file: {str(e)}")
        return None

# Function to clean the Excel data and create filtered versions
def clean_excel(file, output_path, password=None):
    # If password is provided, decrypt the file first
    if password:
        file = decrypt_excel(file, password)
        if not file:
            return None, []

    df = pd.read_excel(file, engine='openpyxl')

    # Normalization and filtering steps
    if 'PLACEMENT' in df.columns:
        df = df[df['PLACEMENT'] != 'N/A']

    if 'ACTION' in df.columns:
        df = df[df['ACTION'] != 'EXCLUDE IN REPORT']
        df.loc[df['ACTION'] != 'PTP', ['AMT', 'WHEN']] = pd.NA

    if 'REACTION' in df.columns:
        df['REACTION'] = df['REACTION'].replace(0, pd.NA)

    # Convert to BytesIO to avoid saving on disk
    cleaned_file = io.BytesIO()
    df.to_excel(cleaned_file, index=False, engine='openpyxl')

    madpl_files = []

    if 'PLACEMENT' in df.columns:
        madpl_150dpd_df = df[df['PLACEMENT'] == 'MADPL 150DPD']
        madpl_150dpd_file = io.BytesIO()
        madpl_150dpd_df.to_excel(madpl_150dpd_file, index=False, engine='openpyxl')
        madpl_files.append(madpl_150dpd_file)

        madpl1_df = df[df['PLACEMENT'] == 'MADPL1']
        madpl1_file = io.BytesIO()
        madpl1_df.to_excel(madpl1_file, index=False, engine='openpyxl')
        madpl_files.append(madpl1_file)

    # Move the pointer to the beginning of the cleaned file for later download
    cleaned_file.seek(0)

    return cleaned_file, madpl_files

# Streamlit UI logic for app1
def app1_ui():
    st.title('Excel File Cleaner and Filter')
    st.write('MADPL 120 & 150 DAILY REPORT')

    uploaded_file = st.file_uploader("Upload an Excel file", type=['xls', 'xlsx'])

    # Input for password if the file is password protected
    password = st.text_input("Enter the password for the Excel file (if applicable):", type="password")

    if uploaded_file is not None:
        st.write("Processing your file...")

        # Process the file
        cleaned_file, madpl_files = clean_excel(uploaded_file, None, password)

        if cleaned_file is None:
            return  # Exit if file decryption fails

        # Create a zip file containing the cleaned and filtered files
        zip_filename = f"{uploaded_file.name.replace('.xlsx', '')}_cleaned_files.zip"
        zip_filepath = io.BytesIO()

        with ZipFile(zip_filepath, 'w') as zipf:
            # Save the cleaned file to the zip
            cleaned_file.seek(0)  # Ensure we start from the beginning of the BytesIO object
            zipf.writestr(f"{uploaded_file.name.replace('.xlsx', '')}_cleaned.xlsx", cleaned_file.read())

            # Add MADPL files to the zip with unique names
            for i, madpl_file in enumerate(madpl_files, start=1):
                madpl_file.seek(0)  # Ensure we start from the beginning of the BytesIO object
                # Use an index to make each MADPL file's name unique
                madpl_filename = f"{uploaded_file.name.replace('.xlsx', '')}_MADPL_file_{i}.xlsx"
                zipf.writestr(madpl_filename, madpl_file.read())

        # Provide a download link for the zip file
        zip_filepath.seek(0)  # Ensure we start from the beginning of the BytesIO object
        st.download_button(
            label="Download Cleaned Files",
            data=zip_filepath,
            file_name=zip_filename,
            mime="application/zip"
        )

        # Optionally, display the first few rows of the cleaned DataFrame
        cleaned_df = pd.read_excel(cleaned_file, engine='openpyxl')
        st.write("Here are the first few rows of the cleaned data:")
        st.dataframe(cleaned_df.head())

    else:
        st.info("Please upload an Excel file to get started.")