# app2.py

import streamlit as st
import pandas as pd
import pyperclip  # Import the pyperclip module to interact with the clipboard
import msoffcrypto
import io

# Function to handle the decryption logic for password-protected Excel files
def decrypt_excel(file, password):
    try:
        decrypted_file = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)  # Provide the password here
        office_file.decrypt(decrypted_file)

        decrypted_file.seek(0)  # Move the file pointer to the beginning
        return decrypted_file
    except Exception as e:
        st.error(f"Failed to decrypt file: {str(e)}")
        return None

# Function to handle the extraction logic
def extract_data_from_excel(file, address, password=None):
    try:
        if password:
            file = decrypt_excel(file, password)
            if not file:
                return None
        
        df = pd.read_excel(file, engine='openpyxl', header=None)

        extracted_data = {
        "NAME": df.iloc[0, 1],        
        "ACCOUNT NUMBER": df.iloc[1, 1], 
        "ADDRESS": address,           
        "TOTAL/FACE": df.iloc[23, 1],  
        "DP AMOUNT": df.iloc[10, 1],   
        "DP DATE": pd.to_datetime(df.iloc[11, 1]).strftime('%m/%d/%Y'),     
        "REM AMOUNT": df.iloc[18, 1],  
        "TERM": df.iloc[12, 1],        
        "MA": df.iloc[19, 1],          
        "START": pd.to_datetime(df.iloc[20, 1]).strftime('%m/%d/%Y'),     
        "DAY": pd.to_datetime(df.iloc[20, 1]).day,  # Extract the day number from the START date
        "END": pd.to_datetime(df.iloc[21, 1]).strftime('%m/%d/%Y')           
        }

        return extracted_data


    except Exception as e:
        st.error(f"Error extracting data: {str(e)}")
        return None

# Function to copy the extracted data to the clipboard in a horizontal format
def copy_data_to_clipboard(extracted_data):
    try:
        data_str = "\t".join([str(value).zfill(6) if isinstance(value, str) and value.isdigit() else str(value) for value in extracted_data.values()])
        pyperclip.copy(data_str)
        return True
    except Exception as e:
        st.error(f"Error copying data to clipboard: {str(e)}")
        return False

# Main function to build the Streamlit app UI for app2
def main():
    st.title("Excel File Data Extractor (Straight Term)")
    st.write('STRAIGHT TERM AGREEMENT')
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])
    password = st.text_input("Enter the password for the Excel file (if applicable):", type="password", key="text_input_area")
    
    if uploaded_file:
        st.write(f"Uploaded file: {uploaded_file.name}")
        
        address = st.text_area("Paste the address here:", "")

        extracted_data = extract_data_from_excel(uploaded_file, address, password)
        
        if extracted_data:
            st.subheader("Extracted Data:")
            st.write(extracted_data)

            if st.button("Copy Data to Clipboard", key="copy_to_clip_straight"):
                success = copy_data_to_clipboard(extracted_data)
                if success:
                    st.success("Data successfully copied to clipboard in horizontal format.")
                else:
                    st.error("Failed to copy data to clipboard.")
