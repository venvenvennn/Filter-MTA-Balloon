import pandas as pd
import streamlit as st
import pyperclip  # Import the pyperclip module to interact with the clipboard
import msoffcrypto
import io

# Function to handle the decryption logic for password-protected Excel files
def decrypt_excel(file, password):
    try:
        # Open the encrypted file
        decrypted_file = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)  # Provide the password here
        office_file.decrypt(decrypted_file)

        # Move the file pointer to the beginning
        decrypted_file.seek(0)

        return decrypted_file
    except Exception as e:
        st.error(f"Failed to decrypt file: {str(e)}")
        return None

# Function to handle the extraction logic
def extract_data_from_excel(file, address, password=None):
    try:
        # If password is provided, decrypt the file first
        if password:
            file = decrypt_excel(file, password)
            if not file:
                return None
        
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file, engine='openpyxl', header=None)

        # Function to safely convert date to string format (handle NaT or None)
        def safe_date_format(value):
            if pd.isna(value) or value is None:  # Handle NaT or None
                return ""
            return pd.to_datetime(value).strftime('%m/%d/%Y')

        # Function to extract the calendar date (day of the month)
        def extract_day(value):
            if pd.isna(value) or value is None:  # Handle NaT or None
                return ""
            return pd.to_datetime(value).strftime('%d')  # Returns the day of the month

        # Extract data based on the provided mapping
        total_face = df.iloc[23, 1]  # B24: Total/Face
        downpayment = df.iloc[10, 1]  # B11: DP AMOUNT

        # Calculate the remaining balance as Total/Face - Downpayment
        rem_balance = total_face - downpayment if pd.notna(total_face) and pd.notna(downpayment) else ""

        # Extract the MA1, MA2, and MA3 values based on the format you described
        ma_values = df.iloc[19, 1].split(";")  # B20: MA values

        # Extract the first value (before 'x') for each MA
        ma1 = ma_values[0].split("x")[0].strip() if len(ma_values) > 0 else ""
        ma2 = ma_values[1].split("x")[0].strip() if len(ma_values) > 1 else ""
        ma3 = ma_values[2].split("x")[0].strip() if len(ma_values) > 2 else ""

        # Extract the term value
        term_value = df.iloc[12, 1]  # B13: TERM
        if pd.notna(term_value):
            term_value = float(term_value)
        else:
            term_value = 0

        # Calculate the division result (Term value / 3)
        months_per_term = term_value / 3 if term_value else 0

        # Extract the start date for MA1
        start_date_ma1 = pd.to_datetime(df.iloc[20, 1])  # B21

        # Calculate the end date for MA1 as Start + 11 months
        end_date_ma1 = start_date_ma1 + pd.DateOffset(months=11)

        # For MA2, Start date will be 1 month after the End date of MA1
        start_date_ma2 = end_date_ma1 + pd.DateOffset(months=1)
        # End date for MA2 is Start + 11 months
        end_date_ma2 = start_date_ma2 + pd.DateOffset(months=11)

        # For MA3, Start date will be 1 month after the End date of MA2
        start_date_ma3 = end_date_ma2 + pd.DateOffset(months=1)
        # End date for MA3 is Start + 11 months
        end_date_ma3 = start_date_ma3 + pd.DateOffset(months=11)

        # Extract the day of the calendar date for START, START 2, and START 3
        start_day_ma1 = extract_day(start_date_ma1)
        start_day_ma2 = extract_day(start_date_ma2)
        start_day_ma3 = extract_day(start_date_ma3)

        # Format the START, DAY, MONTH, and END dates for MA 1, MA 2, and MA 3
        extracted_data = {
            "NAME": df.iloc[0, 1],        # B1
            "ACCOUNT NUMBER": df.iloc[1, 1], # B2
            "ADDRESS": address,           # Use the provided address
            "TOTAL/FACE": total_face,  # B24
            "DP AMOUNT": downpayment,   # B11
            "DP DATE": safe_date_format(df.iloc[11, 1]),     # B12
            "REM BAL": rem_balance,  # Remaining balance
            "TERM": term_value,        # B13
            "MA 1": ma1,                   # MA 1 (first number before 'x')
            "START": safe_date_format(start_date_ma1),     # Start date for MA 1
            "DAY": start_day_ma1,                     # Day for MA 1
            "MONTH": months_per_term,  # Divided Term value, applied for MA1's Month
            "END": safe_date_format(end_date_ma1),           # End date for MA 1
            "MA 2": ma2,                   # MA 2 (first number before 'x')
            "START 2": safe_date_format(start_date_ma2),     # Start date for MA 2 (1 month after MA1's End)
            "DAY 2": start_day_ma2,                  # Day for MA 2
            "MONTH 2": months_per_term,  # Divided Term value, applied for MA2's Month
            "END 2": safe_date_format(end_date_ma2),           # End date for MA 2
            "MA 3": ma3,                   # MA 3 (first number before 'x')
            "START 3": safe_date_format(start_date_ma3),     # Start date for MA 3 (1 month after MA2's End)
            "DAY 3": start_day_ma3,                  # Day for MA 3
            "MONTH 3": months_per_term,  # Divided Term value, applied for MA3's Month
            "END 3": safe_date_format(end_date_ma3)           # End date for MA 3
        }
        return extracted_data

    except Exception as e:
        st.error(f"Error extracting data: {str(e)}")
        return None

# Function to copy the extracted data to the clipboard in a horizontal format
def copy_data_to_clipboard(extracted_data):
    try:
        # Convert all extracted data values to string explicitly to preserve leading zeros
        data_str = "\t".join([str(value).zfill(6) if isinstance(value, str) and value.isdigit() else str(value) for value in extracted_data.values()])
        
        # Use pyperclip to copy the string to the clipboard
        pyperclip.copy(data_str)
        return True
    except Exception as e:
        st.error(f"Error copying data to clipboard: {str(e)}")
        return False


# Streamlit app structure
def main():
    st.title("Excel File Data Extractor (Balloon Term)")
    st.write('BALLOON TERM AGREEMENT (3 YEARS)')
 
    # File upload - Add a unique key to avoid duplicate ID errors
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"], key="unique_file_uploader_key")

    if uploaded_file:
        # Show the file name
        st.write(f"Uploaded file: {uploaded_file.name}")
        
        # Allow the user to paste the address
        address = st.text_area("Paste the address here:", key="address_text_area")
        
        # Password input for encrypted Excel files
        password = st.text_input("Enter the password for the Excel file (if applicable):", type="password", key="password_text_area")

        # Extract data from the uploaded file
        extracted_data = extract_data_from_excel(uploaded_file, address, password)
        
        if extracted_data:
            # Display extracted data
            st.subheader("Extracted Data:")
            st.write(extracted_data)

            # Add a button to copy the data to clipboard in horizontal format
            if st.button("Copy Data to Clipboard", key="copy_to_clip_balloon"):
                # Copy the extracted data to clipboard in horizontal format
                success = copy_data_to_clipboard(extracted_data)

                if success:
                    st.success("Data successfully copied to clipboard in horizontal format.")
                else:
                    st.error("Failed to copy data to clipboard.")

# Run the Streamlit app
if __name__ == "__main__":
    main()
