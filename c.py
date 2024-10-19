import pandas as pd
import re
import streamlit as st
from io import BytesIO

# Streamlit App Title
st.title("Extract Phone Numbers from Excel File")

# File uploader for Excel files
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "csv"])

# Function to extract phone numbers from text
def extract_phone_numbers(text):
    phone_pattern = r'(\b\d{5}\s\d{5}\b|\b\d{10}\b)'  # Matches '12345 67890' or '1234567890'
    phone_numbers = re.findall(phone_pattern, str(text))
    return phone_numbers

# Function to convert DataFrame to Excel and return as a downloadable link
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Phone Numbers')
        writer.close()
    processed_data = output.getvalue()
    return processed_data

# Check if a file is uploaded
if uploaded_file is not None:
    try:
        # Read all sheets from the Excel file
        sheets_dict = pd.read_excel(uploaded_file, sheet_name=None)

        # Initialize an empty list to store all extracted phone numbers
        all_phone_numbers = []

        # Loop through each sheet
        for sheet_name, df in sheets_dict.items():
            # Loop through every cell in the DataFrame
            for index, row in df.iterrows():
                for cell in row:
                    # Extract phone numbers from the cell
                    phone_numbers = extract_phone_numbers(cell)
                    if phone_numbers:
                        all_phone_numbers.extend(phone_numbers)

        # Remove duplicates
        all_phone_numbers = list(set(all_phone_numbers))

        # Display the extracted phone numbers
        if all_phone_numbers:
            st.write(f"{len(all_phone_numbers)} phone numbers extracted:")
            for number in all_phone_numbers:
                st.write(number)

            # Create a DataFrame from the phone numbers
            phone_numbers_df = pd.DataFrame(all_phone_numbers, columns=["Phone Numbers"])

            # Convert the DataFrame to an Excel file
            excel_data = convert_df_to_excel(phone_numbers_df)

            # Provide a download button for the Excel file
            st.download_button(label="Download Phone Numbers as Excel", 
                               data=excel_data, 
                               file_name="extracted_phone_numbers.xlsx", 
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.write("No phone numbers found.")
    except Exception as e:
        st.error(f"Error reading the file: {e}")
else:
    st.write("Please upload an Excel file.")
