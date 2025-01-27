import dateutil.utils
import pandas as pd
import streamlit
import streamlit as st
import time
from io import BytesIO
import uuid
import re
from dateutil.utils import today
from requests import delete
from streamlit import session_state
import numpy as np
import math
import openpyxl






def load_file(uploaded_file):
    """
    Load the file into a DataFrame.
    Handles CSV and Excel files.
    """
    try:
        if uploaded_file.name.endswith(".csv"):
            return pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(".xlsx"):
            return pd.read_excel(uploaded_file)
        else:
            st.error(f"Unsupported file type: {uploaded_file.name}")
            return None
    except Exception as e:
        st.error(f"Error reading file {uploaded_file.name}: {e}")
        return None


def process_uploaded_files(uploaded_files):
    """
    Process the uploaded files and return the corresponding DataFrames.
    """
    properties_df = None
    beneficiaries_df = None
    tenants_df = None
    invoices_df = None
    payments_df = None
    master_df = None

    success_message = st.empty()

    for uploaded_file in uploaded_files:
        file_df = load_file(uploaded_file)
        if file_df is not None:
            if "properties" in uploaded_file.name.lower():
                properties_df = file_df
                # st.success(f"Successfully loaded {uploaded_file.name}")
            elif "beneficiaries" in uploaded_file.name.lower():
                beneficiaries_df = file_df
            # st.success(f"Successfully loaded {uploaded_file.name}")
            elif "tenants" in uploaded_file.name.lower():
                tenants_df = file_df
            # st.success(f"Successfully loaded {uploaded_file.name}")

            elif "invoices" in uploaded_file.name.lower():
                invoices_df = file_df
            #  st.success(f"Successfully loaded {uploaded_file.name}")
            elif "payments" in uploaded_file.name.lower():
                payments_df = file_df
            #  st.success(f"Successfully loaded {uploaded_file.name}")
            elif "master" in uploaded_file.name.lower():
                master_df = file_df
                master_df["Unique ID"] = [f"{i + 1:05d}" for i in range(len(master_df))]


            #  st.success(f"Successfully loaded {uploaded_file.name}")

        st.session_state['properties_df'] = properties_df
        st.session_state['beneficiaries_df'] = beneficiaries_df
        st.session_state['tenants_df'] = tenants_df
        st.session_state['invoices_df'] = invoices_df
        st.session_state['payments_df'] = payments_df
        st.session_state['master_df'] = master_df

    return properties_df, beneficiaries_df, tenants_df, invoices_df, payments_df, master_df



def clean_master():
    master_df = st.session_state['master_df']  # Assuming master_df is loaded into session state
    agent_name = st.session_state['agent_name']


    # Function to apply proper case
    def proper_case(val):
        if isinstance(val, str):
            val = val.strip().replace('  ', ' ').replace('  ', ' ')  # Remove double spaces
            return val.title()  # Capitalize each word
        return val

    # Function to convert to uppercase
    def upper_case(val):
        if isinstance(val, str):
            val = val.strip()
            val = re.sub(r'\s+', ' ', val)  # Replace multiple spaces with a single space
            return val.upper()  # Convert to uppercase
        return val

    # Function to clean up values
    def clean_case(val):
        if isinstance(val, str):
            val = val.strip()
            val = re.sub(r'\s+', ' ', val)  # Replace multiple spaces with a single space
            return val.capitalize()  # Capitalize the first word and lowercase the rest
        return val

    # Function to convert number to string
    def convert_number_to_string(val):
        if isinstance(val, float) and math.isnan(val):
            return ""

        # Handle floats specifically
        if isinstance(val, float):
            # If the number is effectively an integer (like 123.0), treat it as an integer
            if val.is_integer():
                return str(int(val)).replace(',', '')
            return str(val).replace('.', '').replace(',', '')

        # For integers, just return as a string (removes commas if any)
        if isinstance(val, int):
            return str(val).replace(',', '')

        # If it's anything else, just return as a string
        return str(val)


    # Function to handle number formatting to 2 decimals
    def number_format(val, decimals=2):
        try:
            return round(float(val), decimals)
        except (ValueError, TypeError):
            return ""  # Or return None if preferred

    # Function to handle date formatting
    def format_date(val):
        try:
            return pd.to_datetime(val).strftime('%Y/%m/%d')
        except (ValueError, TypeError):
            return val

    # Function to update tenant address with property address if tenant address is None
    def update_tenant_address(row):
        # List of tenant address columns
        tenant_columns = [
            'Tenant Address 1 (Same as Property Address. Leave blank)',
            'Tenant Address 2',
            'Tenant Address 3',
            'Tenant City',
            'Tenant County',
            'Tenant Postcode',
            'Tenant Country'
        ]

        # Check if all tenant address columns are None or empty
        if all(pd.isna(row[col]) or row[col] == '' for col in tenant_columns):
            # Copy values from property address columns to tenant address columns
            row['Tenant Address 1 (Same as Property Address. Leave blank)'] = row['Address 1']
            row['Tenant Address 2'] = row['Address 2']
            row['Tenant Address 3'] = row['Address 3']
            row['Tenant City'] = row['City']
            row['Tenant County'] = row['County']
            row['Tenant Postcode'] = row['Postcode']
            row['Tenant Country'] = row['Country']

        return row

    # Function to clean the VAT Included? column
    def clean_vat_included(val):
        if isinstance(val, str):
            val = val.strip().lower()  # Remove any extra spaces and convert to lowercase
            if val in ['no', 'yes']:
                return 'N' if val == 'no' else 'Y'
        return val

    # Apply formatting rules to columns
    formatting_rules = {
        'Property Name': proper_case,
        'Address 1': proper_case,
        'Address 2': proper_case,
        'Address 3': proper_case,
        'City': proper_case,
        'County': proper_case,
        'Postcode': upper_case,
        'Country': upper_case,
        'Rent Amount': lambda x: number_format(x, 2),
        'Rent Due Date': lambda x: number_format(x, 0),
        'Rent Frequency': clean_case,
        'Next Due Date (If Not Monthly)': format_date,
        'Tenancy End Date': format_date,
        'Commission/Management Percentage': number_format,
        'Or Fixed Management Fee': lambda x: number_format(x, 2),
        'VAT Included?': clean_vat_included,  # Apply custom function to this column
        'Agreed Float Amount': lambda x: number_format(x, 2),
        'Service Type/Branch': proper_case,
        'Landlord(s) Name': proper_case,
        'Are The Rent Proceeds Split Between Multiple Bank Accounts? (100% Or 50/50, Etc.)': number_format,
        'Landlord E-Mail Address': lambda x: x.lower() if isinstance(x, str) else x,
        'Additional Landlord E-Mail CC': lambda x: x.lower() if isinstance(x, str) else x,
        'Landlord Mobile Number': convert_number_to_string,
        'Landlord Additional Phone Number': convert_number_to_string,
        'Landlord Contact Address 1': proper_case,
        'Landlord Contact Address 2': proper_case,
        'Landlord Contact Address 3': proper_case,
        'Landlord Contact City': proper_case,
        'Landlord Contact County': proper_case,
        'Landlord Contact Postcode': upper_case,
        'Landlord Country': upper_case,
        'Landlord Account Name': proper_case,
        'Landlord Sort Code (6 Digits No Spaces/Dashes)': convert_number_to_string,
        'Landlord Account Number (8 Digits No Spaces/Dashes)': convert_number_to_string,
        'Additional Bank Information': clean_case,
        'Tenant(s) Name': proper_case,
        'Tenant E-Mail Address': lambda x: x.lower() if isinstance(x, str) else x,
        'Additional Tenant E-Mail CC': lambda x: x.lower() if isinstance(x, str) else x,
        'Tenant Mobile Number': convert_number_to_string,
        'Tenant Additional Phone Number': convert_number_to_string,
        'Tenant Address 1 (Same as Property Address. Leave blank)': proper_case,
        'Tenant Address 2': proper_case,
        'Tenant Address 3': proper_case,
        'Tenant City': proper_case,
        'Tenant County': proper_case,
        'Tenant Postcode': upper_case,
        'Tenant Country': upper_case
    }

    master_df['Commission/Management Percentage'] = master_df['Commission/Management Percentage'] * 100
    master_df['Are The Rent Proceeds Split Between Multiple Bank Accounts? (100% Or 50/50, Etc.)'] = master_df['Are The Rent Proceeds Split Between Multiple Bank Accounts? (100% Or 50/50, Etc.)'] * 100

    if 'Commission/Management Percentage' in master_df.columns:
        master_df['Commission/Management Percentage'] = master_df['Commission/Management Percentage'].apply(
            lambda x: x / 100 if isinstance(x, (int, float)) and x > 100 else x
        )
    if 'Are The Rent Proceeds Split Between Multiple Bank Accounts? (100% Or 50/50, Etc.)' in master_df.columns:
        master_df['Commission/Management Percentage'] = master_df['Commission/Management Percentage'].apply(
            lambda x: x / 100 if isinstance(x, (int, float)) and x > 100 else x
        )

    # Iterate through each column and apply formatting where applicable
    for column, func in formatting_rules.items():
        if column in master_df.columns:
            master_df[column] = master_df[column].apply(func)

    # **Remove values from 'Country' columns if more than two characters**
    country_columns = [col for col in master_df.columns if 'Country' in col]

    for col in country_columns:
        master_df[col] = master_df[col].apply(lambda x: "" if isinstance(x, str) and len(x) > 2 else x)

    # Apply tenant address update function
    master_df = master_df.apply(update_tenant_address, axis=1)

    # **Special handling for Commission/Management Percentage column**
    if 'Commission/Management Percentage' in master_df.columns:
        master_df['Commission/Management Percentage'] = master_df['Commission/Management Percentage'].apply(
            lambda x: x / 100 if isinstance(x, (int, float)) and x > 100 else x
        )

    def clean_mobile_numbers(df):
        # Iterate over all columns
        for column in df.columns:
            # Check if the column name contains 'Mobile Number'
            if 'Mobile Number' in column:
                # Apply a function to clean mobile numbers, skipping None, blank, or empty strings
                df[column] = df[column].apply(
                    lambda x: format_mobile_number(x.strip()) if isinstance(x, str) and x.strip() != '' else x)
        return df

    # Function to format mobile numbers
    def format_mobile_number(number):
        if isinstance(number, str) and number.strip() != '':  # Check for non-empty value
            if number.startswith(('07', '7')):  # For UK mobile numbers
                # Replace the leading 07 or 7 with 44
                return '44' + number[1:] if number.startswith('07') else '44' + number
        return number  # Skip landlines or foreign numbers and invalid entries

    def move_non_uk_mobile_to_additional(df):
        # Iterate through the DataFrame rows
        for index, row in df.iterrows():
            # Check if 'Landlord Mobile Number' is a valid string and not empty
            if isinstance(row['Landlord Mobile Number'], str) and row['Landlord Mobile Number'].strip() != '':
                # Check if 'Landlord Mobile Number' doesn't start with '44'
                if not row['Landlord Mobile Number'].startswith('44'):
                    # Check if 'Landlord Additional Phone Number' already has a value
                    if pd.notna(row['Landlord Additional Phone Number']) and row[
                        'Landlord Additional Phone Number'].strip() != '':
                        # If there is existing data, append the new number with a comma
                        df.at[
                            index, 'Landlord Additional Phone Number'] = f"{row['Landlord Additional Phone Number']}, {row['Landlord Mobile Number']}"
                    else:
                        # If there's no existing data, just set the value without adding a comma
                        df.at[index, 'Landlord Additional Phone Number'] = row['Landlord Mobile Number']

                    # Set 'Landlord Mobile Number' to an empty value (or None)
                    df.at[index, 'Landlord Mobile Number'] = None

            # Repeat the same logic for 'Tenant Mobile Number'
            if isinstance(row['Tenant Mobile Number'], str) and row['Tenant Mobile Number'].strip() != '':
                if not row['Tenant Mobile Number'].startswith('44'):
                    if pd.notna(row['Tenant Additional Phone Number']) and row[
                        'Tenant Additional Phone Number'].strip() != '':
                        df.at[
                            index, 'Tenant Additional Phone Number'] = f"{row['Tenant Additional Phone Number']}, {row['Tenant Mobile Number']}"
                    else:
                        df.at[index, 'Tenant Additional Phone Number'] = row['Tenant Mobile Number']
                    df.at[index, 'Tenant Mobile Number'] = None

        return df


    master_df = clean_mobile_numbers(master_df)
    master_df = move_non_uk_mobile_to_additional(master_df)

    # Assuming master_df is your cleaned DataFrame
    st.session_state['master_df'] = master_df
    st.write(master_df)

    # Create a BytesIO buffer to hold the Excel data
    output = BytesIO()

    # Write the DataFrame to the buffer using pandas ExcelWriter
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        master_df.to_excel(writer, index=False, sheet_name='Master')

    # Ensure the buffer is written completely and reset to the start
    output.seek(0)

    # Create a download button for the Excel file
    st.download_button(
        label="Download Cleaned Master",
        data=output,
        file_name=f"Cleaned Master file {agent_name}.xlsx",  # Ensure the file has the .xlsx extension
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True  # Makes the button span the full width
    )

def build_properties():
    launch_date = st.session_state['launch_date']
    agent_name = st.session_state['agent_name']

    st.info("Properties")

    # Retrieve properties and master data from session state
    properties_df = st.session_state['properties_df']
    master_df = st.session_state['master_df']

    # Remove rows that are completely empty or have blank 'Property name'
    properties_df = properties_df.dropna(subset=['Property name'], how='all')

    # Add unique codes from master_df (this should be after the row deletion)
    properties_df.loc[:, 'Unique ID'] = master_df.loc[properties_df.index, 'Unique ID']

    # Define the column mapping for property information
    column_mapping_property = {
        "Property Name": "Property name",
        "Agent/Branch": "Agent",
        "Service Level": "Service level",
        "Rent Amount": "Monthly payment required",
        "Hold Owner Funds": "Hold owner funds",
        "Agreed Float Amount": "Prop. acc. minimum balance",
        "Address 1": "Address 1",
        "Address 2": "Address 2",
        "Address 3": "Address 3",
        "City": "City",
        "County": "County",
        "Postcode": "Postcode",
        "Country": "Country",
        "Property Customer Reference": "Customer reference"
    }

    # Define input type for each column
    column_input_types = {
        "Property name": "text",
        "Agent": "text",
        "Service level": "select",  # This is for the dropdown selection
        "Monthly payment required": "number",
        "Prop. acc. minimum balance": "number",
        "Hold owner funds": "select",
        "Approval required": "select",
        "Listed from": "date",
        "Listed until": "date",
        "Comment": "text",
        "Allow payments": "select",
        "Allow maintenance": "select", # Update this for 'Y' or 'N' selection
        "Address 1": "text",
        "Address 2": "text",
        "Address 3": "text",
        "City": "text",
        "County": "text",
        "Postcode": "text",
        "Country": "select",
        "Customer Reference": "text",
        "Status": "select",
        "Tags": "text"
    }

    # Precompute the common columns between the mapping and the columns we need
    common_columns = set(master_df.columns).intersection(column_mapping_property.keys())

    # Iterate over the common columns and set the values from master_df to properties_df
    for master_col in common_columns:
        properties_col = column_mapping_property[master_col]

        # Update properties_df with the values from master_df where the Unique ID matches
        for index, row in properties_df.iterrows():
            unique_id = row["Unique ID"]
            # Find the corresponding row in master_df by matching the Unique ID
            master_row = master_df[master_df["Unique ID"] == unique_id]
            if not master_row.empty:
                # Set the value in properties_df to the corresponding value in master_df
                properties_df.at[index, properties_col] = master_row[master_col].values[0]

    if 'Country' in properties_df.columns:
        missing_countries = properties_df['Country'].isna() | (properties_df['Country'] == "")

        with st.expander("Missing country codes"):
            if missing_countries.any():
                st.warning("Missing country codes")
                st.dataframe(properties_df[missing_countries])

                if st.checkbox("Apply 'UK' to all missing/blank countries",key="property auto country"):
                    properties_df.loc[missing_countries, 'Country'] = "UK"
                    st.success("Default value 'UK' applied to all missing/blank countries.")
                    st.session_state['beneficiaries_df'] = properties_df

                elif st.checkbox("Manually add country codes",key='property manual country'):
                    for index in properties_df[missing_countries].index:
                        property_name = properties_df.at[index, 'Property name']
                        country_input = st.text_input(f"Enter country code for {property_name}:", max_chars=2,
                                                      key=f"property_country_input_{index}")
                        if country_input:
                            properties_df.at[index, 'Country'] = country_input

                            st.session_state['properties_df'] = properties_df

        # Check for missing 'Monthly payment required' (NaN or empty string)
        if 'Monthly payment required' in properties_df.columns:
            missing_payment = properties_df['Monthly payment required'].isna() | (
                        properties_df['Monthly payment required'] == "")

            # Expander for missing payment amounts
            with st.expander("Missing Monthly Payment Amounts"):
                if missing_payment.any():
                    st.warning("Missing Monthly Payment Amounts")
                    st.dataframe(properties_df[missing_payment])

                    # Checkbox to apply default value '0.00' to all missing payment entries
                    if st.checkbox("Apply '0.00' to all missing property amounts", key="property_auto_payment"):
                        properties_df.loc[missing_payment, 'Monthly payment required'] = 0.00
                        st.success("Default value '0.00' applied to all missing payment amounts.")
                        st.session_state['properties_df'] = properties_df

                    # Checkbox to manually add payment amounts
                    elif st.checkbox("Manually add payment amounts", key="property_manual_payment"):
                        for index in properties_df[missing_payment].index:
                            property_name = properties_df.at[index, 'Property name']
                            # Input for monthly payment amount
                            payment_input = st.number_input(f"Enter payment amount for {property_name}:", min_value=0.0,
                                                            value=0.0, key=f"property_payment_input_{index}")
                            if payment_input != 0.0:  # If a value is entered, update the DataFrame
                                properties_df.at[index, 'Monthly payment required'] = payment_input
                                st.session_state['properties_df'] = properties_df
                                st.success(f"Payment amount for {property_name} updated to {payment_input}.")


        # Create a BytesIO buffer for download
        output = BytesIO()

        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            # Write the dataframe to the Excel file
            properties_df.to_excel(writer, index=False, sheet_name="Properties")

            # Access the xlsxwriter workbook and worksheet
            workbook = writer.book
            worksheet = writer.sheets["Properties"]

            # Find the index of the 'Business name' column
            property_name_col = properties_df.columns.get_loc('Property name')

            # Freeze the columns starting from 'Business name' (use the column index)
            worksheet.freeze_panes(1, property_name_col + 1)  # Freeze first row and the 'Business name' column

            # Format: Align all data from row 2 downwards to the left
            cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'border': 0})  # Remove borders

            # Apply left alignment for all cells (except the header)
            for row in range(1, len(properties_df) + 1):  # Skipping the header row (which is row 0)
                worksheet.set_row(row, None, cell_format)

            # Auto-size all columns to fit the content
            for col_num, column in enumerate(properties_df.columns.values):
                max_len = max(properties_df[column].astype(str).map(len).max(), len(column)) + 2  # Add some padding
                worksheet.set_column(col_num, col_num, max_len)

            # Set header row width to 48 pixels (approximately 6.4 characters wide, since 1 character ~= 7 pixels)
            header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'border': 0,
            })

            # Add filters to the header row
            worksheet.autofilter(0, 0, 0, len(properties_df.columns) - 1)

    # Move to the beginning of the BytesIO buffer
    output.seek(0)
    xlsx_data = output.read()

    # Create a download button for the Excel file
    st.download_button(
        label="Download Properties",
        data=xlsx_data,
        file_name=f"Properties To Check And Import {agent_name}.xlsx",  # Ensure the file has the .xlsx extension
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True  # Makes the button span the full width
    )


def build_beneficiaries():
    launch_date = st.session_state['launch_date']
    agent_name = st.session_state['agent_name']

    st.info("Beneficiaries")
    beneficiaries_df = st.session_state.get('beneficiaries_df')
    master_df = st.session_state.get('master_df')

    # Remove rows that are completely empty or have blank 'Unique ID' (or other critical columns)
    beneficiaries_df = beneficiaries_df.dropna(subset=['Business name'], how='all')

    beneficiaries_df['Unique ID'] = master_df["Unique ID"]

    column_mapping_beneficiary = {
        "Landlord First Name": "First name",
        "Landlord Last Name": "Last name",
        "Landlord(s) Name": "Business name",
        "Passport no / Driving licence": "Passport no / Driving licence",
        "VAT number": "VAT number",
        "Landlord E-Mail Address": "E-mail address",
        "Additional Landlord E-Mail CC": "E-mail CC",
        "Landlord Mobile Number": "Mobile number",
        "Landlord Additional Phone Number": "Phone number",
        "Landlord Fax Number": "Fax number",
        "Landlord Address 1": "Address 1",
        "Landlord Address 2": "Address 2",
        "Landlord Address 3": "Address 3",
        "Landlord City": "City",
        "Landlord County": "County",
        "Landlord Postcode": "Postcode",
        "Landlord Country": "Country",
        "Landlord Account Name": "Account name",
        "Landlord Bank name": "Bank name",
        "Landlord Sort Code (6 Digits No Spaces/Dashes)": "Sort code",
        "Landlord Swift Code": "SWIFT code",
        "Landlord Branch name": "Branch name",
        "Landlord Account Number (8 Digits No Spaces/Dashes)": "Account number",
        "Landlord Account Number / IBAN Number": "Account number",
        "Landlord Bank Country": "Bank country",
        "Customer reference": "Landlord Customer Reference"
    }

#Build beneficiaries df
    for master_col, beneficiaries_col in column_mapping_beneficiary.items():
        if master_col in master_df.columns:
            # Check if the target beneficiary column exists in beneficiaries_df
            if beneficiaries_col in beneficiaries_df.columns:
                try:
                    mapping = master_df.set_index("Unique ID")[master_col]
                    beneficiaries_df[beneficiaries_col] = beneficiaries_df["Unique ID"].map(mapping)
                except KeyError:
                    pass

    # Drop "Unnamed" columns and "Unique ID" column
    beneficiaries_df = beneficiaries_df.loc[:,
                                       ~beneficiaries_df.columns.str.contains('^Unnamed|Unique ID')]

    beneficiaries_df['Notify text'] = "Y"
    beneficiaries_df['Notify e-mail'] = "Y"



#Check missing country code
    missing_countries = beneficiaries_df['Country'].isna() | (beneficiaries_df['Country'] == "")
    filtered_df = beneficiaries_df.loc[:, beneficiaries_df.notna().any()]
    if missing_countries.any():
     with st.expander("Missing country codes"):
            if missing_countries.any():
                st.warning("Missing country codes")
                st.dataframe(filtered_df[missing_countries])

                if st.checkbox("Apply 'UK' to all missing/blank countries"):
                    beneficiaries_df.loc[missing_countries, 'Country'] = "UK"
                    st.success("Default value 'UK' applied to all missing/blank countries.")
                    st.session_state['beneficiaries_df'] = beneficiaries_df


                elif st.checkbox("Manually add country codes"):
                    for index in beneficiaries_df[missing_countries].index:
                        property_name = beneficiaries_df.at[index, 'Business name']
                        country_input = st.text_input(f"Enter country code for {property_name}:", max_chars=2,
                                                      key=f"country_input_{index}")
                        if country_input:
                            beneficiaries_df.at[index, 'Country'] = country_input
                            st.session_state['beneficiaries_df'] = beneficiaries_df

    st.session_state['beneficiaries_df'] = beneficiaries_df


#Dupliucate beneficiaries

    duplicate_rows = beneficiaries_df[beneficiaries_df.duplicated(subset=['First name', 'Last name','Business name', 'Account number'], keep=False)]
    filtered_df = duplicate_rows.loc[:, beneficiaries_df.notna().any()]
    if not duplicate_rows.empty:
        with (st.expander("Duplicates")):
            st.warning("Duplicate entries found based on 'Business name' and 'Account number'.")
            st.dataframe(filtered_df)
            st.write("Select one option")

            delete_checkbox = st.checkbox("Delete a duplicate row?")

            # If checkbox is ticked, show the multiselect for row selection

            if delete_checkbox:
                row_to_delete = st.multiselect("Select row(s) to delete:", options=duplicate_rows.index)

                # Only proceed to delete if rows are selected
                if row_to_delete:
                    beneficiaries_df = beneficiaries_df.drop(index=row_to_delete)
                    st.success(f"Row(s) {row_to_delete} has been deleted.")
                    st.session_state['beneficiaries_df'] = beneficiaries_df
                else:
                    st.warning("No rows selected for deletion.")

            if st.checkbox("Auto dedupe"):
                beneficiaries_df = beneficiaries_df.drop_duplicates(subset=['First name', 'Last name','Business name', 'Account number'], keep='first')
                st.success("Duplicates successfully removed")
                st.session_state['beneficiaries_df'] = beneficiaries_df

    # Whitelist reqd

    # Create the 'Beneficiary Combination' temporarily
    beneficiaries_df['Beneficiary Combination'] = beneficiaries_df.apply(
        lambda row: (row['First name'], row['Last name'], row['Business name']) if row['Account number'] not in [None,
                                                                                                                 '',
                                                                                                                 ' '] else None,
        axis=1
    )

    # Remove rows with None or empty 'Beneficiary Combination'
    beneficiaries_df = beneficiaries_df[beneficiaries_df['Beneficiary Combination'].notna()]

    # Group by Account number and count unique combinations of beneficiaries
    account_duplication_check = (
        beneficiaries_df.groupby('Account number')['Beneficiary Combination']
        .nunique()
        .reset_index()
    )

    # Filter out account numbers with 3 or more distinct combinations
    duplicate_account_numbers = account_duplication_check[
        account_duplication_check['Beneficiary Combination'] >= 3
        ]

    # Remove 'Beneficiary Combination' column after checking duplicates
    beneficiaries_df.drop(columns=['Beneficiary Combination'], inplace=True)

    # Only show expander if there are flagged account numbers
    if not duplicate_account_numbers.empty:
        with st.expander("Whitelist"):
            st.warning("Whitelist request: Bank account linked to three or more beneficiaries.")

            # Preview rows with duplicate account numbers
            duplicate_accounts = beneficiaries_df[
                beneficiaries_df['Account number'].isin(duplicate_account_numbers['Account number'])]

            # Replacing NaN values with empty strings in 'duplicate_accounts'
            duplicate_accounts = duplicate_accounts.fillna("")  # This line replaces NaN values with ""

            # Remove columns where all values are "" or None for the UI preview (without modifying the original dataframe)
            cleaned_preview = duplicate_accounts.loc[:, (duplicate_accounts != "").any(axis=0)]

            # Display the cleaned dataframe preview
            st.dataframe(cleaned_preview)

            # Select beneficiaries to edit or delete from the flagged account numbers
            beneficiary_options = duplicate_accounts[
                ['First name', 'Last name', 'Business name', 'Account number', 'Sort code']]
            beneficiary_options = beneficiary_options.drop_duplicates(
                subset=['First name', 'Last name', 'Business name', 'Account number'])

            # Multi-select dropdown to delete multiple beneficiaries
            selected_beneficiaries = st.multiselect(
                "Select beneficiaries to delete:",
                options=beneficiary_options['Business name'].tolist(),
                default=[],
                key='delete_whitelist'
            )

            if selected_beneficiaries:
                # Button to delete selected beneficiaries
                if st.button("Delete selected beneficiaries"):
                    # Remove rows where the 'Business name' is in the selected beneficiaries list
                    beneficiaries_df = beneficiaries_df[
                        ~beneficiaries_df['Business name'].isin(selected_beneficiaries)
                    ]

                    # Update the session state with the modified dataframe
                    st.session_state['beneficiaries_df'] = beneficiaries_df

                    st.success(f"Deleted the selected beneficiaries: {', '.join(selected_beneficiaries)}.")
                    st.write(beneficiaries_df)

            # Multi-select dropdown to select beneficiaries to edit
            selected_beneficiaries_to_edit = st.multiselect(
                "Select beneficiaries to edit:",
                options=beneficiary_options['Business name'].tolist(),
                default=[],
                key='edit_whitelist'
            )

            # If beneficiaries are selected for editing
            if selected_beneficiaries_to_edit:
                changes = []

                # Loop through the selected beneficiaries and create editable fields
                for business_name in selected_beneficiaries_to_edit:
                    selected_beneficiary_details = beneficiary_options[
                        beneficiary_options['Business name'] == business_name
                        ].iloc[0]

                    st.info(f"Edit details for: {business_name}")

                    # Editable fields for each selected beneficiary
                    first_name = st.text_input("First Name", value=selected_beneficiary_details['First name'],
                                               key=f"first_name_{business_name}")
                    last_name = st.text_input("Last Name", value=selected_beneficiary_details['Last name'],
                                              key=f"last_name_{business_name}")
                    business_name_input = st.text_input("Business Name",
                                                        value=selected_beneficiary_details['Business name'],
                                                        key=f"business_name_{business_name}")

                    account_number = st.text_input("Account Number",
                                                   value=str(selected_beneficiary_details['Account number']),
                                                   key=f"account_number_{business_name}", max_chars=8)

                    sort_code = st.text_input("Sort code",
                                          value=str(selected_beneficiary_details['Sort code']),
                                          key=f"sort_code_{business_name}",max_chars=6)


                    # Append the changes for this beneficiary to the list
                    changes.append({
                        'Business name': business_name,
                        'First name': first_name,
                        'Last name': last_name,
                        'Account number': account_number,
                        'Sort code': sort_code
                    })

                # Button to save changes for all selected beneficiaries
                if st.button("Save changes for selected beneficiaries"):
                    # Loop through the changes and update beneficiaries_df
                    for change in changes:
                        beneficiaries_df.loc[
                            (beneficiaries_df['Business name'] == change['Business name']),
                            ['First name', 'Last name', 'Business name', 'Account number', 'Sort code']
                        ] = [change['First name'], change['Last name'], change['Business name'],
                             change['Account number'], change['Sort code']]

                    # Update the session state with the modified dataframe
                    st.session_state['beneficiaries_df'] = beneficiaries_df

                    st.success(f"Details for the selected beneficiaries have been updated.")
                    st.write(beneficiaries_df)



    # Step 1: Identify missing bank details
    missing_info = []

    # Check if 'Account name', 'Sort code', 'Account number' are missing
    for index, row in beneficiaries_df.iterrows():
        missing_fields = []

        # Ensure that the value is treated as a string (use str() to prevent non-string issues)
        account_name = str(row['Account name']) if pd.notna(row['Account name']) else ""
        sort_code = str(row['Sort code']) if pd.notna(row['Sort code']) else ""
        account_number = str(row['Account number']) if pd.notna(row['Account number']) else ""

        # Check if fields are blank
        if account_name.strip() == "":
            missing_fields.append('Account name')

        if sort_code.strip() == "":
            missing_fields.append('Sort code')

        if account_number.strip() == "":
            missing_fields.append('Account number')

        if missing_fields:
            # Add beneficiaries with missing fields to the list
            missing_info.append((row['Business name'], missing_fields, index))

    # Step 2: Only show the expander if there are missing bank account details
    if missing_info:
        with st.expander("Missing Bank Account Information"):
            st.warning("Some beneficiaries have missing bank information.")

            # Multi-select dropdown for beneficiaries to edit (with missing bank details)
            business_names_with_missing_info = [info[0] for info in missing_info]
            selected_beneficiaries_to_edit = st.multiselect(
                "Select beneficiaries to edit bank details:",
                options=business_names_with_missing_info,
                default=[]
            )

            if selected_beneficiaries_to_edit:
                # Create a dictionary to store temporary changes for editing
                changes = {}

                # Filter the missing details for selected beneficiaries
                selected_missing_info = [info for info in missing_info if info[0] in selected_beneficiaries_to_edit]

                # Show input fields to edit bank details for selected beneficiaries
                for business_name, fields, row_index in selected_missing_info:
                    st.write(f"Edit details for {business_name}:")

                    # Initialize a dictionary for this specific beneficiary's details
                    beneficiary_changes = {'Business name': business_name}

                    # Editable fields for each selected beneficiary (display all fields, even if not missing)
                    first_name = st.text_input("First Name", value=str(beneficiaries_df.loc[row_index, 'First name']),
                                               key=f"first_name_{business_name}")
                    last_name = st.text_input("Last Name", value=str(beneficiaries_df.loc[row_index, 'Last name']),
                                              key=f"last_name_{business_name}")
                    business_name_input = st.text_input("Business Name",
                                                        value=str(beneficiaries_df.loc[row_index, 'Business name']),
                                                        key=f"business_name_{business_name}")
                    account_number = st.text_input("Account Number",
                                                   value=str(beneficiaries_df.loc[row_index, 'Account number']),
                                                   key=f"account_number_{business_name}", max_chars=8)
                    sort_code = st.text_input("Sort Code",
                                              value=str(beneficiaries_df.loc[row_index, 'Sort code']),
                                              key=f"sort_code_{business_name}", max_chars=6)

                    # Append the changes for this beneficiary to the dictionary
                    beneficiary_changes['First name'] = first_name
                    beneficiary_changes['Last name'] = last_name
                    beneficiary_changes['Business name'] = business_name_input
                    beneficiary_changes['Account number'] = account_number
                    beneficiary_changes['Sort code'] = sort_code

                    # Save the changes for this beneficiary
                    changes[business_name] = beneficiary_changes

                # Button to save changes for all selected beneficiaries
                if st.button("Save changes for selected beneficiaries", key="save_missing_bd_changes"):
                    # Apply all the changes from the 'changes' dictionary to the DataFrame
                    for business_name, change in changes.items():
                        beneficiaries_df.loc[
                            (beneficiaries_df['Business name'] == change['Business name']),
                            ['First name', 'Last name', 'Business name', 'Account number', 'Sort code']
                        ] = [change['First name'], change['Last name'], change['Business name'],
                             change['Account number'], change['Sort code']]

                    # Update the session state with the modified dataframe
                    st.session_state['beneficiaries_df'] = beneficiaries_df

                    st.success(f"Bank details for the selected beneficiaries have been updated.")
                    st.write(beneficiaries_df)

            # Multi-select dropdown for beneficiaries to delete
            selected_beneficiaries_to_delete = st.multiselect(
                "Select beneficiaries to delete:",
                options=business_names_with_missing_info,
                default=[]
            )

            if selected_beneficiaries_to_delete:
                if st.button("Delete selected beneficiaries"):
                    # Remove the rows where the 'Business name' is in the selected delete list
                    beneficiaries_df = beneficiaries_df[
                        ~beneficiaries_df['Business name'].isin(selected_beneficiaries_to_delete)]

                    # Update session state with the modified dataframe
                    st.session_state['beneficiaries_df'] = beneficiaries_df

                    # Show success message
                    st.success(f"Deleted the selected beneficiaries: {', '.join(selected_beneficiaries_to_delete)}.")

                    # Optionally display the updated DataFrame to the user
                    st.write(beneficiaries_df)




    #Notification settings
    with st.expander("Notification Settings"):

        # Set SMS
        options = ["Y", "N"]
        sms = st.selectbox("Notify SMS", options, key="sms_beneficiary")
        beneficiaries_df['Notify text'] = sms

        # Set Email
        email = st.selectbox("Notify e-mail", options, key="email_beneficiary")
        beneficiaries_df['Notify e-mail'] = email

        p_advice = st.selectbox("Notify payment advice", options, key="advice")
        beneficiaries_df['PaymentAdvice'] = p_advice

        # Save updated DataFrame to session state
        st.session_state['beneficiary_df'] = beneficiaries_df

    # EDIT BEN RECORD
    with st.expander("Edit Beneficiary Record"):
        # Drop "Unnamed" columns and "Unique ID" column
        beneficiaries_df = beneficiaries_df.loc[:, ~beneficiaries_df.columns.str.contains('^Unnamed|Unique ID')]

        # Allow the user to select multiple beneficiaries for editing
        selected_beneficiaries = st.multiselect(
            "Select beneficiaries to edit", list(beneficiaries_df['Business name'].unique()))

        if selected_beneficiaries:
            edited_values = {}  # Dictionary to store edited values for each selected beneficiary
            is_valid = True  # Start with validation set to True

            # Loop through each selected beneficiary and display editable fields
            for beneficiary_name in selected_beneficiaries:
                selected_beneficiary = beneficiaries_df[beneficiaries_df['Business name'] == beneficiary_name].iloc[0]

                # Show heading for each beneficiary
                st.subheader(f"Edit details for {beneficiary_name}")

                # Store the edited values for this beneficiary
                edited_values[beneficiary_name] = {}

                for col in beneficiaries_df.columns:  # Loop through all columns for this beneficiary
                    current_value = selected_beneficiary.get(col, "")

                    # If current_value is NaN, set it to an empty string
                    if pd.isna(current_value):
                        current_value = ""

                    # Conditionally render inputs based on column name
                    if col == "Status":
                        edited_value = st.selectbox(f"{col}", options=["Active", "Archived"],
                                                    index=["Active", "Archived"].index(
                                                        current_value) if current_value in ["Active",
                                                                                            "Archived"] else 0,
                                                    key=f'beneficiary_{beneficiary_name}_{col}')
                    elif col in ["Notify e-mail", "Notify text", "PaymentAdvice"]:
                        edited_value = st.selectbox(f"{col}", options=["Y", "N"],
                                                    index=["Y", "N"].index(current_value) if current_value in ["Y",
                                                                                                               "N"] else 0,
                                                    key=f'beneficiary_{beneficiary_name}_{col}')
                    elif col == "VAT number":
                        edited_value = st.number_input(f"{col}",
                                                       value=int(current_value) if current_value.isdigit() else 0,
                                                       step=1, key=f'beneficiary_{beneficiary_name}_{col}')
                    elif col == "Date of birth":
                        try:
                            date_value = pd.to_datetime(current_value, errors='coerce')
                            edited_value = st.date_input(f"{col}",
                                                         value=date_value if pd.notna(date_value) else None,
                                                         key=f'beneficiary_{beneficiary_name}_{col}')
                        except Exception:
                            edited_value = st.date_input(f"{col}", value=None,
                                                         key=f'beneficiary_{beneficiary_name}_{col}')
                    elif col == "Country":
                        edited_value = st.text_input(f"{col}", value=current_value, max_chars=2,
                                                     key=f'beneficiary_{beneficiary_name}_{col}')
                    elif col == "Bank country":
                        edited_value = st.text_input(f"{col}", value=current_value, max_chars=2,
                                                     key=f'beneficiary_{beneficiary_name}_{col}')
                    elif col == "Account number":
                        # Limit Account number input to 8 characters
                        edited_value = st.text_input(f"{col}", value=current_value, max_chars=8,
                                                     key=f'beneficiary_{beneficiary_name}_{col}')
                    elif col == "Sort code":
                        # Limit Sort code input to 6 characters
                        edited_value = st.text_input(f"{col}", value=current_value, max_chars=6,
                                                     key=f'beneficiary_{beneficiary_name}_{col}')
                    else:
                        edited_value = st.text_input(f"{col}", value=current_value,
                                                     key=f'beneficiary_{beneficiary_name}_{col}')

                    # Store the edited value for this beneficiary
                    edited_values[beneficiary_name][col] = edited_value

            # Validation: Check for empty fields and other conditions
            first_last_name_empty = lambda edited_values: (edited_values.get('First name', '').strip() == '') or (
                    edited_values.get('Last name', '').strip() == '')
            business_name_empty = lambda edited_values: (edited_values.get('Business name', '').strip() == '')
            country_code_invalid = lambda edited_values: len(edited_values.get('Country', '').strip()) != 2
            missing_account_name = lambda edited_values: not edited_values.get("Account name")
            missing_account_number = lambda edited_values: not edited_values.get("Account number")
            missing_sort_code = lambda edited_values: not edited_values.get("Sort code")
            sort_code_invalid = lambda edited_values: len(edited_values.get("Sort code", "").strip()) != 6
            account_number_invalid = lambda edited_values: len(edited_values.get("Account number", "").strip()) != 8

            # Validate each beneficiary's input
            for name, edited in edited_values.items():
                if first_last_name_empty(edited) and business_name_empty(edited):
                    st.error(f"Either 'First name & Last name' or 'Business name' must be provided for {name}.")
                    is_valid = False

                if country_code_invalid(edited):
                    st.error(f"Country code must have exactly 2 characters for {name}.")
                    is_valid = False

                if missing_account_name(edited):
                    st.error(f"Account name missing for {name}.")
                    is_valid = False
                if missing_account_number(edited):
                    st.error(f"Account number missing for {name}.")
                    is_valid = False
                if missing_sort_code(edited):
                    st.error(f"Sort code missing for {name}.")
                    is_valid = False

                # Check if Sort code is exactly 6 characters
                if sort_code_invalid(edited):
                    st.error(f"Sort code must be exactly 6 characters for {name}.")
                    is_valid = False

                # Check if Account number is exactly 8 characters
                if account_number_invalid(edited):
                    st.error(f"Account number must be exactly 8 characters for {name}.")
                    is_valid = False

            # Disable Save Changes button if validation fails
            save_button_enabled = is_valid and st.button("Save Changes", key='save_changes_beneficiary_records')

            # Save changes to session state when the user updates the fields
            if save_button_enabled:
                # Loop through each beneficiary and update the dataframe
                for beneficiary_name, edited in edited_values.items():
                    beneficiaries_df.loc[beneficiaries_df['Business name'] == beneficiary_name] = \
                        [edited[col] for col in beneficiaries_df.columns]

                # Save the updated DataFrame back to session state
                st.session_state['beneficiaries_df'] = beneficiaries_df
                st.success("Changes saved successfully!")

        else:
            st.warning("Please select at least one beneficiary to edit.")

    def add_leading_zeros(df):
        # Add leading zeros to 'Sort code' if less than 6 characters
        if 'Sort code' in df.columns:
            df['Sort code'] = df['Sort code'].apply(
                lambda x: str(x).zfill(6) if isinstance(x, str) and len(str(x)) < 6 else x)

        # Add leading zeros to 'Account number' if less than 8 characters
        if 'Account number' in df.columns:
            df['Account number'] = df['Account number'].apply(
                lambda x: str(x).zfill(8) if isinstance(x, str) and len(str(x)) < 8 else x)

        return df

    # Create a BytesIO buffer for download
    output = BytesIO()

    # Apply the leading zeros function to the dataframe
    beneficiaries_df = add_leading_zeros(beneficiaries_df)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Write the dataframe to the Excel file
        beneficiaries_df.to_excel(writer, index=False, sheet_name="Beneficiaries")

        # Access the xlsxwriter workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Beneficiaries"]

        # Find the index of the 'Business name' column
        business_name_col = beneficiaries_df.columns.get_loc('Business name')

        # Freeze the columns starting from 'Business name' (use the column index)
        worksheet.freeze_panes(1, business_name_col + 1)  # Freeze first row and the 'Business name' column

        # Format: Align all data from row 2 downwards to the left
        cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'border': 0})  # Remove borders

        # Apply left alignment for all cells (except the header)
        for row in range(1, len(beneficiaries_df) + 1):  # Skipping the header row (which is row 0)
            worksheet.set_row(row, None, cell_format)

        # Auto-size all columns to fit the content
        for col_num, column in enumerate(beneficiaries_df.columns.values):
            max_len = max(beneficiaries_df[column].astype(str).map(len).max(), len(column)) + 2  # Add some padding
            worksheet.set_column(col_num, col_num, max_len)

        # Set header row width to 48 pixels (approximately 6.4 characters wide, since 1 character ~= 7 pixels)
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 0,
        })

        # Add filters to the header row
        worksheet.autofilter(0, 0, 0, len(beneficiaries_df.columns) - 1)

    output.seek(0)
    xlsx_data = output.read()

    # Download button
    st.download_button(
        label="Download Beneficiaries",
        data=xlsx_data,
        file_name=f"Beneficiaries To Check And Import {agent_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )


def build_tenants():
    launch_date = st.session_state['launch_date']
    agent_name = st.session_state['agent_name']

    st.info("Tenants")
    tenants_df = st.session_state['tenants_df']
    master_df = st.session_state['master_df']

    # Remove rows that are completely empty or have blank 'Unique ID' (or other critical columns)
    tenants_df = tenants_df.dropna(subset=['Business name'], how='all')

    # Add unique codes from master_df
    tenants_df['Unique ID'] = master_df["Unique ID"]

    # Master column headers on the left
    column_mapping_tenant = {
        "First name": "First name",
        "Last name": "Last name",
        "Tenant(s) Name": "Business name",
        "Passport no / Driving licence": "Passport no / Driving licence",
        "Date of birth": "Date of birth",
        "VAT number": "VAT number",
        "Tenant E-Mail Address": "E-mail address",
        "Additional Tenant E-Mail CC": "E-mail CC",
        "Tenant Mobile Number": "Mobile number",
        "Tenant Additional Phone Number": "Phone number",
        "Tenant Fax Number": "Fax number",
        "Tenant Address 1 (Same as Property Address. Leave blank)": "Address 1",
        "Tenant Address 2": "Address 2",
        "Tenant Address 3": "Address 3",
        "Tenant City": "City",
        "Tenant County": "County",
        "Tenant Postcode": "Postcode",
        "Tenant Country": "Country",
        "Tenant Account name": "Account name",
        "Tenant Bank name": "Bank name",
        "Tenant Sort Code": "Sort code",
        "Tenant Account Number": "Account number",
        "Tenant Branch name": "Branch name",
        "Customer reference": "Tenant Customer Reference"
    }

    for master_col, tenants_col in column_mapping_tenant.items():
        # Check if the column exists in master_df
        if master_col in master_df.columns:
            try:
                # Map the values if the column exists
                mapping = master_df.set_index("Unique ID")[master_col]
                tenants_df[tenants_col] = tenants_df["Unique ID"].map(mapping)
            except KeyError:
                # Skip if "Unique ID" is missing in either DataFrame
                st.warning(f"Column '{master_col}' or 'Unique ID' not found.")

    tenants_df['Notify text'] = "Y"
    tenants_df['Notify e-mail'] = "Y"

    # Save the updated DataFrame in session state
    st.session_state['tenants_df'] = tenants_df

    # Ensure 'Country' column exists
    if 'Country' in tenants_df.columns:
        # Check if Business name, First name, and Last name are blank
        non_blank_names = ~tenants_df[['Business name', 'First name', 'Last name']].isna().any(axis=1)

        # Apply the condition to only look for missing countries if names are not blank
        missing_countries = (tenants_df['Country'].isna() | (tenants_df['Country'] == "")) & non_blank_names

        with st.expander("Missing country codes"):
            if missing_countries.any():
                st.warning("Missing country codes")
                st.dataframe(tenants_df[missing_countries])

                if st.checkbox("Apply 'UK' to all missing/blank countries", key="tenant auto country"):
                    tenants_df.loc[missing_countries, 'Country'] = "UK"
                    st.success("Default value 'UK' applied to all missing/blank countries.")
                    st.session_state['tenants_df'] = tenants_df

                elif st.checkbox("Manually add country codes", key='tenant manual country'):
                    for index in tenants_df[missing_countries].index:
                        property_name = tenants_df.at[index, 'Business name']
                        country_input = st.text_input(f"Enter country code for {property_name}:", max_chars=2,
                                                      key=f"tenant_country_input_{index}")
                        if country_input:
                            tenants_df.at[index, 'Country'] = country_input
                            st.session_state['tenants_df'] = tenants_df

    with st.expander("Tenant Settings"):
        # Set lead days
        if tenants_df['Invoice lead days'].isnull().any():
            options = [str(i) for i in range(61)]
            invoice_lead_days = st.selectbox("Select Invoice Lead Days (0-60)", options)
            tenants_df['Invoice lead days'] = invoice_lead_days

        # Set SMS
        options = ["Y", "N"]
        sms = st.selectbox("Notify SMS", options, key="sms_tenant")
        tenants_df['Notify text'] = sms

        # Set Email
        email = st.selectbox("Notify e-mail", options, key="email_tenant")
        tenants_df['Notify e-mail'] = email



    # Save updated DataFrame to session state
    st.session_state['tenants_df'] = tenants_df

    # Create a BytesIO buffer for download
    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Write the dataframe to the Excel file
        tenants_df.to_excel(writer, index=False, sheet_name="Tenants")

        # Access the xlsxwriter workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Tenants"]

        # Find the index of the 'Business name' column
        business_name_col = tenants_df.columns.get_loc('Business name')

        # Freeze the columns starting from 'Business name' (use the column index)
        worksheet.freeze_panes(1, business_name_col +1)  # Freeze first row and the 'Business name' column

        # Format: Align all data from row 2 downwards to the left
        cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'border': 0})  # Remove borders

        # Apply left alignment for all cells (except the header)
        for row in range(1, len(tenants_df) + 1):  # Skipping the header row (which is row 0)
            worksheet.set_row(row, None, cell_format)

        # Auto-size all columns to fit the content
        for col_num, column in enumerate(tenants_df.columns.values):
            max_len = max(tenants_df[column].astype(str).map(len).max(), len(column)) + 2  # Add some padding
            worksheet.set_column(col_num, col_num, max_len)

        # Set header row width to 48 pixels (approximately 6.4 characters wide, since 1 character ~= 7 pixels)
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 0,
        })

        # Add filters to the header row
        worksheet.autofilter(0, 0, 0, len(tenants_df.columns) - 1)

    output.seek(0)
    xlsx_data = output.read()

    # Download button
    st.download_button(
        label="Download Tenants",
        data=xlsx_data,
        file_name=f"Tenants To Check And Import {agent_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )






def build_invoices():

    st.info("Invoices")
    launch_date = st.session_state['launch_date']
    agent_name = st.session_state['agent_name']

    # Load the DataFrames

    invoices_df = st.session_state['invoices_df']
    master_df = st.session_state['master_df']



    # Drop all empty rows from invoices_df
    invoices_df = invoices_df.dropna(how='all').reset_index(drop=True)

    # Flatten invoices_df into a single list of strings for easier processing
    invoices_flat = invoices_df.apply(
        lambda col: col.apply(str)  # Convert all elements to strings
    ).stack().reset_index(drop=True)


    def remove_id_in_brackets(text):
        """
        Removes any ID in square brackets from the property name.
        """
        return re.sub(r'\[.*?\]', '', text).strip()


    # Function to find the full matching cell for a property name
    def find_property_details(property_name, invoices_flat):
        """
        Finds the exact match for property name excluding the ID in square brackets.
        """
        property_name_clean = remove_id_in_brackets(property_name).strip().lower()

        for cell in invoices_flat:
            cell_clean = remove_id_in_brackets(cell).strip().lower()

            # Perform exact match (no substring match)
            if property_name_clean == cell_clean:
                return cell
        return None

    def find_tenant_details(tenant_name, invoices_flat):
        """
        Finds the exact match for tenant name excluding the ID in square brackets.
        """
        tenant_name_clean = remove_id_in_brackets(tenant_name).strip().lower()

        for cell in invoices_flat:
            cell_clean = remove_id_in_brackets(cell).strip().lower()

            # Perform exact match (no substring match)
            if tenant_name_clean == cell_clean:
                return cell
        return None

    # Add property details to master_df
    master_df['property_details'] = master_df['Property Name'].apply(
        lambda x: find_property_details(x, invoices_flat) if isinstance(x, str) else None
    )

    # Add tenant details to master_df
    master_df['tenant_details'] = master_df['Tenant(s) Name'].apply(
        lambda x: find_tenant_details(x, invoices_flat) if isinstance(x, str) else None


    frequencies = []

    # Search through the DataFrame for the 'Frequency list' value
    for col in invoices_df.columns:
        for idx, value in invoices_df[col].items():
            if isinstance(value, str) and "frequency" in value.lower():
                # If 'Frequency list' is found, get the values below it in the same column
                frequencies = invoices_df[col].iloc[idx + 1:].dropna().unique().tolist()
                break
        if frequencies:
            break

    # Store frequencies list in session state for later use
    st.session_state['frequencies'] = frequencies


    categories = []

    # Search through the DataFrame for the 'Frequency list' value
    for col in invoices_df.columns:
        for idx, value in invoices_df[col].items():
            if isinstance(value, str) and "invoice type" in value.lower():
                # If 'Frequency list' is found, get the values below it in the same column
                categories = invoices_df[col].iloc[idx + 1:].dropna().unique().tolist()
                break
        if categories:
            break

    # Store frequencies list in session state for later use
    st.session_state['categories'] = categories



    # Define the mapping between master and invoices columns
    invoices_mapping = {
        'Invoices ID':'ID',
        'Invoices ExternalID':'ExternalID',
        'property_details': 'Property',
        'tenant_details': 'Tenant',
        'Rent Amount': 'Gross amount',
        'VAT On Rent': 'VAT',
        'VAT Amount On Rent': 'VAT amount',
        'Invoice Description': 'Description',
        'Add invoice period': 'Add invoice period',
        'Invoice type': 'Invoice type',
        'Rent Frequency': 'Frequency',
        'Next Due Date (If Not Monthly)': 'From date',
        'Tenancy End Date': 'To date',
        'Rent Due Date': 'Payment day',
        'Direct Debit': 'Direct Debit',
        'Invoices Customer Reference': 'Customer reference',
        'Invoice status':'Status'
    }


    # Define the column types.#'Frequency': 'Select',
    column_types = {
        'Gross amount': 'Number',
        'VAT amount': 'Number',
        'Add invoice period': 'Text',
        'Invoice type': 'Select',
        'From date': 'Date',
        'To date': 'Date',
        'Payment day': 'Number',
        'Direct Debit': 'Select'
    }

    # Clear all rows and values but keep headers
    invoices_df = invoices_df.iloc[0:0]




    # Apply transformations based on column types
    for col, col_type in column_types.items():
        if col in invoices_df.columns:
            if col_type == 'Number':
                invoices_df[col] = pd.to_numeric(invoices_df[col], errors='coerce')
                invoices_df['Gross amount'] = invoices_df['Gross amount'].round(2)

            elif col_type == 'Select':
                invoices_df[col] = invoices_df[col].astype('category')  # Convert to category (select)
            elif col_type == 'Date':
                invoices_df[col] = pd.to_datetime(invoices_df[col]).dt.strftime(
                    '%Y/%m/%d')  # Format as date (yyyy/mm/dd)
            elif col_type == 'Number':
                invoices_df[col] = invoices_df[col].astype('Int64')  # Convert to int (no decimals)
            elif col_type == 'Text':
                invoices_df[col] = invoices_df[col].astype(str)



    # Map and copy data using the defined mapping (only for columns that exist in both)
    for master_col, invoices_col in invoices_mapping.items():
        if master_col in master_df.columns and invoices_col in invoices_df.columns:
            invoices_df[invoices_col] = master_df[master_col]


    # Drop rows where all columns are NaN (None or empty)
    invoices_df = invoices_df.dropna(how='all')


    # Handle missing 'Invoice Description' in master_df
    if 'Invoice Description' not in master_df.columns:
        invoices_df['Description'] = master_df['Property Name'].apply(
                lambda property_val: f'Rent for {property_val}' if pd.notnull(property_val) else 'Rent')
    else:
            invoices_df['Description'] = 'Rent'


    if 'Frequency' in invoices_df.columns:
        invoices_df.loc[invoices_df['Frequency'] != 'Single', 'Add invoice period'] = "Y"


    if 'Invoice type' not in master_df.columns:
        invoices_df['Invoice type'] = "Rent"





    invoices_df.loc[pd.isnull(invoices_df["From date"]), "From date"] = launch_date


                # Check if 'Gross amount' is None or NaN and flag it
    missing_gross_amount = invoices_df[invoices_df['Gross amount'].isna()]

    if not missing_gross_amount.empty:
        with st.expander("Flagged rows with missing 'Gross amount'"):
            st.write("The following properties have missing 'Gross amount':")
            st.write(missing_gross_amount)

            # Provide options to the user to either input values or delete rows
            add_gross_amount = st.checkbox(
                "Add 'Gross amount' manually"
            )
            delete_rows = st.checkbox(
                "Delete rows with missing 'Gross amount'"
            )

            if add_gross_amount:
                for index, row in missing_gross_amount.iterrows():
                    property_name = row['Property']
                    gross_amount_input = st.number_input(
                        f"Enter 'Gross amount' for {property_name}",
                        value=0.0,
                        min_value=0.0,
                        key=f"add_gross_amount_manually_{property_name}"
                    )
                    if gross_amount_input != 0.0:  # If the user provides an amount
                        invoices_df.at[index, 'Gross amount'] = gross_amount_input
                        st.write("Updated 'Gross amount' values:")
                        st.write(invoices_df)

            if delete_rows:
                invoices_df = invoices_df.dropna(subset=['Gross amount'])
                st.write("Rows with missing 'Gross amount' have been deleted.")


    missing_frequency = invoices_df[invoices_df['Frequency'].isna()]

    if not missing_frequency.empty:
        with st.expander("Flagged rows with missing Frequency"):
            st.write("The following rows are missing a frequency:")
            st.write(missing_frequency)

            # Allow the user to choose between editing or deleting the row
            for index, row in missing_frequency.iterrows():
                property_name = row['Property']
                tenant_name = row['Tenant']

                # Radio button to select whether to edit or delete the row
                action = st.radio(
                    f"Choose action for {property_name} - {tenant_name}",
                    options=["Edit Frequency", "Delete Row"],
                    index=None,
                    key=f"action_{index}"  # Ensure each radio button is uniquely identifiable
                )

                if action == "Edit Frequency":
                    # Provide a dropdown for the user to select a frequency
                    selected_frequency = st.selectbox(
                        f"Select 'Frequency' for {property_name} - {tenant_name}",
                        options=["Please select"] + frequencies,  # Add 'Please select' as the first option
                        index=0,  # Set default index to 'Please select'
                        key=f"frequency_{index}"  # Ensure each selectbox is uniquely identifiable
                    )

                    # Update the 'Frequency' column in the DataFrame, only if it's not 'Please select'
                    if selected_frequency != "Please select":
                        invoices_df.at[index, 'Frequency'] = selected_frequency
                    st.success(f"Updated 'Frequency' for {property_name} - {tenant_name} to {selected_frequency}")

                elif action == "Delete Row":
                    # Button to delete this individual row

                        invoices_df.drop(index, inplace=True)
                        st.success(f"Deleted row for {property_name} - {tenant_name}")

            # Button to delete all rows with missing Frequency
            if st.button("Delete all rows with missing Frequency"):
                invoices_df.dropna(subset=["Frequency"], inplace=True)
                st.success("Deleted all rows with missing Frequency")

            #st.write("Updated 'Frequency' values:")
            #st.write(invoices_df[['Property', 'Tenant', 'Frequency']])

        missing_payment_day = invoices_df[invoices_df['Payment day'].isna()]

        if not missing_payment_day.empty:
            with st.expander("Flagged rows with missing Payment day"):
                st.write("The following rows are missing a payment day:")
                st.write(missing_payment_day)

                # Allow the user to choose between editing or deleting the row using checkboxes
                for index, row in missing_payment_day.iterrows():
                    property_name = row['Property']
                    tenant_name = row['Tenant']



                    st.write(f"Choose action for {property_name} - {tenant_name}:")

                    # Checkbox to select "Edit Payment day"
                    edit_action = st.checkbox(
                        "Edit Payment day",
                        key=f"edit_{index}"  # Ensure each checkbox is uniquely identifiable
                    )

                    # Checkbox to select "Delete Row"
                    delete_action = st.checkbox(
                        "Delete Row",
                        key=f"delete_{index}"  # Ensure each checkbox is uniquely identifiable
                    )

                    if edit_action:
                        # Provide a number input for the user to enter a payment day
                        selected_payment_day = st.number_input(
                            f"Enter payment day for {property_name} - {tenant_name}",
                            min_value=1,  # Minimum value set to 1
                            max_value=31,  # Maximum value set to 31
                            step=1,  # Step size of 1
                            key=f"payment_day_{index}"  # Ensure each input box is uniquely identifiable
                        )

                        # Check if a valid number has been entered (not None)
                        if selected_payment_day is not None and 1 <= selected_payment_day <= 31:
                            invoices_df.at[index, 'Payment day'] = selected_payment_day
                            st.success(
                                f"Updated payment day for {property_name} - {tenant_name} to {selected_payment_day}")
                        else:
                            # Don't trigger success if the input is None or outside the valid range
                            st.warning(f"Please enter a valid payment day (1-31) for {property_name} - {tenant_name}")

                    if delete_action:
                        # Button to delete this individual row
                        if st.button(f"Delete row for {property_name} - {tenant_name}", key=f"delete_button_{index}"):
                            invoices_df.drop(index, inplace=True)
                            st.success(f"Deleted row for {property_name} - {tenant_name}")

                # Button to delete all rows with missing Payment day
                if st.button("Delete all rows with missing Payment day"):
                    invoices_df.dropna(subset=["Payment day"], inplace=True)
                    st.success("Deleted all rows with missing payment day")

                #st.write("Updated 'Payment day' values:")
               # st.write(invoices_df[['Property', 'Tenant', 'Payment day']])




            # Update the session state
    st.session_state['invoices_df'] = invoices_df

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Write the dataframe to the Excel file
        invoices_df.to_excel(writer, index=False, sheet_name="Invoices")

        # Access the xlsxwriter workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Invoices"]

        # Find the index of the  column
        property_col = invoices_df.columns.get_loc('Property')

        # Freeze the columns starting from 'Business name' (use the column index)
        worksheet.freeze_panes(1, property_col + 1)  # Freeze first row and the 'Business name' column

        # Format: Align all data from row 2 downwards to the left
        cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'border': 0})  # Remove borders

        # Apply left alignment for all cells (except the header)
        for row in range(1, len(invoices_df) + 1):  # Skipping the header row (which is row 0)
            worksheet.set_row(row, None, cell_format)

        # Auto-size all columns to fit the content
        for col_num, column in enumerate(invoices_df.columns.values):
            max_len = max(invoices_df[column].astype(str).map(len).max(), len(column)) + 2  # Add some padding
            worksheet.set_column(col_num, col_num, max_len)

        # Set header row width to 48 pixels (approximately 6.4 characters wide, since 1 character ~= 7 pixels)
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 0,
        })

        # Add filters to the header row
        worksheet.autofilter(0, 0, 0, len(invoices_df.columns) - 1)

    # Move to the beginning of the BytesIO buffer
    output.seek(0)
    xlsx_data = output.read()

    # Create a download button for the Excel file
    st.download_button(
        label="Download Invoices",
        data=xlsx_data,
        file_name=f"Invoices To Check And Import {agent_name}.xlsx",  # Ensure the file has the .xlsx extension
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        use_container_width=True  # Makes the button span the full width
    )


def build_owner_payments():
    st.info('Payments')
    payments_df = st.session_state['payments_df']
    master_df = st.session_state['master_df']
    launch_date = st.session_state['launch_date']
    agent_name = st.session_state['agent_name']

    # List of substrings to exclude from processing
    exclude_keywords = ['e-mail', 'number', 'code', 'country']

    # Identify columns to include (exclude those containing the keywords)
    columns_to_process = [
        col for col in master_df.columns if not any(keyword in col.lower() for keyword in exclude_keywords)
    ]

    # Clean the selected string columns in master_df
    master_df[columns_to_process] = master_df[columns_to_process].apply(
        lambda x: x.strip().title() if isinstance(x, str) else x
    )

#change commission to percentage

    # Drop all empty rows from invoices_df
    payments_df = payments_df.dropna(how='all').reset_index(drop=True)

    # Flatten payments_df into a single list of strings for easier processing
    payments_flat = payments_df.apply(
        lambda col: col.apply(str)  # Convert all elements to strings
    ).stack().reset_index(drop=True)

    def remove_id_in_brackets(text):
        """
        Removes any ID in square brackets from the property name.
        """
        return re.sub(r'\[.*?\]', '', text).strip()

    def find_property_details(property_name, payments_flat):
        """
        Finds the exact match for property name excluding the ID in square brackets.
        """
        property_name_clean = remove_id_in_brackets(property_name).strip().lower()

        for cell in payments_flat:
            cell_clean = remove_id_in_brackets(cell).strip().lower()

            # Perform exact match (no substring match)
            if property_name_clean == cell_clean:
                return cell
        return None
    # Function to find the full matching cell for a tenant name
    def find_beneficiary_details(beneficiary_name, payments_flat):
        """
        Finds the exact match for beneficiary name excluding the ID in square brackets.
        """
        beneficiary_name_clean = remove_id_in_brackets(beneficiary_name).strip().lower()

        for cell in payments_flat:
            cell_clean = remove_id_in_brackets(cell).strip().lower()

            # Perform exact match (no substring match)
            if beneficiary_name_clean == cell_clean:
                return cell
        return None

    # Add property details to master_df
    master_df['property_details'] = master_df['Property Name'].apply(
        lambda x: find_property_details(x, payments_flat) if isinstance(x, str) else None
    )

    # Add tenant details to master_df
    master_df['beneficiary_details'] = master_df['Landlord(s) Name'].apply(
        lambda x: find_beneficiary_details(x, payments_flat) if isinstance(x, str) else None
    )

    st.session_state['master_df'] = master_df

    frequencies = []

    # Search through the DataFrame for the 'Frequency list' value
    for col in payments_df.columns:
        for idx, value in payments_df[col].items():
            if isinstance(value, str) and "frequency" in value.lower():
                # If 'Frequency list' is found, get the values below it in the same column
                frequencies = payments_df[col].iloc[idx + 1:].dropna().unique().tolist()
                break
        if frequencies:
            break

    # Store frequencies list in session state for later use
    st.session_state['frequencies'] = frequencies

    categories = []

    # Search through the DataFrame for the 'Frequency list' value
    for col in payments_df.columns:
        for idx, value in payments_df[col].items():
            if isinstance(value, str) and "Category" in value.lower():
                # If 'Frequency list' is found, get the values below it in the same column
                categories = payments_df[col].iloc[idx + 1:].dropna().unique().tolist()
                break
        if categories:
            break

    # Store frequencies list in session state for later use
    st.session_state['categories'] = categories

    use_money_from = []

    # Search through the DataFrame for the 'Frequency list' value
    for col in payments_df.columns:
        for idx, value in payments_df[col].items():
            if isinstance(value, str) and "Use money from" in value.lower():
                # If 'Frequency list' is found, get the values below it in the same column
                categories = payments_df[col].iloc[idx + 1:].dropna().unique().tolist()
                break
        if categories:
            break

    # Store frequencies list in session state for later use
    st.session_state['use_money_from'] = use_money_from

    payments_mapping = {
        'Landlord ID': 'ID',
        'Landlord Group ID': 'Group ID',
        'property_details': 'Property',
        'Landlord Category': 'Category',
        'beneficiary_details': 'Beneficiary',
        'Are The Rent Proceeds Split Between Multiple Bank Accounts? (100% Or 50/50, Etc.)': 'Gross percentage',
        'Landlord fixed amount': 'Gross amount',
        'VAT Included?': 'Vat',
        'Landlord Vat Amount': 'Vat amount',
        'Landlord Vat Percentage': 'Vat percentage',
        'Landlord Frequency': 'Frequency',
        'Landlord Payment Day': 'Payment day',
        'From date': 'From date',
        'To date': 'To date',
        'Beneficiary reference': 'Beneficiary reference',
        'Beneficiary Description': 'Description',
        'Use money from': 'Use money from',
        'No commission': 'No commission',
        'No commission amount': 'No commission amount',
        'Enabled': 'Enabled'
    }

    payments_df = pd.DataFrame(columns=[v for v in payments_mapping.values()])
    payments_df = pd.DataFrame(columns=payments_df.columns)

    for master_col, invoices_col in payments_mapping.items():
        if master_col in master_df.columns and invoices_col in payments_df.columns:
            payments_df[invoices_col] = master_df[master_col]

    payments_df['Category'] = 'Owner'
    payments_df['From date'] = launch_date

    for idx, row in payments_df.iterrows():
        if row['Category'] == 'Owner':
            if 'Landlord Description' not in master_df.columns:
                property_val = master_df.loc[idx, 'Property Name']
                payments_df.at[idx, 'Description'] = f'Rent for {property_val}' if pd.notnull(property_val) else ''

    # Collect rows with missing Property
    missing_property_rows = payments_df[payments_df['Property'].isnull()]
    with st.expander("Select Beneficiary reference"):
        payment_type = st.selectbox("Choose", ['Bulk payments', 'Individual payments'])

        if payment_type == 'Bulk payments':
            # Remove anything in parentheses and limit to 18 characters
            beneficiary_reference = agent_name.split('(')[0].strip()[:18]
            st.write(f"Beneficiary reference set to '{beneficiary_reference}'.")

            override_checkbox = st.checkbox("Override Beneficiary reference (optional)")

            if override_checkbox:
                # Limit the text input to 18 characters
                user_input = st.text_input("Enter your Beneficiary reference", value=beneficiary_reference,
                                           max_chars=18)

            # Apply to all rows in payments_df for Bulk payments
            payments_df['Beneficiary reference'] = beneficiary_reference

        elif payment_type == 'Individual payments':
            # Apply Property Name from master_df to each row in payments_df
            payments_df['Beneficiary reference'] = master_df['Property Name'].iloc[:len(payments_df)].apply(
                lambda x: x[:18])

            st.write(f"Beneficiary reference set to Property Name (capped to 18 characters).")

    st.session_state[payments_df] = payments_df

    if not missing_property_rows.empty:
        with st.expander("Rows with Missing Property"):
            st.write("Preview of rows with missing Property values:")
            st.write(missing_property_rows)

            # Add delete all checkbox for user to delete all rows at once
            delete_all = st.checkbox("Delete All Rows with Missing Property")

            if delete_all:
                # Remove all rows with missing properties from the DataFrame
                payments_df = payments_df[payments_df['Property'].notnull()]
                st.session_state['payments_df'] = payments_df
                st.write("All rows with missing properties have been deleted.")
            else:
                # Iterate over each missing row and provide options for adding a property or deleting
                for idx, row in missing_property_rows.iterrows():
                    st.write(f"Row {idx + 1}:")

                    # Allow user to select a property
                    property_options = master_df['property_details'].dropna().unique().tolist()
                    selected_property = st.selectbox(f"Select Property for row {idx + 1}", property_options,
                                                     key=f"property_{idx}")

                    # Provide a delete button for each row
                    delete_row = st.checkbox(f"Delete row {idx + 1}", key=f"delete_{idx}")

                    if delete_row:
                        # Remove the row from the DataFrame
                        payments_df = payments_df.drop(idx)

                    # Allow the user to update the property for the row if they select a property
                    payments_df.at[idx, 'Property'] = selected_property

                st.session_state['payments_owners_df'] = payments_df
                st.session_state['payments_commission_df'] = payments_df
    st.session_state['payments_owners_df'] = payments_df
    st.session_state['payments_commission_df'] = payments_df



def build_commission_payments():
    payments_owners_df = st.session_state['payments_owners_df']
    payments_df = st.session_state['payments_df']
    master_df = st.session_state['master_df']
    payments_commission_df = st.session_state['payments_commission_df']
    launch_date = st.session_state['launch_date']
    agent_name = st.session_state['agent_name']

    commission_mapping = {
        'Commission ID': 'ID',
        'Commission Group ID': 'Group ID',
        'property_details': 'Property',
        'Commission Category': 'Category',
        'predefined agent':'Beneficiary',
        'Commission/Management Percentage': 'Gross percentage',
        'Or Fixed Management Fee': 'Gross amount',
        'VAT Included?': 'Vat',
        'Commission Vat Amount': 'Vat amount',
        'Commission Vat Percentage': 'Vat percentage',
        'Commission Frequency': 'Frequency',
        'Commission Payment Day': 'Payment day',
        'From date': 'From date',
        'To date': 'To date',
        'Commission reference': 'Beneficiary reference',
        'Commission Description': 'Description',
        'Use money from': 'Use money from',
        'No commission': 'No commission',
        'No commission amount': 'No commission amount',
        'Enabled': 'Enabled'
    }

    payments_commission_df = pd.DataFrame(columns=[v for v in commission_mapping.values()])
    payments_commission_df = pd.DataFrame(columns=payments_commission_df.columns)

    for master_col, commission_col in commission_mapping.items():
        if master_col in master_df.columns and commission_col in payments_commission_df.columns:
            payments_commission_df[commission_col] = master_df[master_col]



    # Drop all empty rows from payments_df
    payments_df = payments_df.dropna(how='all').reset_index(drop=True)
    # Flatten payments_df into a single list of strings for easier processing
    payments_flat = payments_df.apply(
        lambda col: col.apply(str)  # Convert all elements to strings
    ).stack().reset_index(drop=True)
    # Find the value in payments_flat that includes '[C]' and set it as agent_predefined
    agent_predefined = None
    for value in payments_flat:
        if '[C]' in value:
            agent_predefined = value
            break



    payments_commission_df['Category'] = "Commission"
    payments_commission_df['Beneficiary'] = agent_predefined
    payments_commission_df['From date'] = launch_date

    # Link 'property_details' in master_df to 'Property' in payments_commission_df and fetch the Property Name
    for idx, row in payments_commission_df.iterrows():
        if 'Commission Description' not in master_df.columns:
            property_details_value = row['Property']  # 'Property' in payments_commission_df
            # Find the corresponding row in master_df using 'property_details'
            matched_row = master_df[master_df['property_details'] == property_details_value]
            if not matched_row.empty:
                property_name = matched_row.iloc[0]['Property Name']  # Retrieve the Property Name
                payments_commission_df.at[idx, 'Description'] = f'Commission for {property_name}' if pd.notnull(
                    property_name) else ''


    # Find rows where both 'Gross Percentage' and 'Gross Amount' are missing
    missing_rows = payments_commission_df[
        payments_commission_df['Gross percentage'].isna() & payments_commission_df['Gross amount'].isna()]

    # Show a single expander for all missing rows
    if not missing_rows.empty:
        with st.expander("Enter Gross Percentage or Gross Amount for Missing Rows"):
            # Loop through missing rows and provide input fields for each row
            for idx, row in missing_rows.iterrows():
                property_name = master_df.loc[idx, 'Property Name'] if pd.notnull(
                    master_df.loc[idx, 'Property Name']) else 'Unknown Property'
                st.write(f"Property: {property_name} (Row {idx + 1})")

                # Allow user to input values for both 'Gross Percentage' and 'Gross Amount'
                gross_percentage = st.number_input(f"Enter Gross Percentage (%) for {property_name}", min_value=0.00,
                                                   max_value=100.00, step=0.01, format="%.2f",key='add_commission_percentage')
                gross_amount = st.number_input(f"Enter Gross Amount for {property_name}", min_value=0.00, step=0.01,
                                               format="%.2f",key='add_commission_amount')

                # Update the dataframe with the user input
                if gross_percentage > 0:
                    payments_commission_df.at[idx, 'Gross percentage'] = gross_percentage
                if gross_amount > 0:
                    payments_commission_df.at[idx, 'Gross amount'] = gross_amount

    # Combine the payments_owners_df and payments_commission_df into one DataFrame
    payments_combined_df = pd.concat([payments_owners_df, payments_commission_df], ignore_index=True)

    # Ensure that all columns are aligned
    payments_combined_df = payments_combined_df[payments_owners_df.columns]

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Write the dataframe to the Excel file
        payments_combined_df.to_excel(writer, index=False, sheet_name="Payments")

        # Access the xlsxwriter workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets["Payments"]

        # Find the index of the  column
        property_col = payments_combined_df.columns.get_loc('Property')

        # Freeze the columns starting from 'Business name' (use the column index)
        worksheet.freeze_panes(1, property_col + 1)  # Freeze first row and the 'Business name' column

        # Format: Align all data from row 2 downwards to the left
        cell_format = workbook.add_format({'align': 'left', 'valign': 'top', 'border': 0})  # Remove borders

        # Apply left alignment for all cells (except the header)
        for row in range(1, len(payments_combined_df) + 1):  # Skipping the header row (which is row 0)
            worksheet.set_row(row, None, cell_format)

        # Auto-size all columns to fit the content
        for col_num, column in enumerate(payments_combined_df.columns.values):
            max_len = max(payments_combined_df[column].astype(str).map(len).max(), len(column)) + 2  # Add some padding
            worksheet.set_column(col_num, col_num, max_len)

        # Set header row width to 48 pixels (approximately 6.4 characters wide, since 1 character ~= 7 pixels)
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 0,
        })

        # Add filters to the header row
        worksheet.autofilter(0, 0, 0, len(payments_combined_df.columns) - 1)



    output.seek(0)
    xlsx_data = output.read()

    # Provide the download button
    st.download_button(
        label="Download Payments",
        data=xlsx_data,
        file_name=f"Payments To Check And Import {agent_name}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

def main():


    st.header("PayProp Data Portal")

    # Form for entering agency name and launch date
    with st.form("Enter agency name and select launch date"):
        agent_name = st.text_input("*Agent Name & (ID)")
        launch_date = st.date_input("*Launch date", value=None)

        if st.form_submit_button("Submit"):
            if not agent_name or launch_date is None:
                st.error("Please add required info")
            else:
                st.session_state.agent_name = agent_name
                st.session_state.launch_date = launch_date
                st.success("Saved")

    # Ensure that both agent_name and launch_date are populated before proceeding
    if 'agent_name' not in st.session_state or 'launch_date' not in st.session_state:
        return  # Exit if the required info is not entered

    # Proceed with file upload logic
    st.subheader("Upload files:")
    uploaded_files = st.file_uploader(
        "Upload Master data, Properties, Tenants, Beneficiaries, Invoices, and Payments files",
        type=["csv", "xlsx"], accept_multiple_files=True)

    if uploaded_files:
        properties_df, beneficiaries_df, tenants_df, invoices_df, payments_df, master_df = process_uploaded_files(uploaded_files)

        # Store the uploaded files in session state
        st.session_state.uploaded_files = uploaded_files
        st.session_state['properties_df'] = properties_df
        st.session_state['beneficiaries_df'] = beneficiaries_df
        st.session_state['tenants_df'] = tenants_df
        st.session_state['invoices_df'] = invoices_df
        st.session_state['payments_df'] = payments_df
        st.session_state['master_df'] = master_df

        if master_df is not None:
            clean_master()
        else:
            st.write("Master data is None")
        # Check if both beneficiaries_df and master_df are available
        if beneficiaries_df is not None and master_df is not None:
            build_beneficiaries()

        if tenants_df is not None and master_df is not None:
            build_tenants()

        if properties_df is not None and master_df is not None:
            build_properties()

        if invoices_df is not None and master_df is not None:
            build_invoices()

        if payments_df is not None and master_df is not None:
            build_owner_payments()
            build_commission_payments()



if __name__ == "__main__":
    main()
