import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt

# Function to process data and create the table and exceptions
def process_data(file):
    data = pd.read_excel(file, sheet_name='contacts')

    # Extract relevant columns
    relevant_data = data[['Name', 'Licence', 'Expected_Renewal', 'Expected_Revenue', 'LicenceChange', 'RenewalStatus']]
    
    # Convert 'Expected_Renewal' to datetime
    relevant_data['Expected_Renewal'] = pd.to_datetime(relevant_data['Expected_Renewal'], errors='coerce')
    
    # Filter out rows where Expected_Renewal is NaT
    relevant_data = relevant_data.dropna(subset=['Expected_Renewal'])
    
    # Create a new column with year-month format
    relevant_data['YearMonth'] = relevant_data['Expected_Renewal'].dt.strftime('%Y-%m')
    
    # Create a pivot table with year-month as columns and names as rows, including 'Licence'
    pivot_table = relevant_data.pivot_table(index=['Name', 'Licence'], columns='YearMonth', values='Expected_Revenue', aggfunc='sum', fill_value=0).reset_index()
    
    # Additional data columns to be merged
    additional_columns = data[['Name', 'LicenceChange', 'RenewalStatus', 'Updated']].drop_duplicates()
    
    # Merge the additional columns into the pivot table
    merged_data = pd.merge(pivot_table, additional_columns, on='Name', how='left')
    
    # Reorder columns to move 'LicenceChange' and 'RenewalStatus' next to 'Name'
    cols = ['Name', 'LicenceChange', 'RenewalStatus', 'Updated'] + [col for col in merged_data.columns if col not in ['Name', 'LicenceChange', 'RenewalStatus','Updated']]
    merged_data = merged_data[cols]
    
    # Create a dataframe to identify exceptions
    exceptions = data.copy()
    
    # Convert dates to datetime for comparison
    exceptions['Expected_Renewal'] = pd.to_datetime(exceptions['Expected_Renewal'], errors='coerce')
    exceptions['renewal_date'] = pd.to_datetime(exceptions['renewal_date'], errors='coerce')
    
    # Identify missing expected renewal dates
    exceptions['Missing Expected Date'] = exceptions['Expected_Renewal'].isna()
    
    # Identify late renewals where expected_renewal date is later than the renewal_date
    exceptions['Late renewal'] = (exceptions['Expected_Renewal'] > exceptions['renewal_date'])
    
    # Filter the relevant columns for display
    exceptions_output = exceptions[['Name', 'Licence', 'Expected_Renewal', 'renewal_date', 'Missing Expected Date', 'Late renewal']]
    
    return merged_data, exceptions_output

# Streamlit app
st.title('Renewals Data Analysis')

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    merged_data, exceptions_output = process_data(uploaded_file)
    
    st.header("Relevant Data")
    st.dataframe(merged_data)
    
    st.header("Exceptions")
    st.dataframe(exceptions_output)
    
    # Create a bar chart for the new data
    st.header("Bar Chart: Expected Revenue by Year-Month")
    merged_data_melted = merged_data.melt(id_vars=['Name', 'LicenceChange', 'RenewalStatus', 'Licence'], var_name='YearMonth', value_name='Expected_Revenue')
    merged_data_melted = merged_data_melted[merged_data_melted['YearMonth'].str.match(r'\d{4}-\d{2}')]

    pivot_table_sum = merged_data_melted.groupby('YearMonth')['Expected_Revenue'].sum().reset_index()
    pivot_table_sum.columns = ['YearMonth', 'Expected_Revenue']
    
    fig, ax = plt.subplots()
    ax.bar(pivot_table_sum['YearMonth'], pivot_table_sum['Expected_Revenue'])
    ax.set_xlabel('Year-Month')
    ax.set_ylabel('Expected Revenue')
    ax.set_title('Expected Revenue by Year-Month')
    st.pyplot(fig)
    
    # Function to convert dataframe to Excel and provide a download link
    def to_excel(df):
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        df.to_excel(writer, index=True, sheet_name='Sheet1')
        writer.close()
        processed_data = output.getvalue()
        return processed_data
    
    # Provide download links
    merged_data_excel = to_excel(merged_data)
    exceptions_output_excel = to_excel(exceptions_output)
    
    st.download_button(label="Download Relevant Data as Excel", data=merged_data_excel, file_name='relevant_data.xlsx')
    st.download_button(label="Download Exceptions as Excel", data=exceptions_output_excel, file_name='exceptions.xlsx')
