import pandas as pd

# List of Excel files to process
excel_files = ['./invoice_1.xlsx', './invoice_2.xlsx', './invoice_3.xlsx']

# Store grouped data from all files
all_grouped_data = []

unique_account_names = set()
unique_InvoiceNumbers = set()

for file_path in excel_files:
    # Load all sheets from each file
    xls = pd.ExcelFile(file_path)
    
    for sheet_name in xls.sheet_names:
        # Read each sheet with header on row 2 (index 1)
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)

        # Get columns up to 'Ext Price'
        cols_to_use = ['Account Number', 'AccountName', 'Frequency', 'Invoice Number', 'Invoice Date', 'Badge Type', 'Quantity', 'Unit Price', 'Ext Price']
        df_subset = df[cols_to_use]

        # Handle missing Frequency values
        df_subset['Frequency'] = df_subset['Frequency'].fillna('Missing')
        df_subset['Badge Type'] = df_subset['Badge Type'].fillna('Fees')

        # Convert to datetime in case of mixed types
        df_subset['Invoice Date'] = pd.to_datetime(df_subset['Invoice Date'], errors='coerce').dt.strftime('%d %B %Y')


        # Group with aggregation that includes non-numeric fields
        grouped = df_subset.groupby(['Account Number'], as_index=False).agg({
            'Quantity': 'sum',
            'Unit Price': 'sum',
            'Ext Price': 'sum',
            'Invoice Date': 'first',
            'Invoice Number': 'first'
        })

        # Merge back AccountName
        account_names = df[['Account Number', 'AccountName']].drop_duplicates()
        grouped = grouped.merge(account_names, on='Account Number', how='left')

        all_grouped_data.append(grouped)

# Combine all grouped data into one DataFrame
combined = pd.concat(all_grouped_data, ignore_index=True)



# Final grouping with full aggregation
final_grouped = combined.groupby(['Account Number', 'AccountName'], as_index=False).agg({
    
    'Invoice Number': 'first',
    'Invoice Date': 'first',
    'Quantity': 'sum',
    'Unit Price': 'sum',  # Or sum if you prefer
    'Ext Price': 'sum'
})


unique_account_names.update(df['AccountName'].dropna().unique())
unique_InvoiceNumbers.update(df['Invoice Number'].dropna().unique())

# Add the total row
total_row = {
    'AccountName': f"TOTAL | Unique Accounts: {len(unique_account_names)}",
    'Account Number': '',
    
    'Invoice Number': '',
    'Invoice Date': '',
    
    'Quantity': final_grouped['Quantity'].sum(),
    'Unit Price': '',
    'Ext Price': final_grouped['Ext Price'].sum()
}


# Append total row
final_grouped_with_total = pd.concat([final_grouped, pd.DataFrame([total_row])], ignore_index=True)


# Reorder columns: AccountName first
desired_order = ['AccountName', 'Account Number', 'Invoice Number', 'Invoice Date', 'Quantity', 'Unit Price', 'Ext Price']

final_grouped_with_total = final_grouped_with_total[desired_order]

# Show result
print("Final Combined Result from All Files:")
print(final_grouped_with_total)

# output_path = './final_summary.xlsx'

# # Export the DataFrame to Excel
# final_grouped_with_total.to_excel(output_path, index=False)

# print(f"✅ Excel file saved as: {output_path}")
