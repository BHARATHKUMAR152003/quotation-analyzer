import pandas as pd

# Step 1: Read main Excel file
df = pd.read_excel(r"C:\Users\A552440\Desktop\New folder\Demo\ADVANCED CONSTRUCTION TECHNOLOGIES PVT LTD_QuoteToOrder_Weekly_2026329_1774823539930.xlsx")

# Step 2: Read remove list CSV file
remove_df = pd.read_csv(r"C:\Users\A552440\Desktop\New folder\Demo\Exclusion Data for Q 2 C.csv")

# Step 3: Clean CustomerName columns (important for matching)
df['CustomerName'] = df['CustomerName'].astype(str).str.strip().str.lower()
remove_df['CustomerName'] = remove_df['CustomerName'].astype(str).str.strip().str.lower()

# Step 4: Store rows that will be removed (for tracking)
removed_data = df[df['CustomerName'].isin(remove_df['CustomerName'])]

# Step 5: Remove matching CustomerNames
df = df[~df['CustomerName'].isin(remove_df['CustomerName'])]

# Step 6: Clean 'Converted?' column
df['Converted?'] = df['Converted?'].astype(str).str.strip().str.capitalize()

# Step 7: Group by QuotationNumber and count Yes/No
summary = df.groupby('QuotationNumber')['Converted?'].value_counts().unstack(fill_value=0)

# Step 8: Ensure both 'Yes' and 'No' columns exist
for col in ['Yes', 'No']:
    if col not in summary.columns:
        summary[col] = 0

# Step 9: Reset index and rename columns
summary = summary.reset_index()
summary = summary.rename(columns={
    'QuotationNumber': 'Unique QuotationNumber',
    'Yes': 'Converted Yes',
    'No': 'Converted No'
})

# Step 10: Add Total and Conversion %
summary['Total'] = summary['Converted Yes'] + summary['Converted No']
summary['Conversion %'] = (summary['Converted Yes'] / summary['Total']) * 100

# Step 11: Count rows where value is NOT zero
yes_count = (summary['Converted Yes'] > 0).sum()
no_count = (summary['Converted No'] > 0).sum()

print("Count of Converted Yes (excluding 0):", yes_count)
print("Count of Converted No (excluding 0):", no_count)

# Step 12: Save outputs
summary.to_excel("output.xlsx", index=False)
removed_data.to_excel("removed_customers.xlsx", index=False)

print("Filtered output generated successfully!")