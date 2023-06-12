import pandas as pd

# Read the Excel file
df = pd.read_excel('data.xlsx')

# Create a pivot table
pivot_df = df.pivot_table(values='Measure Values', index=['Material Number', 'Planning Week'], columns='Measure '
                                                                                                       'Names')

# Reset the index
pivot_df.reset_index(inplace=True)
pivot_df.fillna(0, inplace=True)

# Calculate running sum values
pivot_df['Running Sum On-site Inventory'] = pivot_df.groupby('Material Number')['Inventory Qty'].cumsum()
pivot_df['Running Sum Off-site Inventory'] = pivot_df.groupby('Material Number')['Offsite Inventory Qty'].cumsum()
pivot_df['Running Sum Open PO Quantity'] = pivot_df.groupby('Material Number')['Open Po Qty'].cumsum()
pivot_df['Running Sum Demand'] = pivot_df.groupby('Material Number')['Demand Qty'].cumsum()

# Calculate BoH and WoS
pivot_df['BoH'] = (pivot_df['Running Sum On-site Inventory'] + pivot_df['Running Sum Off-site Inventory'] + pivot_df[
    'Running Sum Open PO Quantity']) - pivot_df['Running Sum Demand']
pivot_df['WoS'] = pivot_df['BoH'] / pivot_df['Demand Qty']

# Converting the processed data to excel
pivot_df.to_excel('processed_data.xlsx')
