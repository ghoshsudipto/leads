import pandas as pd
import os

# Set the folder path
folder_path = 'C:/Users/Ghosh/Dropbox/lead'

# Load the 'SPEQ CRM Leads.csv' file
file_path_speq = os.path.join(folder_path, 'SPEQ CRM  Leads.csv')
df_speq = pd.read_csv(file_path_speq)

# Load the 'combined.csv' file
file_path_combined = os.path.join(folder_path, 'combined.csv')
df_combined = pd.read_csv(file_path_combined)

# Compare the 'Mobile' column in SPEQ CRM Leads with 'Phone number' in combined.csv
df_missing = df_combined[~df_combined['Phone number'].isin(df_speq['Mobile'])]

# Select the relevant columns
df_to_append = df_missing[['Name', 'Phone number', 'Email address', 'Interested segment', 'City']]

# Load existing 'dropped_rows.xlsx' file if it exists, otherwise create a new one
dropped_rows_path = os.path.join(folder_path, 'dropped_rows.xlsx')

if os.path.exists(dropped_rows_path):
    df_dropped = pd.read_excel(dropped_rows_path)
    df_dropped = pd.concat([df_dropped, df_to_append], ignore_index=True)
else:
    df_dropped = df_to_append

# Save the updated 'dropped_row.xlsx' file
df_dropped.to_excel(dropped_rows_path, index=False)

print(f"Data appended to {dropped_rows_path}")

# Existing functionality: Define BDE allocation percentages
bde_allocation = {
    'TN': {'vaishalini@speqresearch.com': 90,
           'priyanka@speqresearch.com': 10},

    'WB': {'roy@speqresearch.com': 40,
           'rahul@speqresearch.com': 40,
           'Abhishek@speqresearch.com': 0,
           'ali@speqresearch.com': 10,
           'pinki@speqresearch.com': 10},

    'AS': {'jeemoni@speqresearch.com': 70,
           'roy@speqresearch.com': 25},

    'GJ': {'devanshu@speqresearch.com': 30,
           'rahul@speqresearch.com': 20,
           'pinki@speqresearch.com': 10,
           'Sahil@speqresearch.com': 10,
           'jaya@speqresearch.com': 10,
           'shubham@speqresearch.com': 5,
           'nikhil@speqresearch.com': 5,
           'Mohammad@speqresearch.com': 0,
           'priyanka@speqresearch.com': 10},

    'MH': {'Sahil@speqresearch.com': 15,
           'jaya@speqresearch.com': 15,
           'nikhil@speqresearch.com': 15,
           'jeemoni@speqresearch.com': 10,
           'shubham@speqresearch.com': 10,
           'pinki@speqresearch.com': 5,
           'manzar@speqresearch.com': 10,
           'Mohammad@speqresearch.com': 0,
           'priyanka@speqresearch.com': 20},

    'OD': {'ali@speqresearch.com': 70,
           'manzar@speqresearch.com': 30},

    'IN': {'priyanka@speqresearch.com': 20,
           'Sahil@speqresearch.com': 20,
           'jaya@speqresearch.com': 20,
           'nikhil@speqresearch.com': 15,
           'rahul@speqresearch.com': 0,
           'shubham@speqresearch.com': 10,
           'Abhishek@speqresearch.com': 0,
           'jeemoni@speqresearch.com': 15,
           'Mohammad@speqresearch.com': 0,
           'ali@speqresearch.com': 0,
           'devanshu@speqresearch.com': 0},
}

# Initialize a list to hold the allocation results
allocations = []

# Allocate BDEs based on region
for region, bdes in bde_allocation.items():
    region_df = df_speq[df_speq['City'] == region]
    total_rows = len(region_df)

    if total_rows > 0:
        start_index = 0
        allocated_so_far = 0

        # Calculate allocation for each BDE
        for bde, percentage in bdes.items():
            if percentage > 0:
                allocate_rows = int((percentage / 100) * total_rows)
                allocated_so_far += allocate_rows

                end_index = start_index + allocate_rows
                allocated_df = region_df.iloc[start_index:end_index].copy()
                allocated_df["BDE email"] = bde
                allocations.append(allocated_df)

                start_index = end_index

        # Handle any remaining rows due to rounding errors
        remaining_rows = total_rows - allocated_so_far
        if remaining_rows > 0:
            remaining_df = region_df.iloc[start_index:start_index + remaining_rows].copy()
            # Allocate remaining rows to the BDE with the highest percentage
            primary_bde = max(bdes, key=bdes.get)
            remaining_df["BDE email"] = primary_bde
            allocations.append(remaining_df)

# Concatenate all allocated DataFrames
allocation_df = pd.concat(allocations, ignore_index=True)

# Drop the specified columns
columns_to_drop = ['S.No', 'Lead ID', 'Owner', 'Status', 'Actions', 'Description', 'Modified', 'LeadSource',
                   'LeadResponse', 'Paid']
allocation_df.drop(columns=columns_to_drop, inplace=True)
allocation_df["Interested segment"] = ""
allocation_df["Phone number"] = allocation_df["Mobile"]
allocation_df["Email address"] = allocation_df["Email"]
allocation_df = allocation_df[['Name', 'Phone number', 'Email address', 'Interested segment', 'City', "BDE email"]]
# Save the allocation DataFrame to a CSV file
output_path = os.path.join(folder_path, 'bde_allocation.csv')
allocation_df.to_csv(output_path, index=False)

print(f"BDE allocation saved to: {output_path}")

