import os
import pandas as pd
import re
import numpy as np
from unidecode import unidecode

# Define the folder path
folder_path = r'C:\Users\Ghosh\Dropbox\lead'

# Define the expected column names
expected_columns = ['Created', 'Name', 'Email address', 'Source', 'Form', 'Channel', 'Stage', 'Owner', 'Labels',
                    'Phone number', 'Secondary Phone Number']

# Step 1: Read all CSV files and create DataFrames
dfs = []
for filename in os.listdir(folder_path):
    if filename.endswith('.csv'):
        file_path = os.path.join(folder_path, filename)
        df = pd.read_csv(file_path)

        # Check if the number of columns matches the expected
        if len(df.columns) == len(expected_columns):
            # Step 2: Rename columns
            df.columns = expected_columns
        else:
            print(f"Column mismatch in file: {filename}, skipping this file.")
            continue

        dfs.append(df)

# Step 3: Remove blank spaces from "Email address" and "Phone number"
for df in dfs:
    df['Email address'] = df['Email address'].str.strip()

    # Convert the 'Phone number' to string and handle NaN values
    df['Phone number'] = df['Phone number'].astype(str).replace('nan', '')
    df['Phone number'] = df['Phone number'].str.replace(" ", "").str.strip()

    # Convert the 'Secondary Phone Number' to string and handle NaN values
    df['Secondary Phone Number'] = df['Secondary Phone Number'].astype(str).replace('nan', '')
    df['Secondary Phone Number'] = df['Secondary Phone Number'].str.replace(" ", "").str.strip()


# Step 4: Convert phone numbers to string format and avoid scientific notation
for df in dfs:
    df['Phone number'] = df['Phone number'].apply(lambda x: "'" + str(x).replace('.0', ''))

    df['Secondary Phone Number'] = df['Secondary Phone Number'].apply(lambda x: "'" + str(x).replace('.0', ''))

# Step 5: Create a new column "Final number" with the last 10 characters from "Phone number"
for df in dfs:
    df['Final number'] = df['Phone number'].apply(lambda x: x[-10:] if len(x) >= 10 else x)


# Step 6: Format the "Name" column
def format_name(row):
    # Step 1: Check if "Name" is empty, if so, use the part of the "Email address" before "@"
    if pd.isna(row['Name']) or row['Name'].strip() == '':
        row['Name'] = row['Email address'].split('@')[0]

    # Step 2: Convert to plain English using unidecode
    plain_name = unidecode(row['Name'])

    # Keep only letters, periods, and spaces
    plain_name = re.sub(r'[^a-zA-Z\s]', '', plain_name)

    # Capitalize first letter of each word and lower the rest
    plain_name = plain_name.title()
    return plain_name


# Step 7: Create a new column "City" based on the "Form" column
def get_city(segment):
    segment = segment.lower() if pd.notna(segment) else ''
    if any(keyword in segment for keyword in ['assam', 'asam', 'asaam']):
        return 'AS'
    elif 'wb' in segment or 'bengali' in segment:
        return 'WB'
    elif 'gujarat' in segment or 'gujrat' in segment:
        return 'GJ'
    elif 'tamil' in segment:
        return 'TN'
    elif any(keyword in segment for keyword in ['marathi', 'maharashtra', 'maratha']):
        return 'MH'
    elif 'odia' in segment:
        return 'OD'
    elif 'hindi' in segment or 'english' in segment:
        return 'IN'
    elif any(keyword in segment for keyword in ['punjabi', 'pb']):
        return 'PB'
    elif any(keyword in segment for keyword in ['telagu', 'andhra', 'hyd']):
        return 'TS'
    else:
        return np.nan


# Apply the format_name function to the "Name" column and get_city function for "City"
for df in dfs:
    df['Name'] = df.apply(format_name, axis=1)
    df['City'] = df['Form'].apply(get_city)

# Step 8: Join all DataFrames
if dfs:
    combined_df = pd.concat(dfs, ignore_index=True)

    # Step 9: Drop the specified columns
    columns_to_drop = ['Created', 'Source', 'Channel', 'Stage', 'Owner', 'Labels', 'Phone number',
                       'Secondary Phone Number']
    combined_df = combined_df.drop(columns=columns_to_drop, errors='ignore')

    # Step 10: Initialize dropped_rows DataFrame
    dropped_rows = pd.DataFrame(columns=combined_df.columns)

    # Step 11: Find duplicate rows in "Final number" and move all duplicates (including the first occurrence) to dropped_rows
    duplicates = combined_df[combined_df.duplicated(subset=['Final number'], keep=False)]
    dropped_rows = pd.concat([dropped_rows, duplicates])

    # Remove duplicates from combined_df, but keep one occurrence of each unique value
    combined_df = combined_df.drop_duplicates(subset=['Final number'], keep='first')

    # Step 12: Find rows with non-numeric characters in "Final number"
    non_numeric_rows = combined_df[~combined_df['Final number'].str.isnumeric()]
    dropped_rows = pd.concat([dropped_rows, non_numeric_rows])

    # Remove non-numeric rows from combined_df
    combined_df = combined_df[combined_df['Final number'].str.isnumeric()]

    # Step 13: Find rows where "Final number" starts with 0, 1, 2, 3, 4, 5
    invalid_start_rows = combined_df[combined_df['Final number'].str.startswith(tuple('012345'))]
    dropped_rows = pd.concat([dropped_rows, invalid_start_rows])

    # Remove rows starting with 0-5 from combined_df
    combined_df = combined_df[~combined_df['Final number'].str.startswith(tuple('012345'))]

    # Step 14: Rename "Final number" to "Phone number" and "Form" to "Interested segment"
    combined_df = combined_df.rename(columns={
        'Final number': 'Phone number',
        'Form': 'Interested segment'
    })

    combined_df["BDE email"] = ""

    # Step 15: Rearrange the columns to the desired order
    combined_df = combined_df[['Name', 'Phone number', 'Email address', 'Interested segment', 'City', "BDE email"]]

    # Step 16: Save the combined DataFrame and dropped_rows
    output_file = os.path.join(folder_path, 'combined.csv')
    dropped_file = os.path.join(folder_path, 'dropped_rows.xlsx')  # Save dropped_rows as xlsx

    combined_df.to_csv(output_file, index=False)

    # Save dropped_rows as xlsx
    dropped_rows.to_excel(dropped_file, index=False, engine='openpyxl')

    print(f'Combined CSV saved at {output_file}')
    print(f'Dropped rows XLSX saved at {dropped_file}')
else:
    print("No valid files to process.")
