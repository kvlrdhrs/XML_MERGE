import os
import pandas as pd

# Path to the folder containing XLSX files
folder_path = r'C:\Users\kyana\Desktop\test\test'

# List of all XLSX files in the folder
file_list = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

# Create an empty DataFrame to store the merged data
merged_data = pd.DataFrame()

# Specified column structure
desired_columns = ['AccountNum', 'ItemNum', 'ItemNumDF', 'OperationTime', 'OperationDate', 'PaymentName', 'PopMSKTime',
                   'ItemNumBlock', 'DebitAmount', 'CreditAmount', 'DocType', 'DocDate', 'DocNum',
                   'BPName', 'BPBIK', 'BPINN', 'BPKPP', 'PPName', 'PPAccountNum',
                   'CorrespondentAccount', 'Position', 'FirstName', 'MiddleName', 'LastName']

# Iterate through each XLSX file
for file_name in file_list:
    file_path = os.path.join(folder_path, file_name)

    # Read the XLSX file and get data from the fourth row
    df = pd.read_excel(file_path, header=None, skiprows=3)

    # Check for existing columns in the file
    existing_columns = set(df.iloc[0].tolist())
    columns_to_include = list(set(desired_columns) & existing_columns)

    # Check if there's at least one column to include
    if columns_to_include:
        # Reorder and include only existing columns
        df = df[columns_to_include]

        # Merge data using concat method
        merged_data = pd.concat([merged_data, df], ignore_index=True)

# Check the size of the merged data
if merged_data.shape[0] > 900000:
    # Split data into chunks with a limit of 900000 rows
    num_chunks = merged_data.shape[0] // 900000 + 1

    for i in range(num_chunks):
        start_idx = i * 900000
        end_idx = (i + 1) * 900000

        chunk_df = merged_data[start_idx:end_idx]

        # Create a new XLSX file based on the chunk of data
        output_file_name = f"merge{i+1}.xlsx"
        output_file_path = os.path.join(folder_path, output_file_name)
        chunk_df.to_excel(output_file_path, index=False)
else:
    # Create a new XLSX file based on the merged data
    output_file_path = os.path.join(folder_path, "merged.xlsx")
    merged_data.to_excel(output_file_path, index=False)
