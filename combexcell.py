import pandas as pd
import re

# Read the data from the first sheet and the first column of the Excel file
input_file_path = "D:\\adhvik\\adh\\dad project\\4000.xlsx"
df = pd.read_excel(input_file_path, header=None, sheet_name=0, usecols=[0])

# Convert each cell to strings
df = df.applymap(str)

# Initialize a list to store dictionaries, each representing a row in the output DataFrame
output_data = []

# Iterate through each row in the input data
for index, row in df.iterrows():
    data = row.iloc[0]  # Assuming the data is in the first (and only) column

    # Extract serial number if present
    serial_match = re.search(r'^\s*(\d+)\s*\.', data)
    if serial_match:
        serial_no = serial_match.group(1).strip()

        # Extract Voter ID if present
        voter_id_match = re.search(r'\b([A-Z]+\d+)\b', data)
        voter_id = voter_id_match.group(1).strip() if voter_id_match else ""

        # Extract Name if present, excluding words "கணவர்" or "தந்தை"
        name_match = re.search(rf'{voter_id}\s*:\s*((?:(?!கணவர்|தந்தை).)+)\s*', data)
        name = name_match.group(1).strip() if name_match else ""

        # Extract Father/Husband Name if present after "கணவர்" or "தந்தை"
        father_husband_name_match = re.search(r'(?:கணவர்|தந்தை)\s*:\s*(.+?)\s*(?:(?:\d+\s*\.\s*)|$)', data)
        father_husband_name = father_husband_name_match.group(1).strip() if father_husband_name_match else ""

        # Create a dictionary representing a row in the output DataFrame
        row_data = {
            "Ward": "",  # Add appropriate values for these columns
            "Election": "",
            "Booth": "",
            "Voter ID": voter_id,
            "Serial No": serial_no,
            "Part No": "",
            "Page No": "",
            "Male/Female": "",
            "Age": "",
            "Phone Num": "",
            "Aadhar Num": "",
            "Name": name,
            "Father/Husband Name": father_husband_name,
            "Door Num": "",
            "Address": "",
            "Family": "",
            "Caste": "",
            "Area Guide": "",
            "Guide Number": "",
            "Event": "",
            "Memo": "",
            "Qualification": "",
            "Stay": "",
            "Voted": "",
            "Voted Date": "",
        }

        # Append the row data to the output list
        output_data.append(row_data)

# Create a DataFrame from the list of dictionaries
output_df = pd.DataFrame(output_data)

# Reorder columns to match the specified order
column_order = ["Ward", "Election", "Booth", "Voter ID", "Serial No", "Part No", "Page No", 
                "Male/Female", "Age", "Phone Num", "Aadhar Num", "Name", "Father/Husband Name",
                "Door Num", "Address", "Family", "Caste", "Area Guide", "Guide Number", 
                "Event", "Memo", "Qualification", "Stay", "Voted", "Voted Date"]

output_df = output_df[column_order]

# Write the transformed data to the second Excel sheet
output_file_path = "D:\\adhvik\\adh\\dad project\\100A.xlsx"
output_df.to_excel(output_file_path, index=False)
