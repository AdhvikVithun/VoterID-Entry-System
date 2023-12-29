import pandas as pd

# Read the data from the first sheet and the first column of the Excel file
input_file_path = "D:\\adhvik\\adh\\dad project\\4000.xlsx"
df = pd.read_excel(input_file_path, header=None, sheet_name=0, usecols=[0])

# Convert each cell to strings
df = df.applymap(str)
print("Df ")
print(df)

# Initialize a list to store dictionaries, each representing a row in the output DataFrame
output_data = []

# Initialize variables to store values for each column
ward = ""
area = ""

# Iterate through each row in the input data
for index, row in df.iterrows():
    data = row.iloc[0]  # Assuming the data is in the first (and only) column

    # Check for the presence of வார்டு and update the Ward and Area accordingly
    if "வார்டு" in data:
        ward = data.split()[-1]
        area = data.split("வார்டு")[-1].strip()
        continue

    # Extract information based on the provided rules
    parts = data.split(":", 2)
    if len(parts) == 3 and (parts[0].startswith("WTD") or parts[0].startswith("JBB")):
        serial_no, voter_info, rest = parts
        voter_id = voter_info.split()[0]
        name = voter_info.split(":")[1].split()[0]
        father_husband_name = voter_info.split(":")[2].split()[0]

        # Create a dictionary representing a row in the output DataFrame
        row_data = {
            "Ward": ward,
            "Area": area,
            "Serial No": serial_no,
            "Voter ID": voter_id,
            "Name": name,
            "Father/Husband Name": father_husband_name,
        }

        # Append the row data to the output list
        output_data.append(row_data)

# Print some intermediate information
print("Output Data:")
print(output_data)

# Create a DataFrame from the list of dictionaries
output_df = pd.DataFrame(output_data)

# Print the DataFrame
print("Output DataFrame:")
print(output_df)

# Write the transformed data to the second Excel sheet
output_file_path = "D:\\adhvik\\adh\\dad project\\100A Test.xlsx"
output_df.to_excel(output_file_path, index=False)

