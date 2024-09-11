import pandas as pd
import json
import re

# Function to read and process the Excel file
def read_excel(file_path):
    sheets = pd.read_excel(file_path, sheet_name=None)
    all_data = []
    for sheet_name, df in sheets.items():
        if 'Student Name' in df.columns and 'Gender' in df.columns:
            df['Gender'] = df['Gender'].str.lower().str.strip().replace({'m': 'male', 'f': 'female'})
            df['Special Character'] = df['Student Name'].apply(lambda x: 'yes' if re.search(r"[^a-zA-Z\s]", x) else 'no')
            all_data.append(df)
    return pd.concat(all_data, ignore_index=True)

# Function to read the JSON file
def read_json(file_path):
    with open(file_path, 'r') as f:
        return json.load(f)

# Function to merge and format the data
def merge_and_format_data(excel_data, json_data):
    formatted_data = []
    for idx, row in excel_data.iterrows():
        student_name = row['Student Name']
        similar_names = [entry for entry in json_data if entry['male_name'] == student_name or entry['female_name'] == student_name]
        name_similar = 'yes' if similar_names else 'no'
        dob = row.get('DoB', 'N/A')
        if isinstance(dob, pd.Timestamp):
            dob = dob.strftime('%Y-%m-%d')  # Convert Timestamp to string
        formatted_data.append({
            "id": str(idx),
            "student_number": row.get('Student Number', 'N/A'),
            "additional_details": {
                "dob": dob,
                "gender": row['Gender'],
                "special_character": row['Special Character'],
                "name_similar": name_similar
            }
        })
    return formatted_data

# Paths to the files
excel_file_path = r"C:\Users\Tevin\Links\test_files.xlsx"
json_file_path = 'similarity_results.json'

# Read and process the files
excel_data = read_excel(excel_file_path)
json_data = read_json(json_file_path)

#Debud: print the first few rows of the DoB column
print(excel_data['DoB'].head())

# Merge and format the data
formatted_data = merge_and_format_data(excel_data, json_data)

# Save the formatted data as a JSON file
with open('merged_data.json', 'w') as f:
    json.dump(formatted_data, f, indent=4)

print("Merged data saved to merged_data.json")