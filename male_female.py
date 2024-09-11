import os
import pandas as pd
import re
import logging

# Configure logging
log_file_path = 'male_female.log'
logging.basicConfig(filename=log_file_path, level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Check if the log file is created and writable
if not os.access(log_file_path, os.W_OK):
    print(f"Log file {log_file_path} is not writable. Check file permissions.")
else:
    print(f"Logging to {log_file_path}")

# Import the Excel file
file_path = r"C:\Users\Tevin\Links\test_files.xlsx"

# Method to generate separate lists of male and female students
def generate_gender_lists(file_path):
    try:
        # Read all sheets from the Excel file
        sheets = pd.read_excel(file_path, sheet_name=None)
    except FileNotFoundError:
        logging.error(f"File not found: {file_path}")
        return [], [], []
    except Exception as e:
        logging.error(f"An error occurred while reading the Excel file: {e}")
        return [], [], []

    male_students = []
    female_students = []
    special_char_names = []

    # Iterate through each sheet
    for sheet_name, df in sheets.items():
        # Check if the DataFrame has 'Student Name' and 'Gender' columns
        if 'Student Name' not in df.columns or 'Gender' not in df.columns:
            logging.error(f"The sheet '{sheet_name}' must contain 'Student Name' and 'Gender' columns.")
            continue
        
        # Ensure case insensitivity and strip whitespace
        df['Gender'] = df['Gender'].str.lower().str.strip()

        
        # Debug: Print unique values in the 'Gender' column
        unique_genders = df['Gender'].unique()
        logging.info(f"Unique gender values in sheet '{sheet_name}': {unique_genders}")
        
        # Separate male and female students
        male_students = df[df['Gender'] == 'm']['Student Name'].tolist()
        female_students = df[df['Gender'] == 'f']['Student Name'].tolist()
        
        # Debug: Log the number of males and females found in the current sheet
        logging.info(f"Number of males in sheet '{sheet_name}': {len(male_students)}")
        logging.info(f"Number of females in sheet '{sheet_name}': {len(female_students)}")
        
       
        
        # Find names with special characters
        for name in df['Student Name']:
            if re.search(r"[^a-zA-Z\s]", name):
                special_char_names.append(name)
    
    # Log the number of male and female students
    logging.info(f"Total number of male students: {len(male_students)}")
    logging.info(f"Total number of female students: {len(female_students)}")
    
    # Log names with special characters
    if special_char_names:
        logging.info(f"Names with special characters: {', '.join(special_char_names)}")
    else:
        logging.info("No names with special characters found.")
    
    print("Male Students:", male_students)
    print("Female Students:", female_students)
    print("Special Character Names:", special_char_names)
   


# Example usage
generate_gender_lists(file_path)

#Ensure that the log file is closed after logging
logging.shutdown()

