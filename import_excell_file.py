import pandas as pd
import re
import logging



#configure the logging
logging.basicConfig(filename='student_emails.log', level=logging.INFO, 
                    format='%(asctime)s - %(levelname)s - %(message)s')

# import the excel file
file_path = r"C:\Users\Tevin\Links\test_files.xlsx"
 

# method to generate unique emails
def generate_unique_emails(file_path):
    #read the excel file
    sheets = pd.read_excel(file_path, sheet_name=None)
    
    #track the unique emails
    unique_emails = set()

    #store the results
    results = []
    
    #get the sheets
    for sheet_name, df in sheets.items():
        # Check if the DataFrame has a 'Student Name' column
        if 'Student Name' not in df.columns:
            logging.error(f"The sheet '{sheet_name}' must contain a 'Student Name' column.")
            continue

        #iterate through the rows
        for name in df ['Student Name']: 
            #remove special characters 
            clean_name = re.sub(r'[^a-zA-Z\s]', '', name)
            name_parts = clean_name.split()

            #construct the email address
            if len(name_parts) == 1:
                email = f"{name_parts[0].lower()}@gmail.com"
            else:
                email = f"{name_parts[0][0].lower()}{name_parts[-1].lower()}@gmail.com"

            #check if the email is unique
            original_email = email
            counter = 1
            while email in unique_emails:
                email = f"{original_email.split('@')[0]}{counter}@gmail.com"
                counter += 1


            #add the email to the unique emails
            unique_emails.add(email)

            #add the email to the results
            results.append(f"Student Name: {name}, Email address: {email}")
            logging.info(f"Generated email for {name}: {email}")

    return results


#example
emails = generate_unique_emails(file_path)
for email in emails:
    print(email)
    logging.info(email)
    
    
