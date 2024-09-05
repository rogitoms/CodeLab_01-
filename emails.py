import pandas as pd
import re


def clean_name(name):
    """Cleans the name by removing special characters and trimming spaces."""
    name = re.sub(r'[^a-zA-Z\s]', '', name)  # Remove non-alphabetic characters
    return name.strip()


def generate_email(name):
    """Generates an email address based on the student's name."""
    # Split the name into parts
    name_parts = clean_name(name).split()

    if len(name_parts) == 2:
        # Use the first initial and full last name for two-part names
        email = f"{name_parts[0][0]}{name_parts[1]}@gmail.com".lower()
    elif len(name_parts) >= 3:
        # Use the first initial and the last part of the name for three-part names
        email = f"{name_parts[0][0]}{name_parts[-1]}@gmail.com".lower()
    else:
        # If there's only one name, use it as the email
        email = f"{name_parts[0]}@gmail.com".lower()

    return email


def ensure_unique_email(email, existing_emails):
    """Ensures the email is unique by appending a number if needed."""
    base_email = email
    count = 1
    while email in existing_emails:
        email = f"{base_email.split('@')[0]}{count}@gmail.com"
        count += 1
    existing_emails.add(email)
    return email


def process_sheet(sheet_name, df, existing_emails):
    """Generates emails for a specific sheet and ensures uniqueness across all sheets."""
    emails = []
    # Loop through each student name and generate an email
    for name in df['Student Name']:
        email = generate_email(name)
        unique_email = ensure_unique_email(email, existing_emails)
        emails.append(unique_email)
    # Add the emails to the DataFrame
    df['Email Address'] = emails
    return df


def generate_emails_from_excel(file_path):
    """Reads student names from an Excel file with multiple sheets (File_A and File_B)
    and generates unique email addresses."""

    # Load both sheets into DataFrames
    sheet_names = ['File_A', 'File_B']
    xls = pd.ExcelFile(file_path)

    if not all(sheet in xls.sheet_names for sheet in sheet_names):
        print("The Excel file must contain 'File_A' and 'File_B' sheets.")
        return

    existing_emails = set()  # Keep track of emails to ensure uniqueness across both sheets
    sheet_results = {}

    # Process each sheet
    for sheet_name in sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        # Check if the 'Student Name' column exists
        if 'Student Name' not in df.columns:
            print(f"The sheet '{sheet_name}' does not have a 'Student Name' column.")
            continue

        df_with_emails = process_sheet(sheet_name, df, existing_emails)
        sheet_results[sheet_name] = df_with_emails

    # Save the result for both sheets into a new Excel file
    output_file = 'student_emails_output.xlsx'
    with pd.ExcelWriter(output_file) as writer:
        for sheet_name, df in sheet_results.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"Email addresses generated and saved to {output_file}.")


if __name__ == '__main__':
    # Provide the path to your Excel file here
    file_path = 'testfiles.xlsx'
    generate_emails_from_excel(file_path)
