from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import os

# Authenticate and create the PyDrive client
gauth = GoogleAuth()
gauth.LocalWebserverAuth()  # Creates local webserver and automatically handles authentication

drive = GoogleDrive(gauth)

# List of files to upload
files_to_upload = ['similarity_results.json', 'male_female.log', 'student_emails.log','merged_data.json','import_excell_file.py','main.py','male_female.py','save_to_google_drive.py','similar_names.py']

# Upload files to Google Drive
for file_name in files_to_upload:
    file_path = os.path.join(os.getcwd(), file_name)
    if os.path.exists(file_path):
        gfile = drive.CreateFile({'title': file_name})
        gfile.SetContentFile(file_path)
        gfile.Upload()
        print(f"Uploaded {file_name} to Google Drive")
    else:
        print(f"File {file_name} does not exist")

print("All files uploaded to Google Drive")