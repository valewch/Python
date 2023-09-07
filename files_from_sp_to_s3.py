from shareplum import Site
from shareplum.site import Version
from requests_ntlm import HttpNtlmAuth
import getpass
import pandas as pd
import boto3

# Modify all the Upper Cased strings.
s3 = boto3.resource(
    service_name='s3',
    region_name='REGION',
    aws_access_key_id='YOUR ACCESS KEY',
    aws_secret_access_key='YOUR SECRET ACCESS KEY'
)

# username and password
username = getpass.getuser()
password = None  # password is not needed as Windows credentials will be used

# Specify the document library or folder and file name
site_url = "https://share.COMPANY.com/sites/SITE_NAME/"
auth = HttpNtlmAuth(username, password)
site = Site(site_url, version=Version.v2016, auth=auth, verify_ssl=False)
folder = site.Folder('Shared Documents/FOLDER NAME')  # Replace with the actual library or folder name

files = folder.files
for file in files:
    file_name = file['Name'] # Replace with the name of all the files in the folder
    file_contents=folder.get_file(file_name) 
    read_file = pd.read_excel(file_contents, None)
    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
    # Iterate over the dictionary items and save each DataFrame to a separate sheet
        for sheet_name, df in read_file.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)


file_path = "C:/Users/YOUR_USER/FOLDER_PATH/"
for file in files:
    file_name = file['Name']
    file_name_path = file_path + file_name
    s3_folder = "S3 FOLDER PATH" 
    s3_file_name = s3_folder + file_name
    s3.Bucket('quicksigth-dashboards').upload_file(Filename = file_name_path, Key = s3_file_name)


print("Files uploaded successfully.")


