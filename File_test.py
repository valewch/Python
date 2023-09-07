from shareplum import Site
from shareplum.site import Version
from requests_ntlm import HttpNtlmAuth
import getpass
import pandas as pd
import boto3


# username and password
username = getpass.getuser()
password = None  # password is not needed as Windows credentials will be used

# Specify the document library or folder and file name
site_url = "https://share.amazon.com/sites/sjocopsqa/"
auth = HttpNtlmAuth(username, password)
site = Site(site_url, version=Version.v2016, auth=auth, verify_ssl=False)
folder = site.Folder('Shared Documents/Recalls/Recalls Compiled Files')  # Replace with the actual library or folder name
file_name = "MW Compiled File.xlsx"  # Replace with the name of the file you want to download

# Download the file and save it to your local directory
file_contents=folder.get_file(file_name)

read_file = pd.read_excel(file_contents, sheet_name=['MW Data', 'Raw Data Errors & Root Cause'])
#read_file.to_excel('File_test.xlsx', index=False)

with pd.ExcelWriter('File_test.xlsx', engine='xlsxwriter') as writer:
    # Iterate over the dictionary items and save each DataFrame to a separate sheet
    for sheet_name, df in read_file.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("File saved successfully.")

#file = site.DocumentLibrary(document_library_name).get_file(file_name)
#file.download_file(file_name)
#pd = pd.DataFrame(file_contents)
#pd.to_csv(r'D:\Users\valesch\Documents\Python tests', index=False)



