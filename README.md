# upload_file_sharepoint_python
This  template contains a python script that helps to upload a file to a SharePoint location from your local server.


## Pre-Requisites
- Python with Office365-REST-Python-Client package installed
- AWS CLI Installed
- Export the following variables on the terminal using command `export sharepoint_url = "https://companyname.sharepoint.com/sites/site-name"`:

- SHAREPOINT_URL - SharePoint URL from where you want to download the file.
- SHAREPOINT_USERNAME - SharePoint Username for authentication
- SECRET_ID - SecretID for retrieving the SharePoint password from AWS Secret Manager
- FOLDER_URL - Folder relative path in SharePoint where you want to upload the file
- LOCAL_PATH - Local File Path that needs to be uploaded to SharePoint

## How to run the script

- Clone the Repo
- Run the script using the following command `python3 upload_file_sharepoint.py`