from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
import boto3
import json
import os

# retrieve sharepoint password from AWS Secret Manager
def retrieve_password_secret_manager(secretID):
    client = boto3.client('secretsmanager', region_name='us-east-1')
    response = client.get_secret_value(
        SecretID = secretID
    )
    sharepoint_secrets = json.loads(response['SecretString'])
    password = sharepoint_secrets['password']
    return password

def upload_file_sharepoint(url, username, password, folder_url, local_file_path):
    print("Authenticating with SharePoint")
    #Connect to SharePoint
    ctx_auth = AuthenticationContext(url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(url, ctx_auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
        print("Authentication Successful")
        print("Web Title: {0}".format(web.properties['Title']))
    
        #Upload the file
        with open(local_file_path, "rb") as content_file:
            file_content = content_file.read()
        file_path = os.path.basename(local_file_path)
        file = ctx.web.get_folder_by_server_relative_url(folder_url).upload_file(file_path, file_content).execute_query()
        print("[OK] file has been uploaded to url: {0}".format(file.serverRelativeUrl))
            

if __name__ == "__main__":
    
    sharepoint_url = os.environ["SHAREPOINT_URL"] #SharePoint URL from where you want to upload the file.
    username = os.environ["SHAREPOINT_USERNAME"] #SharePoint Username for authentication
    secretID = os.environ["SECRET_ID"] #SecretID for retrieving the SharePoint password from AWS Secret Manager
    password = retrieve_password_secret_manager(secretID)
    folder_url = os.environ["FOLDER_URL"] #Folder relative path in SharePoint where you want to upload the file.
    local_file_path = os.environ["LOCAL_FILE_PATH"] #Local File Path that needs to be uploaded to SharePoint.
    
    
    upload_file_sharepoint(sharepoint_url, username, password, folder_url, local_file_path)
    