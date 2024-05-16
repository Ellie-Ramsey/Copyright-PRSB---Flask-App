import os, requests, shutil, json, base64, urllib.parse, warnings
from openpyxl import load_workbook
from datetime import datetime
from docx import Document

#  MICROSOFT API ACCESS TOKEN CODE: This function acquires an access token from Microsoft Graph API using the client credentials flow.
def get_graph_access_token(CLIENT_ID, CLIENT_SECRET,AUTHORITY):
    data = {
      "grant_type": "client_credentials",
      "client_id": CLIENT_ID,
      "scope": "https://graph.microsoft.com/.default",
      "client_secret": CLIENT_SECRET
    }
  
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    token_url = f"{AUTHORITY}/oauth2/v2.0/token"
    response = requests.post(token_url, headers=headers, data=data)
  
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        print(f"Error acquiring Graph access token: {response.text}")
        return None


# FILE DOWNLOAD: master function to download a file from SharePoint
def DownloadFile(file_name, access_token, SITE_ID, DOCUMENT_LIBRARY_ID, RESOURCE):
    document_id = search_for_document(access_token, file_name, SITE_ID, DOCUMENT_LIBRARY_ID)
    
    if document_id:
        fetch_sharepoint_document_by_item_id(file_name, access_token, document_id, RESOURCE, SITE_ID, DOCUMENT_LIBRARY_ID)

#  SEARCH FOR DOCUMENT BY NAME: returns the document ID if found, or None if not found.
def search_for_document(access_token, file_name, SITE_ID, DOCUMENT_LIBRARY_ID):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # URL encode the file name to handle spaces and special characters
    encoded_file_name = urllib.parse.quote(file_name)
    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DOCUMENT_LIBRARY_ID}/root/search(q='{encoded_file_name}')"
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        items = response.json().get('value', [])
        for item in items:
            if item['name'] == file_name:
                return item['id']
    else:
        return None

# ACCESS SHAREPOINT CODE: This function fetches a document from SharePoint using the item ID.
def fetch_sharepoint_document_by_item_id(file_name, access_token, item_id, RESOURCE, SITE_ID, DOCUMENT_LIBRARY_ID):
    folder_path = 'downloadedFiles'

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    file_path = os.path.join(folder_path, file_name)

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    url = f"{RESOURCE}/sites/{SITE_ID}/drives/{DOCUMENT_LIBRARY_ID}/items/{item_id}/content"
    response = requests.get(url, headers=headers, allow_redirects=True)

    if response.status_code == 200:
        with open(file_path, 'wb') as file:
            file.write(response.content)
    else:
        print(f"Error fetching SharePoint document by item ID: {response.text}")
