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


