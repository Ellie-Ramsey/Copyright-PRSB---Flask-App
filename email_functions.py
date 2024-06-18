import os, json, requests, base64

# EMAIL DRAFT FUNCTION: Creates a draft email (no attachment version, need to remove to use the attachment version)
def Plain_Email_Draft(email_type, access_token, userEmail):
    # Read JSON data from file
    with open('data/email_content.json', 'r') as file:
        email_contents = json.load(file)
    
    # Extract email subject and message based on email_type
    if email_type in email_contents:
        subject = email_contents[email_type]['title']
        content = email_contents[email_type]['message']
        content = content.replace("\\n", "<br>")
    else:
        print("Email type not found.")
        return

    # Set up parameters for the draft email
    user_email = "ellie@ramseysystems.onmicrosoft.com" 
    create_draft_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
    headers = {
        "Authorization": "Bearer " + access_token,
        "Content-Type": "application/json"
    }
    mail_json = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": content
        },
        "toRecipients": [{"emailAddress": {"address": userEmail}}],
        "ccRecipients": [
            {"emailAddress": {"address": "ellie@ramseysystems.onmicrosoft.com"}}
        ]
    }
    
    # Send request to create draft email
    response = requests.post(create_draft_url, headers=headers, json=mail_json)
    if response.status_code == 201:
        print("Draft email created successfully!")
    else:
        print(f"Failed to create draft email: {response.text}")

# RETURNS SIZE OF FILE IN MB
def get_file_size(file_path):
    """Returns the size of the file in megabytes."""
    return os.path.getsize(file_path) / 1024 / 1024

def prepare_attachments(files):
    """Prepare attachment groups based on size limit."""
    max_size = 21  # Maximum size in MB for attachments
    groups = []
    current_group = []
    current_size = 0
    
    for file in files:
        file_size = get_file_size(file)
        if current_size + file_size > max_size:
            groups.append(current_group)
            current_group = []
            current_size = 0
        current_group.append(file)
        current_size += file_size
    if current_group:
        groups.append(current_group)
    
    return groups

def encode_attachment(file_path):
    """Encodes a file as a base64 string for attachment."""
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

# not a draft currently a send
def Attachment_Email_Draft(email_type, files, access_token, userEmail, userID):
    with open('data/email_content.json', 'r') as file:
        email_contents = json.load(file)

    if email_type in email_contents:
        subject = email_contents[email_type]['title'] + f" - {userID}"
        content = email_contents[email_type]['message']
        content = content.replace("\\n", "<br>")
    else:
        print("Email type not found.")
        return

    attachment_groups = prepare_attachments(files) if files else [[]]

    for attachments in attachment_groups:
        user_email = "ellie@ramseysystems.onmicrosoft.com"
        send_mail_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail"
        headers = {
            "Authorization": "Bearer " + access_token,
            "Content-Type": "application/json"
        }
        
        attachment_data = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": os.path.basename(attachment),
            "contentBytes": encode_attachment(attachment),
            "contentType": "application/octet-stream"
        } for attachment in attachments] if attachments else []

        mail_json = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": content
                },
                "toRecipients": [{"emailAddress": {"address": userEmail}}],
                "ccRecipients": [
                    {"emailAddress": {"address": "ellie@ramseysystems.onmicrosoft.com"}}
                ],
                "attachments": attachment_data
            },
            "saveToSentItems": "true"
        }

        response = requests.post(send_mail_url, headers=headers, json=mail_json)
        
        # Debugging
        print(f"Response status code: {response.status_code}")
        print(f"Response content: {response.content}")

        if response.status_code == 202:
            print("Email sent successfully!")
        else:
            print(f"Failed to send email: {response.text}")

# EMAIL MOVE FUNCTIONS        
def get_all_folders(access_token, user_email):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    list_folders_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders"
    
    folders = []
    while list_folders_url:
        response = requests.get(list_folders_url, headers=headers)
        
        if response.status_code == 200:
            response_json = response.json()
            folders.extend(response_json.get('value', []))
            list_folders_url = response_json.get('@odata.nextLink')
        else:
            print(f"Failed to retrieve folders: {response.status_code}")
            print(response.json())
            return []
    
    return folders

def find_folder_by_name(folders, folder_name):
    for folder in folders:
        print(f"Checking folder: {folder['displayName']} (ID: {folder['id']})")
        if folder['displayName'].lower() == folder_name.lower():
            print(f"Match found: {folder['displayName']} (ID: {folder['id']})")
            return folder['id']
    return None

def get_folder_id_by_name(access_token, folder_name):
    user_email = "ellie@ramseysystems.onmicrosoft.com"
    headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
    }
    
    all_folders = get_all_folders(access_token, user_email)
    if not all_folders:
        return None

    folder_id = find_folder_by_name(all_folders, folder_name)
    if folder_id:
        return folder_id

    # Explore nested folders
    for folder in all_folders:
        child_folders_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/{folder['id']}/childFolders"
        while child_folders_url:
            response = requests.get(child_folders_url, headers=headers)
            if response.status_code == 200:
                child_folders = response.json().get('value', [])
                folder_id = find_folder_by_name(child_folders, folder_name)
                if folder_id:
                    return folder_id
                child_folders_url = response.json().get('@odata.nextLink')
            else:
                print(f"Failed to retrieve child folders: {response.status_code}")
                print(response.json())
                break

    return None

def move_email_to_folder(access_token, email_id, folder_id):
    user_email = "ellie@ramseysystems.onmicrosoft.com"
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    move_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages/{email_id}/move"
    move_payload = {
        "destinationId": folder_id
    }
    
    print(f"Moving email with ID {email_id} to folder ID {folder_id}")
    
    response = requests.post(move_url, headers=headers, json=move_payload)
    
    if response.status_code in [200, 201]:
        print("Email moved successfully.")
    else:
        print(f"Failed to move email: {response.status_code}")
        print(response.json())