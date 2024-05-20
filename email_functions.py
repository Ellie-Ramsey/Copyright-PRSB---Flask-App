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

# PREPARE ATTACHMENTS FUNCTION: Prepares attachment groups based on size limit
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

# ENCODE ATTACHMENT FUNCTION: Encodes a file as a base64 string for attachment
def encode_attachment(file_path):
    """Encodes a file as a base64 string for attachment."""
    with open(file_path, "rb") as f:
        return base64.b64encode(f.read()).decode()

# DRAFT EMAIL FUNCTION: Creates a draft email with attachments
def Attachment_Email_Draft(email_type, files, access_token, userEmail):
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

    # Prepare attachment groups
    attachment_groups = prepare_attachments(files) if files else [[]]

    for attachments in attachment_groups:
        # Set up parameters for the draft email
        user_email = "ellie@ramseysystems.onmicrosoft.com"
        create_draft_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages"
        headers = {
            "Authorization": "Bearer " + access_token,
            "Content-Type": "application/json"
        }
        
        # Prepare attachment data
        attachment_data = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": os.path.basename(attachment),
            "contentBytes": encode_attachment(attachment),
            "contentType": "application/octet-stream"
        } for attachment in attachments] if attachments else []

        mail_json = {
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
        }

        # Send request to create draft email
        response = requests.post(create_draft_url, headers=headers, json=mail_json)
        if response.status_code == 201:
            print("Draft email created successfully!")
        else:
            print(f"Failed to create draft email: {response.text}")
