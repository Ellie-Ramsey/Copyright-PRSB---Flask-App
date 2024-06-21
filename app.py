from flask import Flask, render_template, request, jsonify
from main_functions import get_graph_access_token, email_content_generation, get_folder_id_by_name, move_email_to_folder, manual_content_generation
from secret_vars import CLIENT_ID, CLIENT_SECRET, AUTHORITY
import requests, re, html, time, threading

app = Flask(__name__)
access_token = get_graph_access_token(CLIENT_ID, CLIENT_SECRET, AUTHORITY)
user_email = "ellie@ramseysystems.onmicrosoft.com"
headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

# Define state variables for controlling the automation
automation_paused = threading.Event()
automation_running = threading.Event()

def get_emails_with_subject_pattern(subject_pattern):
    url = f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/inbox/messages"
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        emails = response.json().get('value', [])
        filtered_emails = [email for email in emails if re.match(subject_pattern, email['subject'])]
        return filtered_emails
    else:
        print(f"Failed to retrieve emails: {response.status_code}")
        print(response.json())
        return []

def parse_email_body(email_body):
    email_body = html.unescape(email_body)

    user_id_pattern = r'user_ID\s*=\s*"([^"]+)"'
    user_name_pattern = r'user_name\s*=\s*"([^"]+)"'
    documents_pattern = r'documents_to_search\s*=\s*\[([^\]]+)\]'

    user_id_match = re.search(user_id_pattern, email_body)
    user_name_match = re.search(user_name_pattern, email_body)
    documents_match = re.search(documents_pattern, email_body)

    user_id = user_id_match.group(1) if user_id_match else None
    user_name = user_name_match.group(1) if user_name_match else None
    documents = [doc.strip().strip('"') for doc in documents_match.group(1).split(',')] if documents_match else []

    return user_id, user_name, documents

def process_content_request_email(email):
    email_id = email['id']
    email_body = email['body']['content']
    user_id, user_name, documents_to_search = parse_email_body(email_body)
    email_of_sender = email['from']['emailAddress']['address']
    
    print(f"Email ID: {email_id}")
    print(f"User ID: {user_id}")
    print(f"User Name: {user_name}")
    print(f"Documents to Search: {documents_to_search}")
    print(f"Email of Sender: {email_of_sender}")

    email_content_generation(access_token, user_id, email_of_sender, user_name, documents_to_search, email_id)

def process_requested_content_email(email):
    email_id = email['id']
    subject = email['subject']
    match = re.search(r'Requested Content - (.+)', subject)
    if match:
        folder_name = match.group(1).strip()
        folder_id = get_folder_id_by_name(access_token, folder_name)
        if folder_id:
            move_email_to_folder(access_token, email_id, folder_id)
        else:
            print(f"Folder '{folder_name}' not found.")

def handle_content_requests():
    # Check if automation is paused
    if automation_paused.is_set():
        print("Automation is paused. Skipping content request handling.")
        return

    # Set automation running state
    automation_running.set()

    content_request_pattern = r'Content Request'
    requested_content_pattern = r'Requested Content - .+'
    
    content_request_emails = get_emails_with_subject_pattern(content_request_pattern)
    requested_content_emails = get_emails_with_subject_pattern(requested_content_pattern)
    
    for email in content_request_emails:
        if automation_paused.is_set():
            print("Automation paused during processing. Exiting...")
            break
        process_content_request_email(email)
    
    for email in requested_content_emails:
        if automation_paused.is_set():
            print("Automation paused during processing. Exiting...")
            break
        process_requested_content_email(email)

    # Clear automation running state
    automation_running.clear()

# Route to handle email notifications
@app.route('/notifications', methods=['POST'])
def notifications():
    if not automation_paused.is_set():
        handle_content_requests()
    else:
        print("Automation is paused. Notification handling is skipped.")
    return jsonify({'message': 'Notification received'}), 202

# Route to handle the manual form submission
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        user_id = request.form['userID']
        user_email = request.form['userEmail']
        user_name = request.form['userName']
        documents_to_search = request.form.getlist('documents_to_search')

        value = manual_content_generation(access_token, user_id, user_email, user_name, documents_to_search)

        # Start a background thread to crawl the inbox after a delay
        def delayed_crawl():
            time.sleep(30)
            if not automation_paused.is_set():
                handle_content_requests()
            else:
                print("Automation is paused. Delayed crawl is skipped.")
        
        threading.Thread(target=delayed_crawl).start()

        return jsonify({'message': 'Submission successful - please check the inbox for your drafted email.', 'value': value})

    return render_template('index.html')

# Route to pause the automation
@app.route('/pause_automation', methods=['POST'])
def pause_automation():
    if automation_running.is_set():
        message = "Automation is currently running. It will be paused as soon as possible."
    else:
        automation_paused.set()
        message = "Automation paused."

    return jsonify({'message': message}), 200

# Route to resume the automation
@app.route('/resume_automation', methods=['POST'])
def resume_automation():
    automation_paused.clear()
    return jsonify({'message': 'Automation resumed'}), 200

# Periodically crawl the inbox every 15 minutes
def periodic_crawl():
    while True:
        if not automation_paused.is_set():
            handle_content_requests()
        time.sleep(900)  # Sleep for 900 seconds (15 minutes)

if __name__ == '__main__':
    # Start the background thread for periodic crawling
    threading.Thread(target=periodic_crawl, daemon=True).start()
    app.run(host='0.0.0.0', port=8080)