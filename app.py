from flask import Flask, render_template, request
from secret_vars import CLIENT_ID, CLIENT_SECRET, AUTHORITY, RESOURCE, SITE_ID, DOCUMENT_LIBRARY_ID
from content_generation import get_graph_access_token, DownloadFile, checkAuthorisation, checkUnrequestedDocuments, saveAudit, WordEditingCode, upload_file_to_sharepoint
from email_functions import Plain_Email_Draft, Attachment_Email_Draft
import os, json

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        user_id = request.form['userID']
        user_email = request.form['userEmail']
        user_name = request.form['userName']
        documents_to_search = request.form.getlist('documents_to_search')
       
        access_token = get_graph_access_token(CLIENT_ID, CLIENT_SECRET, AUTHORITY)
        with open('data/content_names.json', 'r') as file:
             content_names = json.load(file)


        if not access_token:
            return "Failed to acquire access token."

        DownloadFile('Licensee.xlsx', access_token, SITE_ID, DOCUMENT_LIBRARY_ID, RESOURCE)


        if not checkAuthorisation(user_email, user_id):
            Plain_Email_Draft('UnauthorisedEmail', access_token, user_email)
            os.remove('downloadedFiles/' + 'Licensee.xlsx')
            return "The user is not authorised."

        documents_found, documents_to_search = checkUnrequestedDocuments(documents_to_search, user_id)

        Nigel_Required = False
        for document in documents_to_search:
            if document == "Strengths and Difficulties Questionnaire (SDQ)":
                Nigel_Required = True
                documents_to_search.remove(document)
                print("Main - IMPORTANT: Nigel Specific Document Required")
                print("Main - Updated douments NOT Previously Requested: ", documents_to_search)
                break

        saveAudit(user_id, documents_to_search, user_email, user_name)

        for document in documents_to_search:
            DownloadFile(content_names[document] + ".docx", access_token, SITE_ID, DOCUMENT_LIBRARY_ID, RESOURCE)

        for file_name in os.listdir('downloadedFiles'):
            if file_name.endswith('.docx'):
                WordEditingCode(user_id, file_name)
                os.remove('downloadedFiles/' + file_name)


        for file_name in os.listdir('downloadedFiles'):
            if file_name.endswith('.docx'):
                
                nhs_index = file_name.find('NHS')
                if nhs_index != -1:
                    original_file_name = file_name[nhs_index:-5]

                    for key, value in content_names.items():
                        if value == original_file_name:
                            content_name = key
                            break

                upload_result = upload_file_to_sharepoint(access_token, "Permissions/Completed/" + content_name + "/Sub-Licenses Granted", "downloadedFiles/" + file_name, SITE_ID, DOCUMENT_LIBRARY_ID)


        if documents_to_search:
            print("Main - Audit was uploaded")
            upload_result = upload_file_to_sharepoint(access_token, "Licensees" , "downloadedFiles/Licensee.xlsx", SITE_ID, DOCUMENT_LIBRARY_ID)
        else:
            print("Main - Audit was not uploaded")
        os.remove('downloadedFiles/Licensee.xlsx')

        documents_found = [user_id + content_names[book] for book in documents_found]

        special = False
        for document in documents_found:
            print("Downloading: ", document)
            DownloadFile(document + ".docx", access_token, SITE_ID, DOCUMENT_LIBRARY_ID, RESOURCE)

            if document == "NHS England Sub-Licence - Content Specific Terms A_ReQoL_10" or "NHS England Sub-Licence - Content Specific Terms A_ReQoL_20" and special == False:
                special = True
                DownloadFile("ReQoL - Important Notes.pdf", access_token, SITE_ID, DOCUMENT_LIBRARY_ID, RESOURCE)



        files = [os.path.join('downloadedFiles', file_name) for file_name in os.listdir('downloadedFiles')]

        Attachment_Email_Draft('RequestedContent', files, access_token, user_email)

        for file_name in os.listdir('downloadedFiles'):
            file_path = os.path.join('downloadedFiles', file_name)
            if os.path.isfile(file_path):
                os.remove(file_path)


        return "Process completed successfully!"

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
