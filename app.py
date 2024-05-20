from flask import Flask, render_template, request
from secret_vars import CLIENT_ID, CLIENT_SECRET, AUTHORITY, RESOURCE, SITE_ID, DOCUMENT_LIBRARY_ID
from content_generation import get_graph_access_token, DownloadFile, checkAuthorisation, checkUnrequestedDocuments, saveAudit, WordEditingCode
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
            # Plain_Email_Draft('UnauthorisedEmail')
            os.remove('downloadedFiles/' + 'Licensee.xlsx')
            return "The user is not authorised."

        documents_found, documents_to_search = checkUnrequestedDocuments(documents_to_search, user_id)

        saveAudit(user_id, documents_to_search, user_email, user_name)

        for document in documents_to_search:
            DownloadFile(content_names[document] + ".docx", access_token, SITE_ID, DOCUMENT_LIBRARY_ID, RESOURCE)

        for file_name in os.listdir('downloadedFiles'):
            if file_name.endswith('.docx'):
                WordEditingCode(user_id, file_name)
                os.remove('downloadedFiles/' + file_name)




        return "Process completed successfully!"

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
