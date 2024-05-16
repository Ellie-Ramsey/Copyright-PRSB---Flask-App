from flask import Flask, render_template, request
from secret_vars import CLIENT_ID, CLIENT_SECRET, AUTHORITY, RESOURCE, SITE_ID, DOCUMENT_LIBRARY_ID
from content_generation import get_graph_access_token, DownloadFile, checkAuthorisation
import os

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # user_id = request.form['userID']
        # user_email = request.form['userEmail']
        # user_name = request.form['userName']
        # documents_to_search = request.form.getlist('documents_to_search')

        user_name = "Steven Stevenson"
        user_email = "ellie@ramseysystems.co.uk"
        user_id = "SP-The Steven Group-0062"
        documents_to_search = ["Adult Carer Quality of Life (AC-QoL)", "Clinical Frailty Scale (CFS)", "Recovering Quality of Life Questionnaire (ReQoL)"]

       
        access_token = get_graph_access_token(CLIENT_ID, CLIENT_SECRET, AUTHORITY)


        if not access_token:
            return "Failed to acquire access token."

        DownloadFile('Licensee.xlsx', access_token, SITE_ID, DOCUMENT_LIBRARY_ID, RESOURCE)


        if not checkAuthorisation(user_email, user_id):
            # Plain_Email_Draft('UnauthorisedEmail')
            os.remove('downloadedFiles/' + 'Licensee.xlsx')
            return "The user is not authorised."



        return "Process completed successfully!"

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
