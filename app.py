from flask import Flask, render_template, request
from secret_vars import CLIENT_ID, CLIENT_SECRET, AUTHORITY, RESOURCE, SITE_ID, DOCUMENT_LIBRARY_ID
from content_generation import get_graph_access_token, DownloadFile
app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        user_id = request.form['userID']
        user_email = request.form['userEmail']
        user_name = request.form['userName']
        documents_to_search = request.form.getlist('documents_to_search')

       
        access_token = get_graph_access_token(CLIENT_ID, CLIENT_SECRET, AUTHORITY)


        if not access_token:
            return "Failed to acquire access token."

        DownloadFile('Licensee.xlsx', access_token, SITE_ID, DOCUMENT_LIBRARY_ID, RESOURCE)






        return "Process completed successfully!"

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
