from flask import Flask, render_template, request
from main_functions import content_generation

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        user_id = request.form['userID']
        user_email = request.form['userEmail']
        user_name = request.form['userName']
        documents_to_search = request.form.getlist('documents_to_search')
       
        value = content_generation(user_id, user_email, user_name, documents_to_search)

        return value

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
