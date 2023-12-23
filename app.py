from flask import Flask, render_template, request, send_file, redirect, url_for
from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm
from io import BytesIO
import pyrebase
import re

app = Flask(__name__)

#----------------------Firebase Configuration----------------------#
config = {
  "apiKey": "AIzaSyDLK1VyXj-NcsQoMXV3xnu6UoInZt3W6cE",
  "authDomain": "formlogin-jf.firebaseapp.com",
  "databaseURL": "https://formlogin-jf-default-rtdb.asia-southeast1.firebasedatabase.app",
  "storageBucket": "formlogin-jf.appspot.com"
}

#----------------------Firebase Initialization----------------------#
firebase = pyrebase.initialize_app(config)
auth = firebase.auth()
db = firebase.database()

#----------------------person who's logged in----------------------#
person = {"is_logged_in": False, "name": "", "email": "", "uid": ""}

#Signin Page
@app.route('/')
def signin():
    return render_template('signin.html')

#Signup Page
@app.route('/signup')
def signup():
    return render_template('signup.html')

#Word formatter
@app.route('/home')
def home():
    if person['is_logged_in']:
        return render_template('index.html')
    else:
        return redirect(url_for('signin'))

#Authentication
@app.route('/result', methods=['POST', 'GET'])
def result():
    if request.method == "POST":
        result = request.form
        email = result["email"]
        password = result["pass"]
        try: 
            #try signing in with the credentials
            user = auth.sign_in_with_email_and_password(email, password)
            global person
            person['is_logged_in'] = True
            person['email'] = user['email']
            person['uid'] = user['localId']

            data = db.child("users").get()
            person['name'] = data.val()[person['uid']]['name']

            #redirect to the index page
            return redirect(url_for("home"))
        except:
            #if any errors
            return redirect(url_for('signin'))
    else:
        if person['is_logged_in']:
            return render_template('home')
        else:
            return redirect(url_for('signin'))

#Adding new users
@app.route('/register', methods = ['POST', 'GET'])
def register():
    if request.method == 'POST':
        result = request.form
        register_code = "4346"
        if result['regcode'] != register_code:
            return redirect(url_for('signup'))
        else:
            email = result['email']
            password = result['pass']
            name = result['name']
            try:
                auth.create_user_with_email_and_password(email, password)
                user = auth.sign_in_with_email_and_password(email, password)
                global person
                person['is_logged_in'] = True
                person['email'] = user['email']
                person['uid'] = user['localId']
                person['name'] = name
                data = {'name': name, 'email': email}
                db.child("users").child(person['uid']).set(data)

                return redirect(url_for('home'))
            except Exception as e:
                return redirect(url_for('signup'))
    else:
        if person['is_logged_in']:
            return render_template('index.html')
        else:
            return redirect(url_for('register'))        

@app.route('/generate_paper', methods=['POST'])
def generate_paper():
    # ----------------------Modify for Rich Text Fields----------------------#
    def match_pattern(s):
        pattern = re.compile(r'(\^\^\w+|__\w+)')
        matches = re.split(pattern, s)
        return [match for match in matches if match.strip()]

    def modify_text(text):
        if isinstance(text, str):
            text = [text]
        result = []
        for line in text:
            filtered_line = match_pattern(line)
            rt = RichText()
            for i in filtered_line:
                if i.startswith('^^'):
                    word = i.replace('^^', '')
                    rt.add(word, superscript=True)
                elif i.startswith('__'):
                    word = i.replace('__', '')
                    rt.add(word, subscript=True)
                else:
                    rt.add(i)
            result.append(rt)
        return result

    template = request.form.get('journal-type')
    doc = DocxTemplate(f"./word-template/{template}.docx")
    vol_inp = request.form.get('volume')
    month_inp = request.form.get('month')
    issue_inp = request.form.get('issuedate')
    issn_inp = request.form.get('issndate')
    doi_inp = request.form.get('doi')
    title_inp = request.form.get('title')
    address_inp = request.form.get('address').split('\n')
    author_inp = request.form.get('authors')
    subdate_inp = request.form.get('sub_date')
    accdate_inp = request.form.get('acc_date')
    revdate_inp = request.form.get('rev_date')
    abstracts_inp = request.form.get('abstract')
    keywords_inp = request.form.get("keyword")
    page_inp = request.form.get("page_no")
    ref = request.form.get('reference').split('\n')

    sections_inp = []
    added_sections = int(request.form.get('sectionIndex'))
    if added_sections > 0:
        i = 1
        while i <= added_sections:
            section_title = request.form.get(f'section_title_{i}')
            section_content = request.form.get(f'section_content_{i}')

            # Create a new section dictionary which will have table or image path afterwards
            section_data = [{
                'title': modify_text(section_title)[0],
                'text': modify_text(section_content)[0],
            }]

            added_fields = request.form.get(f'field_index_{i}')
            if added_fields and int(added_fields) > 0:
                elements = request.form.get(f'element_type_{i}').split(',')
                for j in range(0, int(added_fields) + 1):
                    if elements[j] == 'table':
                        table_lbl = request.form.get(f'section_table_label_{i}_{j}')
                        table_inp = request.form.get(f'section_table_{i}_{j}')
                        rows = table_inp.strip().split('\n') if table_inp else []
                        raw_data = [{'cols': row.strip().split('\t')} for row in rows]
                        element_dict = {'table': raw_data, 'table_lbl': table_lbl}
                        section_data.append(element_dict)
                    elif elements[j] == 'image':
                        section_image = request.files.get(f'section_image_{i}_{j}')
                        image_path = f"uploads/section_image_{i}.jpg"
                        section_image.save(image_path)
                        image_lbl = request.form.get(f'section_image_label_{i}_{j}')
                        element_dict = {'image_center': InlineImage(doc, image_path, width=Mm(100)), 'image_lbl': image_lbl}
                        section_data.append(element_dict)
            # Append section data to the list
            sections_inp.append(section_data)
            i += 1
    sections_inp.append([])

    context = {
        "vol": vol_inp,
        "issue": issue_inp,
        "address": modify_text(address_inp),
        "month": month_inp,
        "pp": page_inp,
        "issn": issn_inp,
        "doi": doi_inp,
        "title": modify_text(title_inp)[0],
        "authors": modify_text(author_inp)[0],
        "sub_date": subdate_inp,
        "acc_date": accdate_inp,
        "rev_date": revdate_inp,
        "abstract": modify_text(abstracts_inp)[0],
        "keywords": modify_text(keywords_inp)[0],
        'sections': sections_inp,
        'references': modify_text(ref),
    }

    doc.render(context)

    output_stream = BytesIO()
    doc.save(output_stream)
    output_stream.seek(0)
    return send_file(
        output_stream,
        as_attachment=True,
        download_name='formatter_output.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )

if __name__ == '__main__':
    app.run(debug=True)
