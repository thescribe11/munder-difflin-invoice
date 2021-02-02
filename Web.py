from main import generate_invoice, send_email
from flask import *

app = Flask(__name__)
app.debug = True
app.secret_key = 'panzerfaust'

CURRENT_DATA = {'name': '',
                'email': '',
                'address': '',
                'transactions': []}


def format_price(raw) -> str:
    if raw[0] == '$':
        raw = raw[1:len(raw)]
    return '$%.2f' % (float(raw), )


@app.route('/add/', methods=['POST', 'GET'])
def add_data():
    if request.method == 'GET':
        return render_template('add.html', data=CURRENT_DATA)
    elif request.method == 'POST':
        if 'Add' in request.form:
            CURRENT_DATA['transactions'].append([
                request.form.get('Description'),
                format_price(request.form.get('Price')),
                int(request.form.get('Amount')),
            ])
            return render_template('add.html', data=CURRENT_DATA)
        elif 'Generate' in request.form:
            print(CURRENT_DATA)
            generate_invoice(CURRENT_DATA['name'], CURRENT_DATA['email'], CURRENT_DATA['address'],
                             CURRENT_DATA['transactions'])
            send_email(CURRENT_DATA['email'], CURRENT_DATA['name'])
            flash("This is a test.")
            return redirect(url_for('main'))


@app.route('/', methods=['POST', 'GET'])
def main():
    if request.method == 'GET':
        return render_template('input.html', data=CURRENT_DATA)
    elif request.method == 'POST':
        CURRENT_DATA['name'] = request.form.get('Name')
        CURRENT_DATA['email'] = request.form.get('Email')
        CURRENT_DATA['address'] = request.form.get('Address')
        return redirect(url_for('add_data'))


if __name__ == "__main__":
    app.run()
