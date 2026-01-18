from flask import Flask, render_template, request
import pandas as pd
import numpy as np
import os
import re
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/submit', methods=['POST'])
def submit():
    file = request.files['file']
    weights = request.form['weights']
    impacts = request.form['impacts']
    email = request.form['email']

    # Email validation
    if not valid_email(email):
        return render_template("index.html", message="Invalid email format")

    weights = weights.split(',')
    impacts = impacts.split(',')

    if len(weights) != len(impacts):
        return render_template("index.html", message="Weights and impacts count mismatch")

    if not all(i in ['+', '-'] for i in impacts):
        return render_template("index.html", message="Impacts must be + or -")

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)

    # Read CSV or Excel
    try:
        data = pd.read_csv(input_path)
    except:
        data = pd.read_excel(input_path)

    if data.shape[1] < 3:
        return render_template("index.html", message="File must contain at least 3 columns")

    criteria = data.iloc[:, 1:].astype(float)
    weights = np.array(weights, dtype=float)

    # TOPSIS
    norm = criteria / np.sqrt((criteria ** 2).sum())
    weighted = norm * weights

    ideal_best, ideal_worst = [], []

    for i in range(len(impacts)):
        if impacts[i] == '+':
            ideal_best.append(weighted.iloc[:, i].max())
            ideal_worst.append(weighted.iloc[:, i].min())
        else:
            ideal_best.append(weighted.iloc[:, i].min())
            ideal_worst.append(weighted.iloc[:, i].max())

    ideal_best = np.array(ideal_best)
    ideal_worst = np.array(ideal_worst)

    dist_best = np.sqrt(((weighted - ideal_best) ** 2).sum(axis=1))
    dist_worst = np.sqrt(((weighted - ideal_worst) ** 2).sum(axis=1))

    score = dist_worst / (dist_best + dist_worst)

    data['Topsis Score'] = score
    data['Rank'] = data['Topsis Score'].rank(ascending=False)

    result_path = os.path.join(RESULT_FOLDER, "result.csv")
    data.to_csv(result_path, index=False)

    send_email(email, result_path)

    return render_template(
        "index.html",
        table=data.to_html(index=False),
        message="Result sent to email and displayed below"
    )

def send_email(to_email, attachment_path):
    EMAIL_USER = os.getenv("EMAIL_USER")
    EMAIL_PASS = os.getenv("EMAIL_PASS")
    msg = EmailMessage()
    msg['Subject'] = "TOPSIS Result"
    msg['From'] = "bhavyaagarwal00000@gmail.com"
    msg['To'] = to_email
    msg.set_content("Please find the TOPSIS result attached.")

    with open(attachment_path, 'rb') as f:
        msg.add_attachment(f.read(), maintype='text', subtype='csv', filename="result.csv" or "result.xlsx")

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(msg)

if __name__ == '__main__':
    app.run(debug=True)
