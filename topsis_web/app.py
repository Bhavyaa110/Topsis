from flask import Flask, render_template, request
import pandas as pd
import numpy as np
import os
import re
import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_ENV_PATH = os.path.join(os.path.dirname(BASE_DIR), ".env")
if os.path.exists(ROOT_ENV_PATH):
    load_dotenv(ROOT_ENV_PATH)
else:
    load_dotenv()

app = Flask(__name__)

def valid_email(email):
    return re.match(r"[^@]+@[^@]+\.[^@]+", email)

@app.route('/')
def index():
    return render_template("index.html")

@app.route('/submit', methods=['POST'])
def submit():
    file = request.files['file']
    raw_weights = request.form['weights']       
    raw_impacts = request.form['impacts']
    email = request.form.get('email', '')

    if email and not valid_email(email):
        return render_template("index.html", message="Invalid email format")

    weights = raw_weights.split(',')
    impacts = raw_impacts.split(',')

    if len(weights) != len(impacts):
        return render_template("index.html", message="Weights and impacts count mismatch")

    if not all(i in ['+', '-'] for i in impacts):
        return render_template("index.html", message="Impacts must be + or -")

    if file.filename.endswith(".csv"):
        data = pd.read_csv(file)
    elif file.filename.endswith((".xls", ".xlsx")):
        data = pd.read_excel(file)
    else:
        return render_template("index.html", message="Only CSV or Excel allowed")

    if data.shape[1] < 3:
        return render_template("index.html", message="File must contain at least 3 columns")

    criteria = data.iloc[:, 1:].astype(float)
    weights_arr = np.array(weights, dtype=float)

    norm = criteria / np.sqrt((criteria ** 2).sum())
    weighted = norm * weights_arr

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

    csv_data = data.to_csv(index=False)
    html_table = data.to_html(index=False, border=1)
    html_table = html_table.replace(
        '<table border="1" class="dataframe">',
        '<table border="1" cellspacing="0" cellpadding="6" '
        'style="border-collapse:collapse;width:100%;text-align:center;">'
    )

    send_email(email, csv_data, raw_weights, raw_impacts, html_table)

    return render_template(
        "index.html",
        table=data.to_html(classes="table",index=False),
        message="Result sent to email and displayed below"
    )

def send_email(to_email, csv_data, weights, impacts, html_table):
    EMAIL_USER = os.getenv("EMAIL_USER")
    EMAIL_PASS = os.getenv("EMAIL_PASS")

    msg = EmailMessage()
    msg['Subject'] = "TOPSIS Analysis Results"
    msg['From'] = EMAIL_USER
    msg['To'] = to_email

    msg.set_content(
        f"Your TOPSIS analysis results.\n\n"
        f"Analysis Parameters:\n"
        f"- Weights: {weights}\n"
        f"- Impacts: {impacts}\n\n"
        f"Please find the result CSV file attached.\n"
    )

    html_body = f"""
    <html>
      <body style="font-family: Arial, sans-serif; background-color:#f5f5f5; padding:20px;">
        <div style="max-width:700px;margin:0 auto;background:#ffffff;border-radius:8px;padding:24px;">
          <h2 style="margin-top:0;color:#202124;">Your TOPSIS Analysis Results</h2>
          <div style="background:#f1f3f4;border-radius:6px;padding:16px;margin-bottom:20px;">
            <h3 style="margin-top:0;margin-bottom:8px;font-size:16px;color:#202124;">Analysis Parameters:</h3>
            <ul style="margin:0;padding-left:20px;color:#202124;">
              <li><strong>Weights:</strong> {weights}</li>
              <li><strong>Impacts:</strong> {impacts}</li>
            </ul>
          </div>
          <p style="color:#202124;">Please find the result CSV file attached.</p>
          <h3 style="margin-top:0;font-size:16px;color:#202124;">Results Table:</h3>
          <div style="overflow-x:auto;border:1px solid #e0e0e0;border-radius:6px;background:#fafafa;padding:8px;">
            {html_table}
          </div>
          <p style="margin-top:24px;color:#5f6368;">Regards,<br><strong>TOPSIS Web Service</strong></p>
        </div>
      </body>
    </html>
    """

    msg.add_alternative(html_body, subtype='html')

    msg.add_attachment(
        csv_data.encode('utf-8'),
        maintype='text',
        subtype='csv',
        filename="topsis_result.csv"
    )

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(msg)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
