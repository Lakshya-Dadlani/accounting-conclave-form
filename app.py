from flask import Flask, render_template, request, redirect, url_for
import os
from openpyxl import Workbook, load_workbook
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = Flask(__name__)

EXCEL_FILE = "registrations.xlsx"

# Create Excel file if not exists
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Email", "Phone", "Institution", "Panel Preference"])
    wb.save(EXCEL_FILE)

# Email settings (use your own Gmail/SMTP credentials)
EMAIL_USER = "your_email@gmail.com"
EMAIL_PASS = "your_password"  # App password for Gmail
EMAIL_RECEIVER = "your_email@gmail.com"  # Your email for admin notifications

def send_email(to_email, subject, body):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_USER
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL_USER, EMAIL_PASS)
        server.sendmail(EMAIL_USER, to_email, msg.as_string())

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        name = request.form["name"]
        email = request.form["email"]
        phone = request.form["phone"]
        institution = request.form["institution"]
        panel = request.form["panel"]

        # Save to Excel
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
        ws.append([name, email, phone, institution, panel])
        wb.save(EXCEL_FILE)

        # Email to user
        send_email(email, "Accounting Conclave 2025 - Registration Confirmation",
                   f"Hello {name},\n\nThank you for registering for Accounting Conclave 2025.\nDate: 30th August 2025\nVenue: Ahmedabad University\n\nWe look forward to seeing you!\n\nBest Regards,\nOrganizing Team")

        # Email to admin
        send_email(EMAIL_RECEIVER, "New Registration - Accounting Conclave 2025",
                   f"New registration received:\n\nName: {name}\nEmail: {email}\nPhone: {phone}\nInstitution: {institution}\nPanel: {panel}")

        return redirect(url_for("success"))

    return render_template("index.html")

@app.route("/success")
def success():
    return render_template("success.html")

if __name__ == "__main__":
    app.run(debug=True)
