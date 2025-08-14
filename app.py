from flask import Flask, render_template, request, redirect, url_for
import os
from openpyxl import Workbook, load_workbook
import resend

app = Flask(__name__)

EXCEL_FILE = "registrations.xlsx"

# Create Excel file if not exists
if not os.path.exists(EXCEL_FILE):
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Email", "Phone", "Institution", "Panel Preference"])
    wb.save(EXCEL_FILE)

# Resend API setup
resend.api_key = os.getenv("RESEND_API_KEY")
EMAIL_SENDER = os.getenv("EMAIL_SENDER")
EMAIL_RECEIVER = os.getenv("EMAIL_RECEIVER")

def send_email(to_email, subject, body_html):
    """Send email using Resend"""
    try:
        resend.Emails.send({
            "from": EMAIL_SENDER,
            "to": [to_email],
            "subject": subject,
            "html": body_html
        })
        return True
    except Exception as e:
        print(f"Email sending failed: {e}")
        return False

# -------------------- Routes --------------------

@app.route("/")
def home():
    return render_template("home.html")

@app.route("/register", methods=["GET", "POST"])
def register():
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

        # Email to participant
        send_email(
            email,
            "Accounting Conclave 2025 - Registration Confirmation",
            f"<p>Hello {name},</p>"
            "<p>Thank you for registering for Accounting Conclave 2025.</p>"
            "<p><b>Date:</b> 30th August 2025<br><b>Venue:</b> Ahmedabad University</p>"
            "<p>We look forward to seeing you!</p>"
            "<p>Best Regards,<br>Organizing Team</p>"
        )

        # Email to admin
        send_email(
            EMAIL_RECEIVER,
            "New Registration - Accounting Conclave 2025",
            f"<p><b>New registration received:</b></p>"
            f"<p>Name: {name}<br>Email: {email}<br>Phone: {phone}<br>"
            f"Institution: {institution}<br>Panel: {panel}</p>"
        )

        return redirect(url_for("success"))

    return render_template("register.html")

@app.route("/pricing")
def pricing():
    return render_template("pricing.html")

@app.route("/about")
def about():
    return render_template("about.html")

@app.route("/privacy")
def privacy():
    return render_template("privacy.html")

@app.route("/terms")
def terms():
    return render_template("terms.html")

@app.route("/contact")
def contact():
    return render_template("contact.html")

@app.route("/success")
def success():
    return render_template("success.html")

if __name__ == "__main__":
    app.run(debug=True)
