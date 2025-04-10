import pandas as pd
import os
import time
import logging
import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Constants
EXCEL_FILE = "employees.xlsx"
TOKEN_FILE = "token.json"
CREDENTIALS_FILE = "credentials.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def load_excel(file_path):
    """Load Excel file and clean column names."""
    if not os.path.exists(file_path):
        logging.error(f"File not found: {file_path}")
        return None
    try:
        data = pd.read_excel(file_path)
        data.columns = [col.strip() for col in data.columns]
        return data
    except Exception as e:
        logging.error(f"Failed to load Excel file: {e}")
        return None

def verify_columns(data, required_columns):
    """Verify required columns exist in the DataFrame."""
    missing_columns = [col for col in required_columns if col not in data.columns]
    if missing_columns:
        logging.error(f"Missing column(s): {missing_columns}")
        return False
    return True

def calculate_net_salary(data):
    """Calculate Net Salary for employees."""
    try:
        data["Net Salary"] = data["Gross Salary"] - (
            data["Gross Salary"] * data["Tax (%)"] / 100) - data["Deductions"]
        return data
    except Exception as e:
        logging.error(f"Error calculating net salary: {e}")
        return None

def authenticate_gmail():
    """Authenticate with Gmail API."""
    creds = None
    try:
        if os.path.exists(TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
            with open(TOKEN_FILE, "w") as token:
                token.write(creds.to_json())
    except Exception as e:
        logging.error(f"Authentication error: {e}")
    return creds

def create_payslip(employee):
    """Generate a PDF payslip for an employee."""
    pdf_file = f"{employee['Name']}_Payslip.pdf"
    try:
        pdf = canvas.Canvas(pdf_file, pagesize=letter)
        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawString(200, 750, "Employee Payslip")
        pdf.line(50, 740, 550, 740)

        # Add employee details
        pdf.setFont("Helvetica", 12)
        pdf.drawString(50, 710, f"Employee Name: {employee['Name']}")
        pdf.drawString(50, 690, f"Gross Salary: ${employee['Gross Salary']:.2f}")
        pdf.drawString(50, 670, f"Tax Percentage: {employee['Tax (%)']}%")
        pdf.drawString(50, 650, f"Deductions: ${employee['Deductions']:.2f}")
        pdf.drawString(50, 630, f"Net Salary: ${employee['Net Salary']:.2f}")

        pdf.setFont("Helvetica-Oblique", 10)
        pdf.drawString(50, 50, "Thank you for your hard work and dedication!")

        pdf.save()
    except Exception as e:
        logging.error(f"Error creating payslip for {employee['Name']}: {e}")
        return None
    return pdf_file

def send_email(creds, sender_email, recipient_email, subject, body, attachment_path):
    """Send an email with an attachment."""
    try:
        service = build("gmail", "v1", credentials=creds)
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = recipient_email
        message["Subject"] = subject
        message.attach(MIMEText(body, "plain"))

        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as attachment:
                part = MIMEApplication(attachment.read())
                part.add_header("Content-Disposition", "attachment", filename=os.path.basename(attachment_path))
                message.attach(part)
        else:
            logging.error(f"Attachment not found: {attachment_path}")
            return False

        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
        service.users().messages().send(userId="me", body={"raw": raw_message}).execute()
        logging.info(f"Email sent to {recipient_email}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email to {recipient_email}: {e}")
        return False

def main():
    """Main function to execute the workflow."""
    data = load_excel(EXCEL_FILE)
    if data is None or not verify_columns(data, ['Gross Salary', 'Tax (%)', 'Deductions', 'Name', 'Email']):
        return

    data = calculate_net_salary(data)
    if data is None:
        return

    creds = authenticate_gmail()
    sender_email = os.getenv("EMAIL_USER")
    if not creds or not sender_email:
        logging.error("Gmail authentication failed or sender email is missing.")
        return

    for _, employee in data.iterrows():
        payslip_file = create_payslip(employee)
        if payslip_file:
            subject = "Your Payslip"
            body = f"Hello {employee['Name']},\nPlease find your payslip attached."
            if not send_email(creds, sender_email, employee["Email"], subject, body, payslip_file):
                logging.error(f"Failed to send email to {employee['Email']}")
        time.sleep(2)  # To avoid API rate limits

if __name__ == "__main__":
    main()
