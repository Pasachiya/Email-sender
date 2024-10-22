import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os
from dotenv import load_dotenv
import qrcode
from io import BytesIO
import base64

# Load environment variables
load_dotenv()

# Email configuration
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

def load_excel_data(file_path):
    """Load data from Excel file."""
    return pd.read_excel(file_path)

def validate_payment(row):
    """Validate payment (placeholder function)."""
    return row['payment_verified'] == 'Yes'

def generate_qr_code(data):
    """Generate QR code and return as base64 encoded string."""
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buffered = BytesIO()
    img.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()

def generate_ticket_html(row):
    """Generate HTML ticket from template."""
    with open('ticket_template.html', 'r') as file:
        template = file.read()
    
    # Generate QR code
    qr_code = generate_qr_code(row['ticket_number'])
    
    # Replace placeholders with actual data
    ticket_html = template.replace('{NAME}', row['name'])
    ticket_html = ticket_html.replace('{BATCH}', str(row['batch']))
    ticket_html = ticket_html.replace('{TICKET_NUMBER}', str(row['ticket_number']))
    ticket_html = ticket_html.replace('{EVENT_DATE}', '2024-10-18')  # Replace with actual event date
    ticket_html = ticket_html.replace('{EVENT_TIME}', '7:00 PM')  # Replace with actual event time
    ticket_html = ticket_html.replace('{VENUE}', 'University Auditorium')  # Replace with actual venue
    ticket_html = ticket_html.replace('{ADDITIONAL_INFO}', 'Please bring your student ID')  # Replace with actual additional info
    ticket_html = ticket_html.replace('{CONTACT_EMAIL}', 'organizer@example.com')  # Replace with actual contact email
    ticket_html = ticket_html.replace('{QR_CODE}', f'data:image/png;base64,{qr_code}')
    
    # Replace social media URLs
    ticket_html = ticket_html.replace('{FACEBOOK_URL}', 'https://facebook.com/your_event_page')
    ticket_html = ticket_html.replace('{TWITTER_URL}', 'https://twitter.com/your_event_page')
    ticket_html = ticket_html.replace('{INSTAGRAM_URL}', 'https://instagram.com/your_event_page')
    
    return ticket_html

def send_email(to_email, subject, html_content):
    """Send HTML email."""
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(html_content, 'html'))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)

def main():
    # Load Excel data
    df = load_excel_data('responses.xlsx')
    
    for _, row in df.iterrows():
        if validate_payment(row):
            ticket_html = generate_ticket_html(row)
            send_email(row['email'], "Your IT Batches Get-Together Ticket", ticket_html)
            print(f"Sent ticket to {row['name']} ({row['email']})")
        else:
            print(f"Payment not verified for {row['name']} ({row['email']})")

if __name__ == "__main__":
    main()