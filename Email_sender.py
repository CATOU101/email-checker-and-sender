import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import decode_header
import pandas as pd
import re
import dns.resolver

# Add email validation using regex
def is_valid_email(email):
    """Check if the email syntax is correct."""
    email_regex = r'^[a-zA-Z0-9_.&%-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$' #Improve regex for better email validation
    return re.match(email_regex, email) is not None

# Add MX record check for email domains
def has_mx_record(domain):
    """Check if the domain has an MX record."""
    try:
        records = dns.resolver.resolve(domain, 'MX')
        return bool(records)
    except (dns.resolver.NoAnswer, dns.resolver.NXDOMAIN):
        return False
    except (dns.exception.Timeout, dns.resolver.NoNameservers, dns.exception.DNSException):
        return False

# Add function to send emails via SMTP
def send_email(smtp_server, smtp_port, smtp_user, smtp_password, to_email, subject, body): #Add logging to track email sending status
    """Send an email using SMTP."""
    try:
        message = MIMEMultipart()
        message['From'] = smtp_user
        message['To'] = to_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'html'))  # Use 'html' for HTML content
        
        with smtplib.SMTP(smtp_server, smtp_port) as session:
            session.starttls()
            session.login(smtp_user, smtp_password)
            session.sendmail(smtp_user, to_email, message.as_string())
        
        return True
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
        return False

# Add function to check for error replies
def check_reply_emails(imap_server, email_address, password):
    """Check for error replies using IMAP."""
    try:
        with imaplib.IMAP4_SSL(imap_server) as imap_conn:
            imap_conn.login(email_address, password)
            imap_conn.select('inbox')

            status, response = imap_conn.search(None, 'UNSEEN')
            if status != 'OK':
                print("Failed to search inbox.")
                return []

            email_ids = response[0].split()
            error_emails = []

            for email_id in email_ids:
                status, email_data = imap_conn.fetch(email_id, '(RFC822)')
                if status != 'OK':
                    print(f"Failed to fetch email with ID: {email_id}")
                    continue

                msg = email.message_from_bytes(email_data[0][1])
                from_email = msg['From']
                subject, encoding = decode_header(msg['Subject'])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else 'utf-8')

                error_mail_address = 'mailer-daemon@googlemail.com'
                if error_mail_address in from_email:
                    body = get_email_body(msg)
                    error_keywords = ['550', '5.1.1', 'Delivery Status Notification (Failure)', '554']
                    error_emails.append((msg, subject, body))

            return error_emails
    except Exception as e:
        print(f"Error while checking reply emails: {e}")
        return []

# Add function to extract email body
def get_email_body(msg):
    """Extract the body of the email."""
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            if "attachment" not in content_disposition:
                if content_type == "text/plain" or content_type == "text/html":
                    payload = part.get_payload(decode=True)
                    if payload is not None:
                        body = payload.decode()
                        break
    else:
        payload = msg.get_payload(decode=True)
        if payload is not None:
            body = payload.decode()
    return body

def main():
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_user = 'Your@email.com'  # Replace with your email
    smtp_password = 'XXXXXXXX'  # Replace with your App password
    imap_server = 'imap.gmail.com'

    # Add functionality to read email addresses from Excel
    excel_file = 'email.xlsx'  # Replace with your excel file path
    subject = "Your Subject Here"

    try:
        df = pd.read_excel(excel_file)
        email_list = df['Email'].tolist()
    except Exception as e:
        print("Failed to read email IDs from the Excel file:", e)
        return

    valid_emails = []
    invalid_emails = []
    error_details = []

    # Check validity and MX record of each email
    for email_addr in email_list:
        if is_valid_email(email_addr): 
            domain = email_addr.split('@')[1]
            if has_mx_record(domain):
                valid_emails.append(email_addr)
            else:
                invalid_emails.append(email_addr)
        else:
            invalid_emails.append(email_addr)

    print(f"Total valid emails: {len(valid_emails)}")
    print(f"Total invalid emails: {len(invalid_emails)}")

    # Add logging for successful and unsuccessful email sends
    successful_emails = 0 
    unsuccessful_emails = 0 
    
    # Loop through the list of valid emails and send emails
    for email_addr in valid_emails:
        # Find the corresponding row in the DataFrame
        row = df[df['Email'] == email_addr].iloc[0]
        recipient_name = row['Name']
        
        # Personalized email body with recipient's name and Html etc
        body = f"""\
        <html>
            <body>
                <p>Hi {recipient_name},</p>
                <p>This is a personalized message for you.</p>
                <p>Here is your table:</p>
                <p>Regards,<br>Your Name</p>
            </body>
        </html>"""

        if send_email(smtp_server, smtp_port, smtp_user, smtp_password, email_addr, subject, body):
            print(f"Email sent successfully to {email_addr}")
            successful_emails += 1
        else:
            print(f"Failed to send email to {email_addr}")
            error_details.append((email_addr, "Failed to send email"))
            unsuccessful_emails += 1 # Track unsuccessful email attempts and log reasons

    # Check for error replies
    error_emails = check_reply_emails(imap_server, smtp_user, smtp_password)
    for msg, subject, body in error_emails:
        for email_addr in valid_emails:
            if email_addr in subject or email_addr in body:
                error_details.append((email_addr, subject))
                successful_emails -= 1
                unsuccessful_emails += 1 # Track unsuccessful email attempts and log reasons

    print(f"Successful emails after error check: {successful_emails}")
    print(f"Unsuccessful emails after error check: {unsuccessful_emails}")

    # Log error details to Excel for tracking failed emails
    for email_addr in invalid_emails:
        error_details.append((email_addr, "Invalid email format"))

    error_details_df = pd.DataFrame(error_details, columns=["Email", "Reason"])
    error_details_file = 'email_error_details.xlsx'  # Replace with your file path to store error mails
    error_details_df.to_excel(error_details_file, index=False)

if __name__ == "__main__":
    main()
