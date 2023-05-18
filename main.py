import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Read email details from Excel sheet
data = pd.read_excel('email_details.xlsx')
email_list = data['Email']
name_list = data['Name']
subject_list = data['Subject']

# Read email body from a text file
with open('email_body.txt', 'r') as file:
    common_body = file.read()

# Gmail account details
sender_email = 'youremail@gmail.com'
password = 'yourpassword'

# SMTP server configuration
smtp_server = 'smtp.gmail.com'
smtp_port = 587

# Iterate through each email in the list and send the email
for i in range(len(email_list)):
    # Create a multipart message
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = email_list[i]
    message['Subject'] = subject_list[i]

    # Replace placeholders in the common body with recipient's name
    personalized_body = common_body.format(name=name_list[i])

    # Add the body of the email
    message.attach(MIMEText(personalized_body, 'plain'))

    # Attach the file
    attachment_file = 'resume.pdf'
    with open(attachment_file, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={attachment_file}')
    message.attach(part)

    # Connect to the SMTP server and send the email
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, password)
    server.send_message(message)
    server.quit()

    print(f"Email sent to {name_list[i]} ({email_list[i]})")
