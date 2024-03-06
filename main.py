# Python code to send email to a list of emails from a spreadsheet

# import the required libraries
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# change these as per use
your_email = "YOUR_EMAIL"
your_password = "YOUR_PASSWORD"  # Note: Replace with an App Password or environmental variable.

# Try establishing a connection with Gmail
try:
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(your_email, your_password)
except smtplib.SMTPException as e:
    print(f"Error: Unable to login. {e}")
    server.close()
    exit()

# Try reading the spreadsheet
try:
    email_list = pd.read_excel('input/YOUR_FILE.xlsx')
except FileNotFoundError as e:
    print(f"Error: The file was not found. {e}")
    server.close()
    exit()
except Exception as e:
    print(f"An error occurred while reading the Excel file. {e}")
    server.close()
    exit()

# Get the names and the emails
names = email_list['Name']
emails = email_list['Email']

# Iterate through the records
for i in range(len(emails)):
    # For every record, get the name and the email addresses
    name = names[i]
    email = emails[i]

    # Compose the email message
    message = MIMEText("Hello " + name)
    message['Subject'] = 'YOUR_SUBJECT'
    message['From'] = your_email
    message['To'] = email

    # Try sending the email
    try:
        server.send_message(message)
        print(f"Email sent to {email}")
    except smtplib.SMTPException as e:
        print(f"Error: Unable to send email to {email}. {e}")

# Finally, close the SMTP server connection
server.close()