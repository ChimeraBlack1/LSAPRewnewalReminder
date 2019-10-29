"""The first step is to create an SMTP object, each object is used for connection 
with one server."""

import smtplib
server = smtplib.SMTP('documentdirection-ca.mail.protection.outlook.com', 25)

# Import the email modules we'll need
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

fromaddr = "chris@documentdirection.ca"
toaddr = "chris@documentdirection.ca"
msg = MIMEMultipart()
msg['From'] = "chris@documentdirection.ca"
msg['To'] = "chris@documentdirection.ca"
msg['Subject'] = "Expiring Software Licenses"

body = "body line 1 <br>"
msg.attach(MIMEText(body, 'html'))
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)

#Send the mail
# msg = "Helloooo!"
# msg = """From:Chris Parke <chris@documentdirection.ca>
# To: <chris@documentdirection.ca>
# MIME-Version: 1.0
# Content-type: text/html
# Subject: Expiring Software Licenses


# """
# msg = msg + """

# This is an e-mail message to be sent in HTML format \n

# """
# msg = msg + """

# bacon
# """


# server.sendmail("chris@documentdirection.ca", "chris@documentdirection.ca", msg)
print("mail sent...")
# print(msg)